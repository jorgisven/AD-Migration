<#
.SYNOPSIS
    Validates the OU_Map_Draft.csv file.
.DESCRIPTION
    Checks the user-edited OU mapping file for common errors like duplicate target DNs,
    invalid characters, and orphaned OUs before it is used in later transform steps.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$MapPath = Join-Path $config.TransformRoot 'Mapping'
$mapFile = Join-Path $MapPath "OU_Map_Draft.csv"

Write-Log -Message "Starting validation of OU Map..." -Level INFO

if (-not (Test-Path $mapFile)) {
    Write-Host "[-] ERROR: OU Map file not found at $mapFile" -ForegroundColor Red
    Write-Log -Message "OU Map validation failed: File not found." -Level ERROR
    throw "OU Map file not found."
}

$OUMap = Import-Csv $mapFile
$hasErrors = $false
$validDNs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

# 1. Check for duplicate TargetDNs
$duplicates = $OUMap | Where-Object { $_.Action -ne 'Skip' } | Group-Object TargetDN | Where-Object { $_.Count -gt 1 }
if ($duplicates) {
    Write-Host "[-] ERROR: Duplicate TargetDNs found in OU Map. Each TargetDN must be unique." -ForegroundColor Red
    $duplicates.Name | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    $hasErrors = $true
} else {
    Write-Host "[+] PASSED: No duplicate TargetDNs found." -ForegroundColor Green
}

# 2. Check for invalid characters and structure
foreach ($row in $OUMap) {
    if ($row.Action -eq 'Skip') { continue }

    # Add to valid DN set for parent check
    $validDNs.Add($row.TargetDN) | Out-Null

    # Check for invalid characters in the OU name part
    $ouName = ($row.TargetDN -split ",")[0] -replace "OU=",""
    if ($ouName -match '[\\/:*?"<>|]') {
        Write-Host "[-] ERROR: Invalid characters in OU name for TargetDN '$($row.TargetDN)'" -ForegroundColor Red
        $hasErrors = $true
    }
}
Write-Host "[+] PASSED: Basic character validation complete." -ForegroundColor Green

# 3. Check for orphaned OUs (parents must exist)
$orphans = @()
foreach ($row in $OUMap) {
    if ($row.Action -eq 'Skip') { continue }

    $parentDN = $row.TargetDN -replace "^[^,]+,", ""
    
    if ($parentDN -ne $row.TargetDN -and $parentDN -notmatch "^DC=") {
        if (-not $validDNs.Contains($parentDN)) {
            $orphans += $row
        }
    }
}

if ($orphans.Count -gt 0) {
    Write-Host "[-] ERROR: Found $($orphans.Count) orphaned OUs. Their parent DN does not exist in the map." -ForegroundColor Red
    $orphans | ForEach-Object { Write-Host "  - Orphan: $($_.TargetDN) | Missing Parent: $($_.TargetDN -replace '^[^,]+,', '')" -ForegroundColor Red }
    $hasErrors = $true
} else {
    Write-Host "[+] PASSED: All OUs have a valid parent in the map." -ForegroundColor Green
}

Write-Host ""
if ($hasErrors) {
    Write-Host "=== OU Map Validation FAILED ===" -ForegroundColor Red
    Write-Host "Please correct the errors in '$mapFile' and re-run the validation." -ForegroundColor Red
} else {
    Write-Host "=== OU Map Validation Complete ===" -ForegroundColor Green
    Write-Host "OU Map appears to be valid."
}