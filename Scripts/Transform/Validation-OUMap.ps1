<#
.SYNOPSIS
    Validates the OU_Map_Draft.csv file.
.DESCRIPTION
    Checks the user-edited OU mapping file for common errors like duplicate target DNs,
    invalid characters, and orphaned OUs before it is used in later transform steps.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic
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

# 0. Check for unmapped OUs
$unmapped = $OUMap | Where-Object { $_.Action -ne 'Skip' -and [string]::IsNullOrWhiteSpace($_.TargetDN) }
if ($unmapped) {
    Write-Host "[-] ERROR: Found $($unmapped.Count) unmapped OUs. TargetDN cannot be blank unless Action is 'Skip'." -ForegroundColor Red
    $unmapped | ForEach-Object { Write-Host "  - Source: $($_.SourceDN)" -ForegroundColor Red }
    $hasErrors = $true
}

# 0.5 Check for non-empty skipped OUs
$ouExportPath = Join-Path $config.ExportRoot 'OU_Structure'
$emptyListFile = Join-Path $ouExportPath "EmptyOUs.txt"
if (Test-Path $emptyListFile) {
    $emptyOUs = @(Get-Content $emptyListFile)
    $nonEmptySkipped = $OUMap | Where-Object { $_.Action -eq 'Skip' -and $_.SourceDN -notin $emptyOUs }
    
    if ($nonEmptySkipped) {
        Write-Host "[!] WARNING: Found $($nonEmptySkipped.Count) skipped OUs that are NOT EMPTY." -ForegroundColor Yellow
        Write-Host "    Any objects (Users, Computers, GPOs) inside these OUs will be orphaned or fail to migrate!" -ForegroundColor Yellow
        $nonEmptySkipped | ForEach-Object { Write-Host "  - Skipped (Not Empty): $($_.SourceDN)" -ForegroundColor Yellow }
    }
}

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

        # Ensure the TargetDN is actually an OU (prevents mapping directly to a DC or CN)
        if ($row.TargetDN -notmatch "^OU=") {
            Write-Host "[-] ERROR: TargetDN must start with 'OU='. Cannot map to a domain component or container: '$($row.TargetDN)'" -ForegroundColor Red
            $hasErrors = $true
        }

    # Check for invalid characters in the OU name part
        $ouName = ($row.TargetDN -split '(?<!\\),')[0] -replace '^OU=','' -replace '\\,',','
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

    $parentDN = $row.TargetDN -replace "^.*?(?<!\\),", ""
    
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

# 4. Optional Live Target Validation
$msg = "Do you want to connect to the live target domain to verify your mapping anchor points?`n`nThis ensures the parent paths you are mapping into actually exist in the target AD environment.`n(Requires network line-of-sight to the Target Domain)."
$onlineResult = [System.Windows.Forms.MessageBox]::Show($msg, "Live Target Validation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

if ($onlineResult -eq 'Yes') {
    $TargetDomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the FQDN of the TARGET domain (e.g., target.local):", "Live Validation", "")
    if (-not [string]::IsNullOrWhiteSpace($TargetDomain)) {
        Write-Host "`n--- Performing Live Target Validation against '$TargetDomain' ---" -ForegroundColor Cyan
        
        # Find all "Anchor" DNs. An anchor is a parent DN that is NOT being created by this script.
        $anchorDNs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $OUMap) {
            if ($row.Action -eq 'Skip' -or [string]::IsNullOrWhiteSpace($row.TargetDN)) { continue }
            $parentDN = $row.TargetDN -replace "^.*?(?<!\\),", ""
            if (-not $validDNs.Contains($parentDN)) {
                $anchorDNs.Add($parentDN) | Out-Null
            }
        }

        foreach ($anchor in $anchorDNs) {
            try {
                # Use Get-ADObject because the anchor could be an OU or the domain root (DC=...)
                $null = Get-ADObject -Identity $anchor -Server $TargetDomain -ErrorAction Stop
                Write-Host "[+] PASSED: Anchor path exists in target: $anchor" -ForegroundColor Green
            } catch {
                Write-Host "[-] ERROR: Anchor path does not exist in target domain: $anchor" -ForegroundColor Red
                Write-Host "    Details: $($_.Exception.Message)" -ForegroundColor Red
                $hasErrors = $true
            }
        }
    } else {
        Write-Host "[-] Live validation skipped (no domain provided)." -ForegroundColor Yellow
    }
}

Write-Host ""
if ($hasErrors) {
    Write-Host "=== OU Map Validation FAILED ===" -ForegroundColor Red
    Write-Host "Please correct the errors in '$mapFile' and re-run the validation." -ForegroundColor Red
    throw "Validation Failed"
} else {
    Write-Host "=== OU Map Validation Complete ===" -ForegroundColor Green
    Write-Host "OU Map appears to be valid."
}