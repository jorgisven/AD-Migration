<#
.SYNOPSIS
    Validates that accounts are mapped to OUs that exist in the OU Map.
.DESCRIPTION
    Checks the User and Computer account mapping files to ensure the 'TargetOU_DN' for each
    account corresponds to a valid 'TargetDN' in the 'OU_Map_Draft.csv'.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$MapPath = Join-Path $config.TransformRoot 'Mapping'

$ouMapFile = Join-Path $MapPath "OU_Map_Draft.csv"
$userMapFile = Join-Path $MapPath "User_Account_Map.csv"
$computerMapFile = Join-Path $MapPath "Computer_Account_Map.csv"

Write-Log -Message "Starting validation of account placement..." -Level INFO
$hasErrors = $false

# 1. Load the valid Target OUs from the OU Map
if (-not (Test-Path $ouMapFile)) {
    throw "OU Map file not found at '$ouMapFile'. Cannot validate account placement."
}
$validOUs = Import-Csv $ouMapFile | Where-Object { $_.Action -ne 'Skip' } | Select-Object -ExpandProperty TargetDN
$validOUSet = [System.Collections.Generic.HashSet[string]]::new($validOUs, [System.StringComparer]::OrdinalIgnoreCase)
# Also add the domain root as a valid placement target
$domainDN = ($validOUs | Select-Object -First 1) -replace '.*?(DC=.*)', '$1'
if ($domainDN) { $validOUSet.Add($domainDN) | Out-Null }

Write-Host "[+] Loaded $($validOUSet.Count) valid OU destinations from OU Map." -ForegroundColor Green

# 2. Validate User Account Placements
if (Test-Path $userMapFile) {
    $userMap = Import-Csv $userMapFile
    $invalidUserPlacements = @()
    foreach ($user in $userMap) {
        if ($user.Action -eq 'Create' -and -not $validOUSet.Contains($user.TargetOU_DN)) {
            $invalidUserPlacements += $user
        }
    }
    if ($invalidUserPlacements.Count -gt 0) {
        Write-Host "[-] ERROR: Found $($invalidUserPlacements.Count) users mapped to an invalid or skipped OU." -ForegroundColor Red
        $invalidUserPlacements | ForEach-Object { Write-Host "  - User: $($_.TargetSam) -> Invalid OU: $($_.TargetOU_DN)" -ForegroundColor Red }
        $hasErrors = $true
    } else {
        Write-Host "[+] PASSED: All user account placements are valid." -ForegroundColor Green
    }
} else {
    Write-Host "[!] WARNING: User account map not found at '$userMapFile'. Skipping validation." -ForegroundColor Yellow
}

# 3. Validate Computer Account Placements
if (Test-Path $computerMapFile) {
    $computerMap = Import-Csv $computerMapFile
    $invalidComputerPlacements = @()
    foreach ($computer in $computerMap) {
        if ($computer.Action -eq 'Create' -and -not $validOUSet.Contains($computer.TargetOU_DN)) {
            $invalidComputerPlacements += $computer
        }
    }
    if ($invalidComputerPlacements.Count -gt 0) {
        Write-Host "[-] ERROR: Found $($invalidComputerPlacements.Count) computers mapped to an invalid or skipped OU." -ForegroundColor Red
        $invalidComputerPlacements | ForEach-Object { Write-Host "  - Computer: $($_.TargetName) -> Invalid OU: $($_.TargetOU_DN)" -ForegroundColor Red }
        $hasErrors = $true
    } else {
        Write-Host "[+] PASSED: All computer account placements are valid." -ForegroundColor Green
    }
} else {
    Write-Host "[!] WARNING: Computer account map not found at '$computerMapFile'. Skipping validation." -ForegroundColor Yellow
}

Write-Host ""
if ($hasErrors) {
    Write-Host "=== Account Placement Validation FAILED ===" -ForegroundColor Red
    Write-Host "Please correct the 'TargetOU_DN' in the account mapping files and re-run." -ForegroundColor Red
} else {
    Write-Host "=== Account Placement Validation Complete ===" -ForegroundColor Green
    Write-Host "All account placements appear to be valid."
}