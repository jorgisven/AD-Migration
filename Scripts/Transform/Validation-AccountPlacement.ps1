<#
.SYNOPSIS
    Validates that accounts are mapped to OUs that exist in the OU Map.
.DESCRIPTION
    Checks the User and Computer account mapping files to ensure the 'TargetOU_DN' for each
    account corresponds to a valid 'TargetDN' in the 'OU_Map_Draft.csv'.
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

$ouMapFile = Join-Path $MapPath "OU_Map_Draft.csv"
$userMapFile = Join-Path $MapPath "User_Account_Map.csv"
$computerMapFile = Join-Path $MapPath "Computer_Account_Map.csv"

Write-Log -Message "Starting validation of account placement..." -Level INFO
$hasErrors = $false

# 1. Load the valid Target OUs from the OU Map
if (-not (Test-Path $ouMapFile)) {
    throw "OU Map file not found at '$ouMapFile'. Cannot validate account placement."
}
$ouMapData = @(Import-Csv $ouMapFile)
$validOUs = @($ouMapData | Where-Object { $_.Action -ne 'Skip' } | Select-Object -ExpandProperty TargetDN)
$validOUSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$validOUs, [System.StringComparer]::OrdinalIgnoreCase)
# Also add the domain root as a valid placement target
$domainDN = ($validOUs | Select-Object -First 1) -replace '.*?(DC=.*)', '$1'
if ($domainDN) { $validOUSet.Add($domainDN) | Out-Null }

Write-Host "[+] Loaded $($validOUSet.Count) valid OU destinations from OU Map." -ForegroundColor Green

# 2. Validate User Account Placements
if (Test-Path $userMapFile) {
    $userMap = @(Import-Csv $userMapFile)
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
    $computerMap = @(Import-Csv $computerMapFile)
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

# 4. Optional Live Target Validation
$msg = "Do you want to connect to the live target domain to verify your account destination OUs?`n`nThis ensures the OUs you are placing accounts into actually exist in the target AD environment.`n(Requires network line-of-sight to the Target Domain)."
$onlineResult = [System.Windows.Forms.MessageBox]::Show($msg, "Live Target Validation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

if ($onlineResult -eq 'Yes') {
    $TargetDomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the FQDN of the TARGET domain (e.g., target.local):", "Live Validation", "")
    if (-not [string]::IsNullOrWhiteSpace($TargetDomain)) {
        Write-Host "`n--- Performing Live Target Validation against '$TargetDomain' ---" -ForegroundColor Cyan
        
        $targetOUsToCheck = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        if (Test-Path $userMapFile) { foreach ($u in $userMap) { if ($u.Action -eq 'Create' -and $u.TargetOU_DN) { $targetOUsToCheck.Add($u.TargetOU_DN) | Out-Null } } }
        if (Test-Path $computerMapFile) { foreach ($c in $computerMap) { if ($c.Action -eq 'Create' -and $c.TargetOU_DN) { $targetOUsToCheck.Add($c.TargetOU_DN) | Out-Null } } }

        foreach ($ou in $targetOUsToCheck) {
            try {
                $null = Get-ADObject -Identity $ou -Server $TargetDomain -ErrorAction Stop
                Write-Host "[+] PASSED: Destination OU exists in target: $ou" -ForegroundColor Green
            } catch {
                # Check if this OU is planned for creation in the map
                $mapRow = $ouMapData | Where-Object { $_.TargetDN -eq $ou } | Select-Object -First 1
                
                if ($mapRow -and $mapRow.SourceDN -ne 'PRE-EXISTING') {
                    # It's normal for it not to exist yet, because Import-OUs hasn't run.
                    Write-Host "[~] PENDING: Destination OU does not exist yet (Will be created during Import): $ou" -ForegroundColor DarkGray
                } elseif ($ou -eq $domainDN) {
                    Write-Host "[-] ERROR: Domain root does not exist or is unreachable: $ou" -ForegroundColor Red
                    $hasErrors = $true
                } else {
                    Write-Host "[-] ERROR: Destination OU does not exist in target domain: $ou" -ForegroundColor Red
                    Write-Host "    Details: $($_.Exception.Message)" -ForegroundColor Red
                    $hasErrors = $true
                }
            }
        }

        Write-Host "`n--- Checking for Existing Accounts in Target Domain ---" -ForegroundColor Cyan
        $accountConflict = $false
        
        if (Test-Path $userMapFile) {
            foreach ($u in $userMap) {
                if ($u.Action -eq 'Create' -and $u.TargetSam) {
                    try {
                        $null = Get-ADUser -Identity $u.TargetSam -Server $TargetDomain -ErrorAction Stop
                        Write-Host "[-] ERROR: User account '$($u.TargetSam)' already exists in the target domain." -ForegroundColor Red
                        $accountConflict = $true
                        $hasErrors = $true
                    } catch { } # Account does not exist, which is what we want when Action is 'Create'
                } elseif ($u.Action -eq 'Merge' -and $u.TargetSam) {
                    try {
                        $null = Get-ADUser -Identity $u.TargetSam -Server $TargetDomain -ErrorAction Stop
                    } catch {
                        Write-Host "[-] ERROR: User account '$($u.TargetSam)' marked for 'Merge', but does NOT exist in the target domain." -ForegroundColor Red
                        $hasErrors = $true
                    }
                }
            }
        }

        if (Test-Path $computerMapFile) {
            foreach ($c in $computerMap) {
                if ($c.Action -eq 'Create' -and $c.TargetName) {
                    try {
                        $null = Get-ADComputer -Identity $c.TargetName -Server $TargetDomain -ErrorAction Stop
                        Write-Host "[-] ERROR: Computer account '$($c.TargetName)' already exists in the target domain." -ForegroundColor Red
                        $accountConflict = $true
                        $hasErrors = $true
                    } catch { } # Computer does not exist, which is what we want
                } elseif ($c.Action -eq 'Merge' -and $c.TargetName) {
                    try {
                        $null = Get-ADComputer -Identity $c.TargetName -Server $TargetDomain -ErrorAction Stop
                    } catch {
                        Write-Host "[-] ERROR: Computer account '$($c.TargetName)' marked for 'Merge', but does NOT exist in the target domain." -ForegroundColor Red
                        $hasErrors = $true
                    }
                }
            }
        }
        
        if (-not $accountConflict) {
            Write-Host "[+] PASSED: No target account name conflicts found." -ForegroundColor Green
        }
    } else {
        Write-Host "[-] Live validation skipped (no domain provided)." -ForegroundColor Yellow
    }
}

Write-Host ""
if ($hasErrors) {
    Write-Host "=== Account Placement Validation FAILED ===" -ForegroundColor Red
    Write-Host "Please correct the 'TargetOU_DN' in the account mapping files and re-run." -ForegroundColor Red
} else {
    Write-Host "=== Account Placement Validation Complete ===" -ForegroundColor Green
    Write-Host "All account placements appear to be valid."
}