<#
.SYNOPSIS
    Orchestrates the full import process into the target domain.
.DESCRIPTION
    Runs all individual import scripts in the correct sequence.
    This script MUST be run with Administrator privileges in the target domain.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain
)

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# --- Admin Check ---
Add-Type -AssemblyName System.Windows.Forms
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $msg = "This script must be run as an Administrator to create objects in Active Directory.`nPlease restart your PowerShell session with 'Run as Administrator'."
    [System.Windows.Forms.MessageBox]::Show($msg, "Administrator Privileges Required", "OK", "Error")
    Write-Host "FATAL: Administrator privileges required. Exiting." -ForegroundColor Red
    exit
}

# --- Module Load ---
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
if (-not (Test-Path $config.ImportRoot)) {
    New-Item -Path $config.ImportRoot -ItemType Directory -Force | Out-Null
}

$stateFile = Join-Path $config.ImportRoot 'Import-Progress-State.json'

function Save-ImportState {
    param(
        [hashtable]$State
    )

    $State.updatedAt = (Get-Date).ToString('o')
    $State | ConvertTo-Json -Depth 6 | Set-Content -Path $stateFile -Encoding UTF8
}

function New-ImportState {
    param(
        [string]$Domain,
        [string[]]$AllScripts
    )

    return @{
        version          = 1
        targetDomain     = $Domain
        scripts          = $AllScripts
        completedScripts = @()
        failedScripts    = @()
        lastScript       = ''
        status           = 'running'
        startedAt        = (Get-Date).ToString('o')
        updatedAt        = (Get-Date).ToString('o')
    }
}

# --- Get Target Domain ---
if (-not $TargetDomain) {
    try {
        $TargetDomain = (Get-ADDomain).DNSRoot
        $msg = "Detected target domain: '$TargetDomain'. Is this correct?"
        $result = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm Target Domain", "YesNo", "Question")
        if ($result -eq 'No') { $TargetDomain = '' }
    } catch {
        $TargetDomain = ''
    }
}
if (-not $TargetDomain) {
    Add-Type -AssemblyName Microsoft.VisualBasic
    $TargetDomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the FQDN of the TARGET domain:", "Import Phase", "target.local")
    if ([string]::IsNullOrWhiteSpace($TargetDomain)) {
        Write-Host "Import cancelled. Target domain cannot be empty." -ForegroundColor Red
        exit
    }
}

# --- Script Execution ---
$scripts = @(
    "Import-OUs.ps1",
    "Import-Accounts.ps1",
    "Import-WMIFilters.ps1",
    "Import-GPOs.ps1",
    "Import-GPOLinks.ps1",
    "Import-DNS.ps1"
)

Write-Host "=== Starting Full Import for Domain: $TargetDomain ===" -ForegroundColor Cyan
$confirmResult = [System.Windows.Forms.MessageBox]::Show("This will begin creating objects in the domain '$TargetDomain'.`n`nEnsure you have reviewed all mapping files in the 'Transform' directory before proceeding.", "Confirm Import", "OKCancel", "Warning")
if ($confirmResult -ne 'OK') {
    Write-Host "Import cancelled by user." -ForegroundColor Yellow
    exit
}

$startIndex = 0
$importState = $null
if (Test-Path $stateFile) {
    try {
        $existingState = Get-Content -Path $stateFile -Raw | ConvertFrom-Json -ErrorAction Stop
        if ($existingState -and $existingState.status -ne 'completed' -and $existingState.scripts) {
            $existingDomain = [string]$existingState.targetDomain
            $existingLast = [string]$existingState.lastScript
            $resumeMsg = "A previous import progress state was found.`n`nTarget Domain: $existingDomain`nLast Step: $existingLast`n`nYES = Resume from next incomplete step`nNO = Restart import from beginning"
            $resumeChoice = [System.Windows.Forms.MessageBox]::Show($resumeMsg, "Resume Previous Import?", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)

            if ($resumeChoice -eq [System.Windows.Forms.DialogResult]::Cancel) {
                Write-Host "Import cancelled by user." -ForegroundColor Yellow
                exit
            }

            if ($resumeChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
                if ($existingDomain -and $existingDomain -ne $TargetDomain) {
                    $domainMismatch = [System.Windows.Forms.MessageBox]::Show("Saved state domain '$existingDomain' differs from current '$TargetDomain'.`n`nYES = Use saved domain and resume`nNO = Keep current domain and restart", "Domain Mismatch", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    if ($domainMismatch -eq [System.Windows.Forms.DialogResult]::Yes) {
                        $TargetDomain = $existingDomain
                    } else {
                        Remove-Item -Path $stateFile -Force -ErrorAction SilentlyContinue
                    }
                }

                if (Test-Path $stateFile) {
                    $completed = @($existingState.completedScripts)
                    $startIndex = $completed.Count
                    if ($startIndex -gt $scripts.Count) { $startIndex = 0 }
                    $importState = @{
                        version          = 1
                        targetDomain     = $TargetDomain
                        scripts          = $scripts
                        completedScripts = $completed
                        failedScripts    = @($existingState.failedScripts)
                        lastScript       = [string]$existingState.lastScript
                        status           = 'running'
                        startedAt        = [string]$existingState.startedAt
                        updatedAt        = (Get-Date).ToString('o')
                    }
                    Save-ImportState -State $importState
                    Write-Host "Resuming import from step index $startIndex of $($scripts.Count)." -ForegroundColor Yellow
                }
            } else {
                Remove-Item -Path $stateFile -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {
        Write-Host "Warning: existing import state was unreadable and will be ignored." -ForegroundColor Yellow
        Remove-Item -Path $stateFile -Force -ErrorAction SilentlyContinue
    }
}

if ($null -eq $importState) {
    $importState = New-ImportState -Domain $TargetDomain -AllScripts $scripts
    Save-ImportState -State $importState
}

$failedScripts = @()

for ($i = $startIndex; $i -lt $scripts.Count; $i++) {
    $script = $scripts[$i]
    $scriptPath = Join-Path $ScriptRoot $script
    $importState.lastScript = $script
    Save-ImportState -State $importState
    
    if (Test-Path $scriptPath) {
        Write-Host "`n[+] Running $script..." -ForegroundColor Green
        try {
            # Execute script with TargetDomain parameter
            & $scriptPath -TargetDomain $TargetDomain -ErrorAction Stop
            if ($importState.completedScripts -notcontains $script) {
                $importState.completedScripts += $script
            }
            $importState.failedScripts = @($importState.failedScripts | Where-Object { $_ -ne $script })
            Save-ImportState -State $importState
        } catch {
            $errMsg = $_.Exception.Message
            Write-Host "[-] CRITICAL ERROR running $($script): $errMsg" -ForegroundColor Red
            $failedScripts += $script
            if ($importState.failedScripts -notcontains $script) {
                $importState.failedScripts += $script
            }
            $importState.status = 'failed'
            Save-ImportState -State $importState

            # Prompt to continue or abort
            $choice = [System.Windows.Forms.MessageBox]::Show("Step '$script' failed.`n`nError: $errMsg`n`nDo you want to continue with the next step?`n(Click No to abort the sequence)", "Import Step Failed", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Error)
            if ($choice -eq 'No') {
                Write-Host "Import sequence aborted by user." -ForegroundColor Red
                break
            }
        }
    } else {
        Write-Host "[-] Script not found: $scriptPath" -ForegroundColor Red
        $failedScripts += $script
        if ($importState.failedScripts -notcontains $script) {
            $importState.failedScripts += $script
        }
        $importState.status = 'failed'
        Save-ImportState -State $importState
    }
}

# --- Summary ---
if ($failedScripts.Count -gt 0) {
    Write-Host "`n=== Import Sequence Completed with Errors ===" -ForegroundColor Red
    Write-Host "The following steps failed:" -ForegroundColor Red
    foreach ($failed in $failedScripts) {
        Write-Host " - $failed" -ForegroundColor Red
    }
    # Attempt to open log file
    try {
        $config = Get-ADMigrationConfig
        $LogPath = $config.LogRoot
        $LogFile = Join-Path $LogPath "$(Get-Date -Format 'yyyy-MM-dd').log"
        if (Test-Path $LogFile) {
            Write-Host "`nOpening log file: $LogFile" -ForegroundColor Yellow
            Invoke-Item $LogFile
        }
    } catch {}
} else {
    $importState.status = 'completed'
    Save-ImportState -State $importState
    Write-Host "`n=== Import Sequence Complete ===" -ForegroundColor Cyan
    [System.Windows.Forms.MessageBox]::Show("All import scripts completed successfully.`n`nPlease proceed to the validation phase.", "Import Complete", "OK", "Information")
}