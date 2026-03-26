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
    [string]$TargetDomain,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDefaultDomainPolicies,

    [Parameter(Mandatory = $false)]
    [string]$DefaultPolicyNameSuffix = ' (Migrated Copy)',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Prompt', 'Skip', 'Rename')]
    [string]$DefaultDomainPolicyMode = 'Prompt',

    [Parameter(Mandatory = $false)]
    [switch]$RunPreflight,

    [Parameter(Mandatory = $false)]
    [switch]$PreflightOnly
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

function New-ImportStepParams {
    param(
        [string]$ScriptName,
        [string]$Domain,
        [string]$PolicyMode,
        [string]$PolicySuffix,
        [switch]$AsWhatIf
    )

    $stepParams = @{
        TargetDomain = $Domain
    }

    if ($ScriptName -in @('Import-GPOs.ps1', 'Import-GPOLinks.ps1')) {
        $stepParams.IncludeDefaultDomainPolicies = ($PolicyMode -eq 'Rename')
        $stepParams.DefaultPolicyNameSuffix = $PolicySuffix
        $stepParams.DefaultDomainPolicyMode = $PolicyMode
    }

    if ($AsWhatIf) {
        $stepParams.WhatIf = $true
    }

    return $stepParams
}

function Invoke-ImportPreflight {
    param(
        [string]$Domain,
        [string[]]$AllScripts,
        [string]$PolicyMode,
        [string]$PolicySuffix
    )

    $supportsWhatIf = @(
        'Import-OUs.ps1',
        'Import-Accounts.ps1',
        'Import-WMIFilters.ps1',
        'Import-GPOs.ps1',
        'Import-GPOLinks.ps1'
    )

    $results = @()

    Write-Host "`n=== Running Import Preflight (WhatIf Dry-Run) ===" -ForegroundColor Cyan
    Write-Log -Message "Starting import preflight dry-run for domain '$Domain'." -Level INFO

    foreach ($scriptName in $AllScripts) {
        $scriptPath = Join-Path $ScriptRoot $scriptName
        $mode = if ($supportsWhatIf -contains $scriptName) { 'WhatIf' } else { 'PrecheckOnly' }

        if (-not (Test-Path $scriptPath)) {
            $msg = "Script not found: $scriptPath"
            Write-Host "[-] Preflight failed for ${scriptName}: $msg" -ForegroundColor Red
            Write-Log -Message "Preflight failed for '$scriptName': $msg" -Level ERROR
            $results += [PSCustomObject]@{ Script = $scriptName; Mode = $mode; Status = 'Failed'; Details = $msg }
            continue
        }

        if ($mode -eq 'PrecheckOnly') {
            $msg = "WhatIf not supported by this script; only file/parameter precheck was performed."
            Write-Host "[!] Preflight partial for ${scriptName}: $msg" -ForegroundColor Yellow
            Write-Log -Message "Preflight partial for '$scriptName': $msg" -Level WARN
            $results += [PSCustomObject]@{ Script = $scriptName; Mode = $mode; Status = 'Partial'; Details = $msg }
            continue
        }

        try {
            Write-Host "[+] Preflight running $scriptName (WhatIf)..." -ForegroundColor Green
            $stepParams = New-ImportStepParams -ScriptName $scriptName -Domain $Domain -PolicyMode $PolicyMode -PolicySuffix $PolicySuffix -AsWhatIf
            & $scriptPath @stepParams -ErrorAction Stop
            $results += [PSCustomObject]@{ Script = $scriptName; Mode = $mode; Status = 'Passed'; Details = 'No terminating errors during WhatIf dry-run.' }
            Write-Log -Message "Preflight passed for '$scriptName'." -Level INFO
        } catch {
            $errText = $_.Exception.Message
            $results += [PSCustomObject]@{ Script = $scriptName; Mode = $mode; Status = 'Failed'; Details = $errText }
            Write-Host "[-] Preflight failed for ${scriptName}: $errText" -ForegroundColor Red
            Write-Log -Message "Preflight failed for '$scriptName': $errText" -Level ERROR
        }
    }

    $reportPath = Join-Path $config.ImportRoot ("Import-Preflight-Report-{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $results | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
    Write-Log -Message "Import preflight report written to '$reportPath'." -Level INFO
    Write-Host "Preflight report: $reportPath" -ForegroundColor Yellow

    $failedCount = (@($results | Where-Object { $_.Status -eq 'Failed' })).Count
    $partialCount = (@($results | Where-Object { $_.Status -eq 'Partial' })).Count
    if ($failedCount -gt 0) {
        Write-Log -Message "Import preflight completed with failures (failed=$failedCount, partial=$partialCount)." -Level WARN
    } else {
        Write-Log -Message "Import preflight completed without terminating failures (partial=$partialCount)." -Level INFO
    }

    return [PSCustomObject]@{
        Results = $results
        ReportPath = $reportPath
        FailedCount = $failedCount
        PartialCount = $partialCount
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
Write-Log -Message "Starting full import orchestrator for domain '$TargetDomain'." -Level INFO

$backupWarningMsg = @"
CRITICAL WARNING: This import modifies Active Directory objects, memberships, GPOs, and DNS in the target domain.

Before proceeding, create and verify a recoverable backup/snapshot of the target domain controllers.

There is no automated 'un-import' rollback tool in this workflow.
"@

[System.Windows.Forms.MessageBox]::Show($backupWarningMsg, "Critical Backup Requirement", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
Write-Log -Message "Displayed critical backup requirement warning before import." -Level WARN

$backupConfirm = [Microsoft.VisualBasic.Interaction]::InputBox("Type YES to confirm you have a verified domain backup/snapshot and want to continue import.", "Backup Confirmation Required", "")
if ($backupConfirm.Trim().ToUpperInvariant() -ne 'YES') {
    Write-Host "Import cancelled. Backup confirmation was not provided." -ForegroundColor Yellow
    Write-Log -Message "Import orchestrator cancelled: backup confirmation was not provided." -Level WARN
    exit
}

Write-Log -Message "Backup confirmation accepted (user typed YES)." -Level INFO

$effectiveDefaultPolicyMode = $DefaultDomainPolicyMode
if ($effectiveDefaultPolicyMode -eq 'Prompt' -and $IncludeDefaultDomainPolicies) {
    # Backward compatibility with previous behavior.
    $effectiveDefaultPolicyMode = 'Rename'
    Write-Log -Message "Legacy switch IncludeDefaultDomainPolicies was set. Effective default policy mode changed from Prompt to Rename." -Level WARN
}

if ($effectiveDefaultPolicyMode -eq 'Prompt') {
    $defaultPolicyPrompt = "Default domain policies were found in source backups.`n`nYES = Skip them (recommended)`nNO = Import as renamed copies`nCancel = Abort"
    $defaultPolicyChoice = [System.Windows.Forms.MessageBox]::Show($defaultPolicyPrompt, "Default GPO Policy Handling", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)

    if ($defaultPolicyChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
        $effectiveDefaultPolicyMode = 'Skip'
        Write-Log -Message "Default GPO policy prompt selection: Skip" -Level INFO
    } elseif ($defaultPolicyChoice -eq [System.Windows.Forms.DialogResult]::No) {
        $effectiveDefaultPolicyMode = 'Rename'
        Write-Log -Message "Default GPO policy prompt selection: Rename" -Level INFO
    } else {
        Write-Host "Import cancelled by user." -ForegroundColor Yellow
        Write-Log -Message "Import orchestrator cancelled by user at default policy handling prompt." -Level WARN
        exit
    }
}

Write-Log -Message "Default domain policy handling mode selected: $effectiveDefaultPolicyMode" -Level INFO
Write-Log -Message "Default domain policy options: Mode='$effectiveDefaultPolicyMode', Suffix='$DefaultPolicyNameSuffix'" -Level INFO

$runPreflightNow = $RunPreflight.IsPresent -or $PreflightOnly.IsPresent
if (-not $runPreflightNow) {
    $preflightPrompt = [System.Windows.Forms.MessageBox]::Show("Run import preflight dry-run (WhatIf) before making changes?`n`nYES = Run preflight now`nNO = Skip preflight and continue`nCancel = Abort", "Import Preflight", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($preflightPrompt -eq [System.Windows.Forms.DialogResult]::Yes) {
        $runPreflightNow = $true
        Write-Log -Message "User selected to run import preflight dry-run." -Level INFO
    } elseif ($preflightPrompt -eq [System.Windows.Forms.DialogResult]::No) {
        Write-Log -Message "User skipped import preflight dry-run." -Level WARN
    } else {
        Write-Host "Import cancelled by user." -ForegroundColor Yellow
        Write-Log -Message "Import orchestrator cancelled by user at preflight prompt." -Level WARN
        exit
    }
}

if ($runPreflightNow) {
    $preflightAttempt = 0
    $preflightDecision = 'Retry'
    
    do {
        $preflightAttempt++
        Write-Log -Message "Running import preflight attempt #$preflightAttempt." -Level INFO
        
        $preflight = Invoke-ImportPreflight -Domain $TargetDomain -AllScripts $scripts -PolicyMode $effectiveDefaultPolicyMode -PolicySuffix $DefaultPolicyNameSuffix
        
        if ($preflight.FailedCount -eq 0) {
            $preflightDecision = 'Passed'
            Write-Host "`n[+] Preflight passed! All steps completed without errors." -ForegroundColor Green
            Write-Log -Message "Import preflight attempt #$preflightAttempt passed (0 failures)." -Level INFO
            break
        }
        
        if (-not $PreflightOnly.IsPresent) {
            $msg = "Preflight attempt #$preflightAttempt detected $($preflight.FailedCount) failed step(s).`n`nReport: $($preflight.ReportPath)`n`nFix any issues and then:`nYES = Retry preflight`nNO = Continue to real import anyway`nCancel = Stop now"
            $preflightChoice = [System.Windows.Forms.MessageBox]::Show($msg, "Preflight Found Issues", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
            
            if ($preflightChoice -eq [System.Windows.Forms.DialogResult]::Yes) {
                $preflightDecision = 'Retry'
                Write-Host "[*] Retrying preflight scan..." -ForegroundColor Yellow
                Write-Log -Message "User selected Retry after preflight attempt #$preflightAttempt." -Level WARN
                continue
            } elseif ($preflightChoice -eq [System.Windows.Forms.DialogResult]::No) {
                $preflightDecision = 'ContinueAnyway'
                Write-Host "[!] Continuing to import despite preflight failures." -ForegroundColor Yellow
                Write-Log -Message "User selected Continue Anyway after preflight attempt #$preflightAttempt. Report: $($preflight.ReportPath)" -Level WARN
                break
            } else {
                Write-Host "Import cancelled after preflight findings." -ForegroundColor Yellow
                Write-Log -Message "Import cancelled by user after preflight attempt #$preflightAttempt. Report: $($preflight.ReportPath)" -Level WARN
                exit
            }
        } else {
            # PreflightOnly mode
            $preflightDecision = 'PreflightOnlyMode'
            Write-Host "Preflight-only mode complete. No changes will be applied." -ForegroundColor Cyan
            Write-Log -Message "Preflight-only mode complete; import execution skipped." -Level INFO
            exit
        }
    } while ($preflightDecision -eq 'Retry')
    
    if ($PreflightOnly.IsPresent) {
        Write-Host "Preflight-only mode complete. No changes will be applied." -ForegroundColor Cyan
        Write-Log -Message "Preflight-only mode complete after $preflightAttempt attempt(s); import execution skipped." -Level INFO
        exit
    }
}

$confirmResult = [System.Windows.Forms.MessageBox]::Show("This will begin creating objects in the domain '$TargetDomain'.`n`nEnsure you have reviewed all mapping files in the 'Transform' directory before proceeding.", "Confirm Import", "OKCancel", "Warning")
if ($confirmResult -ne 'OK') {
    Write-Host "Import cancelled by user." -ForegroundColor Yellow
    Write-Log -Message "Import orchestrator cancelled by user before execution." -Level WARN
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
        Write-Log -Message "Running import step '$script'." -Level INFO
        try {
            # Execute script with TargetDomain parameter
            $stepParams = New-ImportStepParams -ScriptName $script -Domain $TargetDomain -PolicyMode $effectiveDefaultPolicyMode -PolicySuffix $DefaultPolicyNameSuffix

            & $scriptPath @stepParams -ErrorAction Stop
            if ($importState.completedScripts -notcontains $script) {
                $importState.completedScripts += $script
            }
            $importState.failedScripts = @($importState.failedScripts | Where-Object { $_ -ne $script })
            Save-ImportState -State $importState
            Write-Log -Message "Import step '$script' completed successfully." -Level INFO
        } catch {
            $errMsg = $_.Exception.Message
            Write-Host "[-] CRITICAL ERROR running $($script): $errMsg" -ForegroundColor Red
            $failedScripts += $script
            if ($importState.failedScripts -notcontains $script) {
                $importState.failedScripts += $script
            }
            $importState.status = 'failed'
            Save-ImportState -State $importState
            Write-Log -Message "Import step '$script' failed: $errMsg" -Level ERROR

            # Prompt to continue or abort
            $choice = [System.Windows.Forms.MessageBox]::Show("Step '$script' failed.`n`nError: $errMsg`n`nDo you want to continue with the next step?`n(Click No to abort the sequence)", "Import Step Failed", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Error)
            if ($choice -eq 'No') {
                Write-Host "Import sequence aborted by user." -ForegroundColor Red
                Write-Log -Message "Import sequence aborted by user after step failure in '$script'." -Level WARN
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
        Write-Log -Message "Import step script not found: $scriptPath" -Level ERROR
    }
}

# --- Summary ---
if ($failedScripts.Count -gt 0) {
    Write-Host "`n=== Import Sequence Completed with Errors ===" -ForegroundColor Red
    Write-Host "The following steps failed:" -ForegroundColor Red
    foreach ($failed in $failedScripts) {
        Write-Host " - $failed" -ForegroundColor Red
    }
    $failedList = ($failedScripts | Select-Object -Unique) -join ', '
    Write-Log -Message "Import sequence completed with errors. Failed steps: $failedList" -Level ERROR
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
    Write-Log -Message "Import sequence completed successfully for domain '$TargetDomain'." -Level INFO
    [System.Windows.Forms.MessageBox]::Show("All import scripts completed successfully.`n`nPlease proceed to the validation phase.", "Import Complete", "OK", "Information")
}