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

$failedScripts = @()

foreach ($script in $scripts) {
    $scriptPath = Join-Path $ScriptRoot $script
    
    if (Test-Path $scriptPath) {
        Write-Host "`n[+] Running $script..." -ForegroundColor Green
        try {
            # Execute script with TargetDomain parameter
            & $scriptPath -TargetDomain $TargetDomain -ErrorAction Stop
        } catch {
            $errMsg = $_.Exception.Message
            Write-Host "[-] CRITICAL ERROR running $($script): $errMsg" -ForegroundColor Red
            $failedScripts += $script

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
    Write-Host "`n=== Import Sequence Complete ===" -ForegroundColor Cyan
    [System.Windows.Forms.MessageBox]::Show("All import scripts completed successfully.`n`nPlease proceed to the validation phase.", "Import Complete", "OK", "Information")
}