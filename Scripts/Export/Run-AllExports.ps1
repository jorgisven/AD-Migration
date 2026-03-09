<#
.SYNOPSIS
    Orchestrate the full export process for a specific domain.

.DESCRIPTION
    Runs all individual export scripts (OUs, GPOs, WMI, Accounts) in sequence against a target domain.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDomain
)

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# List of export scripts to run
$scripts = @(
    "Export-OUs.ps1",
    "Export-GPOReports.ps1",
    "Export-WMIFilters.ps1",
    "Export-AccountData.ps1",
    "Export-ACLs.ps1",
    "Export-DNS.ps1"
)

Write-Host "=== Starting Full Export for Domain: $SourceDomain ===" -ForegroundColor Cyan

# GUI Notice: Domain Controllers are not exported
Add-Type -AssemblyName System.Windows.Forms

# Check environment context for elevation on DC
$isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$isDC = (Get-CimInstance Win32_OperatingSystem).ProductType -eq 2

if ($isDC -and -not $isElevated) {
    $elevMsg = "You are running on a Domain Controller without Administrator privileges.`n`nSome export steps (like DNS) require elevation to access local resources.`n`nPlease restart PowerShell as Administrator."
    $elevResult = [System.Windows.Forms.MessageBox]::Show($elevMsg, "Elevation Required", [System.Windows.Forms.MessageBoxButtons]::AbortRetryIgnore, [System.Windows.Forms.MessageBoxIcon]::Error)
    if ($elevResult -eq 'Abort') {
        Write-Host "Exiting per user request. Please run as Administrator." -ForegroundColor Red
        exit
    }
    Write-Host "WARNING: User chose to ignore elevation warning on DC. Some steps may fail." -ForegroundColor Yellow
}

$dcMsg = "NOTE: Domain Controller computer accounts and replication data are NOT included in this export.`n`nDomain Controllers must be manually built and promoted in the target domain.`n`nClick OK to acknowledge and continue."
$dcResult = [System.Windows.Forms.MessageBox]::Show($dcMsg, "Migration Notice", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)

if ($dcResult -eq 'Cancel') {
    Write-Host "Export cancelled by user." -ForegroundColor Yellow
    exit
}

$failedScripts = @()

foreach ($script in $scripts) {
    $scriptPath = Join-Path $ScriptRoot $script
    
    if (Test-Path $scriptPath) {
        Write-Host "`n[+] Running $script..." -ForegroundColor Green
        try {
            # Execute script with SourceDomain parameter
            & $scriptPath -SourceDomain $SourceDomain
        } catch {
            Write-Host "[-] Error running $($script): $_" -ForegroundColor Red
            $failedScripts += $script
        }
    } else {
        Write-Host "[-] Script not found: $scriptPath" -ForegroundColor Red
        $failedScripts += $script
    }
}

if ($failedScripts.Count -gt 0) {
    Write-Host "`n=== Export Sequence Completed with Errors ===" -ForegroundColor Red
    Write-Host "The following steps failed:" -ForegroundColor Red
    foreach ($failed in $failedScripts) {
        Write-Host " - $failed" -ForegroundColor Red
    }

    # Attempt to open log file
    $LogPath = Join-Path (Join-Path (Join-Path $env:USERPROFILE "Documents") "ADMigration") "Logs"
    $LogFile = Join-Path $LogPath "$(Get-Date -Format 'yyyy-MM-dd').log"
    if (Test-Path $LogFile) {
        Write-Host "`nOpening log file: $LogFile" -ForegroundColor Yellow
        Invoke-Item $LogFile
    }
} else {
    # Prompt to create Migration Package
    $pkgScript = Join-Path $ScriptRoot "Export-MigrationPackage.ps1"
    if (Test-Path $pkgScript) {
        $pkgResult = [System.Windows.Forms.MessageBox]::Show("Do you want to create a Migration Package (ZIP) now?", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($pkgResult -eq 'Yes') {
            & $pkgScript
        }
    }
    Write-Host "`n=== Export Sequence Complete ===" -ForegroundColor Cyan
}