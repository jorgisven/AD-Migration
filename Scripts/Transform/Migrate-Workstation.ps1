<#
.SYNOPSIS
    Migrates a workstation to the target Active Directory domain.

.DESCRIPTION
    This script is designed to be run locally on end-user workstations. 
    It removes the computer from its current domain and joins it to the specified target domain.
    
    IMPORTANT: This script DOES NOT perform local profile translation. When the user logs in
    to the new domain, they will receive a fresh desktop profile.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain,

    [Parameter(Mandatory = $false)]
    [string]$TargetOU
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

Write-Host "=== Active Directory Workstation Migration ===" -ForegroundColor Cyan

# 1. Admin Check
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $msg = "This script must be run as an Administrator to change the domain membership of this computer."
    [System.Windows.Forms.MessageBox]::Show($msg, "Administrator Privileges Required", "OK", "Error")
    Write-Host "[-] FATAL: Administrator privileges required." -ForegroundColor Red
    exit
}

# 2. Profile Warning
$profileMsg = "WARNING: Domain Change Pending`n`nJoining a new domain will create NEW, empty user profiles for anyone who logs in.`n`nTheir existing files (Documents, Desktop, etc.) will remain safely on the hard drive in their old 'C:\Users\' folder, but will need to be manually copied over.`n`n(If you need to migrate profiles seamlessly, cancel this script and use a third-party endpoint tool like ForensiT Profwiz).`n`nDo you want to continue?"
$warnResult = [System.Windows.Forms.MessageBox]::Show($profileMsg, "Profile Warning", "YesNo", "Warning")

if ($warnResult -ne 'Yes') {
    Write-Host "[-] Migration cancelled by user." -ForegroundColor Yellow
    exit
}

# 3. Get Target Domain
if ([string]::IsNullOrWhiteSpace($TargetDomain)) {
    $TargetDomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the FQDN of the TARGET domain (e.g., target.local):", "Target Domain", "")
    if ([string]::IsNullOrWhiteSpace($TargetDomain)) {
        Write-Host "[-] Target domain cannot be empty. Exiting." -ForegroundColor Red
        exit
    }
}

# 4. Get Credentials
Write-Host "`nPlease enter credentials authorized to join computers to '$TargetDomain'." -ForegroundColor Yellow
Write-Host "(Typically a Domain Admin or an account delegated to the target OU)." -ForegroundColor Yellow
try {
    $cred = Get-Credential -Message "Enter credentials for $TargetDomain (e.g., TARGET\Admin)"
} catch {
    Write-Host "[-] Credential prompt cancelled. Exiting." -ForegroundColor Red
    exit
}

# 5. Perform the Join
Write-Host "`n[*] Attempting to join '$env:COMPUTERNAME' to '$TargetDomain'..." -ForegroundColor Cyan

try {
    $joinParams = @{
        DomainName = $TargetDomain
        Credential = $cred
        Force      = $true  # Forces unjoin from current domain if applicable
        ErrorAction = 'Stop'
    }
    
    if (-not [string]::IsNullOrWhiteSpace($TargetOU)) {
        $joinParams.OUPath = $TargetOU
    }

    Add-Computer @joinParams
    
    Write-Host "[+] Successfully joined $TargetDomain!" -ForegroundColor Green
    
    $restartMsg = "Welcome to the $TargetDomain domain!`n`nYou must restart this computer to apply these changes.`n`nRestart now?"
    $restartResult = [System.Windows.Forms.MessageBox]::Show($restartMsg, "Success - Restart Required", "YesNo", "Information")
    
    if ($restartResult -eq 'Yes') {
        Restart-Computer -Force
    } else {
        Write-Host "[!] Remember to restart this computer before attempting to log in." -ForegroundColor Yellow
    }
} catch {
    Write-Host "[-] ERROR joining domain: $($_.Exception.Message)" -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show("Failed to join domain: $($_.Exception.Message)", "Migration Error", "OK", "Error")
}