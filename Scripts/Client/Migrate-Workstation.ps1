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

# 2. Pre-Migration Checks
Write-Host "`n[*] Performing pre-migration checks on local system..." -ForegroundColor Cyan

# Check for File Shares
$customShares = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { -not $_.Name.EndsWith('$') -and $_.Name -notin @('SYSVOL', 'NETLOGON', 'print$') }
if ($customShares) {
    Write-Host "[!] WARNING: This workstation is hosting active file shares!" -ForegroundColor Yellow
    foreach ($share in $customShares) {
        Write-Host "    - Share Name: $($share.Name) (Path: $($share.Path))" -ForegroundColor Yellow
    }
    Write-Host "    Changing domains may break network access for other users relying on these shares." -ForegroundColor Yellow
} else {
    Write-Host "[+] No custom file shares detected." -ForegroundColor Green
}

# Check for Custom Local Administrators
Write-Host "`n[*] Checking Local Administrators group..." -ForegroundColor Cyan
try {
    # Find the local Administrators group securely via well-known SID (S-1-5-32-544)
    $adminGroup = Get-LocalGroup | Where-Object { $_.SID -match "S-1-5-32-544" } | Select-Object -First 1
    if ($adminGroup) {
        $customAdmins = Get-LocalGroupMember -Group $adminGroup -ErrorAction Stop | Where-Object {
            $_.Name -notmatch "\\Administrator$" -and
            $_.Name -notmatch "\\Domain Admins$" -and
            $_.Name -notmatch "\\Enterprise Admins$" -and
            $_.SID -notmatch "-500$" # Exclude the built-in local administrator account
        }
        
        if ($customAdmins) {
            Write-Host "[!] WARNING: Found non-standard accounts in the local Administrators group!" -ForegroundColor Yellow
            foreach ($admin in $customAdmins) {
                Write-Host "    - $($admin.Name) ($($admin.ObjectClass) / $($admin.PrincipalSource))" -ForegroundColor Yellow
            }
            Write-Host "    If these domain accounts were added manually (not via GPO), they will LOSE ACCESS after joining the new domain." -ForegroundColor Yellow
        } else {
            Write-Host "[+] Only standard built-in accounts found in local Administrators." -ForegroundColor Green
        }
    }
} catch {
    Write-Host "[-] Could not query local Administrators group." -ForegroundColor DarkGray
}

# Check for Applied Computer GPOs
Write-Host "`n[*] Querying applied Computer Group Policies..." -ForegroundColor Cyan
try {
    $appliedGPOs = Get-CimInstance -Namespace root\rsop\computer -ClassName RSOP_GPO -ErrorAction Stop | Where-Object { $_.Name -ne "Local Group Policy" }
    if ($appliedGPOs) {
        Write-Host "[!] NOTE: The following Group Policies are currently applied to this computer:" -ForegroundColor Yellow
        $appliedGPOs | Select-Object -Unique Name | ForEach-Object { Write-Host "    - $($_.Name)" -ForegroundColor Yellow }
        Write-Host "    These policies will no longer apply once joined to the new domain unless equivalents exist in the target." -ForegroundColor Yellow
    } else {
        Write-Host "[+] No domain GPOs currently applied to this computer." -ForegroundColor Green
    }
} catch {
    Write-Host "[-] Could not query RSOP for applied GPOs." -ForegroundColor DarkGray
}
Write-Host ""

# Check for BitLocker
Write-Host "[*] Checking BitLocker status on OS Drive..." -ForegroundColor Cyan
try {
    $bitlocker = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
    if ($bitlocker.ProtectionStatus -eq 'On') {
        Write-Host "[!] CRITICAL WARNING: BitLocker is ENABLED on the OS drive!" -ForegroundColor Red
        Write-Host "    Changing domain membership and rebooting can sometimes trigger BitLocker recovery." -ForegroundColor Yellow
        Write-Host "    Ensure you have the BitLocker Recovery Key saved externally before proceeding," -ForegroundColor Yellow
        Write-Host "    OR temporarily suspend BitLocker before continuing." -ForegroundColor Yellow
    } else {
        Write-Host "[+] BitLocker is not actively protecting the OS drive." -ForegroundColor Green
    }
} catch {
    Write-Host "[-] Could not verify BitLocker status. Ensure you have recovery keys if encrypted." -ForegroundColor DarkGray
}
Write-Host ""

# 3. Profile Warning
$profileMsg = "WARNING: Domain Change Pending`n`nPlease review the console window behind this prompt for any active file shares, custom local administrators, applied GPOs, or BitLocker warnings discovered on this machine.`n`nJoining a new domain will create NEW, empty user profiles for anyone who logs in. Their existing files (Documents, Desktop, etc.) will remain safely on the hard drive in their old 'C:\Users\' folder, but will need to be manually copied over.`n`n(If you need to migrate profiles seamlessly, cancel this script and use a third-party endpoint tool like ForensiT Profwiz).`n`nDo you want to continue?"
$warnResult = [System.Windows.Forms.MessageBox]::Show($profileMsg, "Profile Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

if ($warnResult -ne [System.Windows.Forms.DialogResult]::Yes) {
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
    $restartResult = [System.Windows.Forms.MessageBox]::Show($restartMsg, "Success - Restart Required", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
    
    if ($restartResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        Restart-Computer -Force
    } else {
        Write-Host "[!] Remember to restart this computer before attempting to log in." -ForegroundColor Yellow
    }
} catch {
    $errMsg = $_.Exception.Message
    if ($errMsg -match "already exists" -or $errMsg -match "2224") {
        Write-Host "[-] ERROR: A computer account named '$env:COMPUTERNAME' already exists in the target domain!" -ForegroundColor Red
        Write-Host "    To prevent hijacking, the join operation was safely aborted." -ForegroundColor Yellow
        [System.Windows.Forms.MessageBox]::Show("A computer named '$env:COMPUTERNAME' already exists in '$TargetDomain'.`n`nThe domain join was safely aborted to prevent overwriting the existing object. Please resolve the naming collision before trying again.", "Name Collision Detected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop)
    } else {
        Write-Host "[-] ERROR joining domain: $errMsg" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("Failed to join domain: $errMsg", "Migration Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}