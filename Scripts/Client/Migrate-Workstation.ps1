<#
.SYNOPSIS
    Migrates a workstation to the target Active Directory domain with graceful DNS configuration.

.DESCRIPTION
    This script is designed to be run locally on end-user workstations. 
    It removes the computer from its current domain and joins it to the specified target domain.
    
    NEW FEATURES (with DNS Migration Support):
    - Pre-migration network diagnostics showing current IPv4 configuration and subnet
    - Warnings about current DNS servers that may not work after domain join
    - Optional automatic DNS server configuration after successful domain join
    - DNS resolution validation after domain join (tests LDAP SRV records, etc.)
    - Enhanced error diagnostics with specific guidance for common failure scenarios
    - Support for loading target DNS servers from migration metadata files
    
    IMPORTANT NOTES:
    - This script DOES NOT perform local profile translation. Users will get fresh desktop profiles.
    - DNS server configuration typically requires local Administrator rights.
    - Target DNS servers should be specified to ensure proper domain resolution.
    - The script will prompt for domain join credentials (or you can pass them via parameters).
    - A restart is REQUIRED to complete the domain join.

.PARAMETER TargetDomain
    FQDN of the target domain (e.g., target.local). If not provided, you will be prompted.

.PARAMETER TargetOU
    Optional DN of the target Organizational Unit (e.g., "OU=Workstations,DC=target,DC=local").
    If not provided, the computer will be placed in the default Computer container.

.PARAMETER TargetDnsServers
    IPv4 addresses of DNS servers to configure after domain join (comma-separated or as array).
    If not provided, you will be prompted interactively.

.PARAMETER MigrationMetadataPath
    Optional path to a migration metadata file (JSON or XML) that contains target DNS server info.
    Useful for automated/unattended migrations. Expected properties: TargetDnsServers or DnsServers.

.EXAMPLE
    # Interactive migration with DNS server prompt
    .\Migrate-Workstation.ps1 -TargetDomain "target.local"

.EXAMPLE
    # Specify DNS servers directly
    .\Migrate-Workstation.ps1 -TargetDomain "target.local" -TargetDnsServers @("192.168.1.10", "192.168.1.11")

.EXAMPLE
    # Use migration metadata for automated deployment
    .\Migrate-Workstation.ps1 -TargetDomain "target.local" -MigrationMetadataPath "\\server\migration\metadata.json"

.EXAMPLE
    # Specify both domain and OU
    .\Migrate-Workstation.ps1 -TargetDomain "target.local" -TargetOU "OU=Workstations,DC=target,DC=local" -TargetDnsServers @("192.168.1.10", "192.168.1.11")
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain,

    [Parameter(Mandatory = $false)]
    [string]$TargetOU,

    [Parameter(Mandatory = $false)]
    [string[]]$TargetDnsServers,

    [Parameter(Mandatory = $false)]
    [string]$MigrationMetadataPath
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# ============================================================================
# Helper Functions for Network and DNS Operations
# ============================================================================

function Get-LocalIPv4Configuration {
    <#
    .SYNOPSIS
        Retrieves local IPv4 interfaces and their subnets.
    #>
    try {
        $adapters = @()
        
        # Try Get-NetIPAddress first (PowerShell 3+)
        if (Get-Command Get-NetIPAddress -ErrorAction SilentlyContinue) {
            $adapters = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | 
                Where-Object { $_.IPAddress -notmatch "^127\.|^169\.254\." } |
                ForEach-Object {
                    @{
                        IPAddress = $_.IPAddress
                        PrefixLength = $_.PrefixLength
                        InterfaceAlias = $_.InterfaceAlias
                    }
                }
        }
        
        # Fallback to WMI if Get-NetIPAddress unavailable
        if (-not $adapters) {
            $adapters = Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled='TRUE'" -ErrorAction SilentlyContinue |
                Where-Object { $_.IPAddress } |
                ForEach-Object {
                    foreach ($ip in $_.IPAddress) {
                        if ($ip -notmatch "^127\.|^169\.254\.") {
                            @{
                                IPAddress = $ip
                                SubnetMask = if ($_.IPSubnet) { $_.IPSubnet[0] } else { $null }
                                Description = $_.Description
                            }
                        }
                    }
                }
        }
        
        return $adapters
    } catch {
        Write-Host "[-] Could not retrieve IPv4 configuration: $($_.Exception.Message)" -ForegroundColor DarkGray
        return @()
    }
}

function Get-CurrentDnsServers {
    <#
    .SYNOPSIS
        Retrieves current DNS server configuration.
    #>
    try {
        $dnsServers = @()
        
        if (Get-Command Get-DnsClientServerAddress -ErrorAction SilentlyContinue) {
            $dnsServers = Get-DnsClientServerAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.ServerAddresses } |
                Select-Object -ExpandProperty ServerAddresses
        } else {
            $dnsServers = Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled='TRUE'" -ErrorAction SilentlyContinue |
                Where-Object { $_.DNSServerSearchOrder } |
                Select-Object -ExpandProperty DNSServerSearchOrder
        }
        
        return $dnsServers | Select-Object -Unique
    } catch {
        Write-Host "[-] Could not retrieve DNS configuration: $($_.Exception.Message)" -ForegroundColor DarkGray
        return @()
    }
}

function Set-DnsServers {
    <#
    .SYNOPSIS
        Configures DNS servers on all active network adapters.
    #>
    param([string[]]$DnsServers)
    
    if (-not $DnsServers -or $DnsServers.Count -eq 0) {
        Write-Host "[-] No DNS servers provided. Skipping DNS configuration." -ForegroundColor Yellow
        return $false
    }
    
    try {
        if (Get-Command Set-DnsClientServerAddress -ErrorAction SilentlyContinue) {
            Get-NetAdapter | Where-Object { $_.Status -eq 'Up' } | ForEach-Object {
                try {
                    Set-DnsClientServerAddress -InterfaceIndex $_.InterfaceIndex -ServerAddresses $DnsServers -Confirm:$false
                    Write-Host "[+] Configured DNS servers on $($_.Name): $(($DnsServers -join ', '))" -ForegroundColor Green
                } catch {
                    Write-Host "[!] Could not configure DNS on $($_.Name): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        } else {
            # Fallback: Use WMI to set DNS (less flexible but works on older systems)
            Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled='TRUE'" -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $_.SetDNSServerSearchOrder([System.Collections.ArrayList]$DnsServers) | Out-Null
                    Write-Host "[+] Configured DNS servers on adapter: $(($DnsServers -join ', '))" -ForegroundColor Green
                } catch {
                    Write-Host "[!] Could not configure DNS: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }
        
        return $true
    } catch {
        Write-Host "[-] DNS configuration failed: $($_.Exception.Message)" -ForegroundColor Yellow
        return $false
    }
}

function Test-DnsResolution {
    <#
    .SYNOPSIS
        Tests DNS resolution for the target domain and critical hosts.
    #>
    param([string]$Domain)
    
    Write-Host "`n[*] Validating DNS resolution for '$Domain'..." -ForegroundColor Cyan
    
    $results = @{}
    $testNames = @(
        $Domain,
        "_ldap._tcp.dc._msdcs.$Domain",
        "gc._tcp.dc._msdcs.$Domain"
    )
    
    foreach ($testName in $testNames) {
        try {
            $resolved = Resolve-DnsName -Name $testName -Type A -ErrorAction Stop
            $results[$testName] = $resolved.IPAddress
            Write-Host "[+] $testName resolves to $($resolved.IPAddress)" -ForegroundColor Green
        } catch {
            $results[$testName] = $null
            Write-Host "[!] Could not resolve $testName" -ForegroundColor Yellow
        }
    }
    
    # Return success if at least the domain and one critical SRV record resolved
    return ($null -ne $results[$Domain] -or ($results.Values | Where-Object { $_ } | Measure-Object).Count -gt 0)
}

function Import-MigrationMetadata {
    <#
    .SYNOPSIS
        Attempts to load DNS server configuration from migration metadata file.
    #>
    param([string]$MetadataPath)
    
    if (-not (Test-Path $MetadataPath)) {
        return $null
    }
    
    try {
        $metadata = @{}
        
        if ($MetadataPath -match "\.json$") {
            $metadata = Get-Content $MetadataPath | ConvertFrom-Json
        } elseif ($MetadataPath -match "\.xml$") {
            [xml]$metadata = Get-Content $MetadataPath
        }
        
        if ($metadata -and $metadata.TargetDnsServers) {
            return $metadata.TargetDnsServers
        }
        
        if ($metadata.DnsServers) {
            return $metadata.DnsServers
        }
    } catch {
        Write-Host "[-] Could not parse migration metadata: $($_.Exception.Message)" -ForegroundColor DarkGray
    }
    
    return $null
}

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

# Check Network Interface Configuration
Write-Host "[*] Checking Network Interface configuration..." -ForegroundColor Cyan
$localInterfaces = Get-LocalIPv4Configuration
if ($localInterfaces) {
    Write-Host "[+] Found network interfaces:" -ForegroundColor Green
    foreach ($iface in $localInterfaces) {
        if ($iface.InterfaceAlias) {
            Write-Host "    - $($iface.InterfaceAlias): $($iface.IPAddress)/$($iface.PrefixLength)" -ForegroundColor Green
        } else {
            $subnet = if ($iface.SubnetMask) { $iface.SubnetMask } else { "Unknown" }
            Write-Host "    - $($iface.Description): $($iface.IPAddress) / $subnet" -ForegroundColor Green
        }
    }
} else {
    Write-Host "[!] Could not retrieve network interface configuration." -ForegroundColor Yellow
}

# Check Current DNS Servers
Write-Host "`n[*] Checking current DNS server configuration..." -ForegroundColor Cyan
$currentDns = Get-CurrentDnsServers
if ($currentDns) {
    Write-Host "[!] Current DNS servers:" -ForegroundColor Yellow
    foreach ($dns in $currentDns) {
        Write-Host "    - $dns" -ForegroundColor Yellow
    }
    Write-Host "    These DNS servers may not function after the domain join if they belong to the current domain." -ForegroundColor Yellow
} else {
    Write-Host "[+] No DNS servers currently configured (DHCP)." -ForegroundColor Green
}

Write-Host ""
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

# 4. Determine Target DNS Servers
Write-Host "`n[*] Determining target DNS servers..." -ForegroundColor Cyan

# Try to load from migration metadata first
if (-not $TargetDnsServers -and $MigrationMetadataPath) {
    $metadataDns = Import-MigrationMetadata -MetadataPath $MigrationMetadataPath
    if ($metadataDns) {
        $TargetDnsServers = $metadataDns
        Write-Host "[+] Loaded target DNS servers from migration metadata: $(($TargetDnsServers -join ', '))" -ForegroundColor Green
    }
}

# If still no DNS servers, prompt the user
if (-not $TargetDnsServers -or $TargetDnsServers.Count -eq 0) {
    $dnsMsg = "Would you like to specify DNS servers for the target domain?`n`nClicking 'Yes' lets you enter DNS server IP addresses that will be configured after joining '$TargetDomain'.`n`nClicking 'No' will leave DNS auto-configured (DHCP).`n`nRecommended: Use the target domain's DNS servers to ensure proper resolution after joining."
    $dnsResult = [System.Windows.Forms.MessageBox]::Show($dnsMsg, "Configure Target DNS Servers?", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    
    if ($dnsResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        $dnsInput = [Microsoft.VisualBasic.Interaction]::InputBox("Enter target DNS servers (comma-separated, e.g., 192.168.1.10, 192.168.1.11):", "Target DNS Servers", "")
        if (-not [string]::IsNullOrWhiteSpace($dnsInput)) {
            $TargetDnsServers = @($dnsInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            Write-Host "[+] Will apply DNS servers: $(($TargetDnsServers -join ', '))" -ForegroundColor Green
        } else {
            Write-Host "[*] No DNS servers specified. Will use DHCP/default configuration." -ForegroundColor Cyan
        }
    } else {
        Write-Host "[*] DNS will be auto-configured. Using DHCP or existing settings." -ForegroundColor Cyan
    }
}

# 5. Get Credentials
Write-Host "`n[*] Obtaining domain join credentials..." -ForegroundColor Cyan
Write-Host "Please enter credentials authorized to join computers to '$TargetDomain'." -ForegroundColor Yellow
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
    
    # Configure DNS servers if specified
    if ($TargetDnsServers -and $TargetDnsServers.Count -gt 0) {
        Write-Host "`n[*] Configuring DNS servers for the target domain..." -ForegroundColor Cyan
        $dnsConfigSuccess = Set-DnsServers -DnsServers $TargetDnsServers
        if ($dnsConfigSuccess) {
            Write-Host "[+] DNS servers configured successfully." -ForegroundColor Green
        } else {
            Write-Host "[!] DNS configuration encountered issues. This can be manually corrected after restart." -ForegroundColor Yellow
        }
    }
    
    # Test DNS resolution before restart warning
    Write-Host "`n[*] Testing DNS resolution..." -ForegroundColor Cyan
    $dnsResolutionOk = Test-DnsResolution -Domain $TargetDomain
    if ($dnsResolutionOk) {
        Write-Host "[+] DNS resolution successful for $TargetDomain!" -ForegroundColor Green
    } else {
        Write-Host "[!] DNS resolution failed or incomplete. This may resolve after restart and domain replication." -ForegroundColor Yellow
    }
    
    $restartMsg = "Welcome to the $TargetDomain domain!`n`n[Domain Join: Success]`n[DNS Configuration: $(@('Failed','Successful')[[int]$dnsConfigSuccess])]`n[DNS Resolution: $(@('Failed','Successful')[[int]$dnsResolutionOk])]`n`nYou must restart this computer to apply these changes and enable all domain functionality.`n`nRestart now?"
    $restartResult = [System.Windows.Forms.MessageBox]::Show($restartMsg, "Success - Restart Required", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Information)
    
    if ($restartResult -eq [System.Windows.Forms.DialogResult]::Yes) {
        Restart-Computer -Force
    } else {
        Write-Host "[!] Remember to restart this computer before attempting to log in." -ForegroundColor Yellow
    }
} catch {
    $errMsg = $_.Exception.Message
    Write-Host "[-] ERROR joining domain: $errMsg" -ForegroundColor Red
    
    # Provide intelligent error diagnostics
    if ($errMsg -match "already exists" -or $errMsg -match "2224") {
        Write-Host "[-] ERROR: A computer account named '$env:COMPUTERNAME' already exists in the target domain!" -ForegroundColor Red
        Write-Host "    To prevent hijacking, the join operation was safely aborted." -ForegroundColor Yellow
        $errorDetail = "A computer account named '$env:COMPUTERNAME' already exists in '$TargetDomain'.`n`nThe domain join was safely aborted to prevent overwriting the existing object.`n`nPlease resolve the naming collision before trying again (delete the old account or rename this computer)."
        [System.Windows.Forms.MessageBox]::Show($errorDetail, "Name Collision Detected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Stop)
    } elseif ($errMsg -match "credentials|credential|authentication|logon|password|775") {
        Write-Host "[-] ERROR: Authentication failed. Check the provided credentials and target domain name." -ForegroundColor Red
        Write-Host "[*] Diagnostic tip: Verify the account has sufficient rights to join computers to the target domain or OU." -ForegroundColor Yellow
        $errorDetail = "Failed to authenticate for domain join.`n`n- Check credentials are correct`n- Verify the account is authorized to join computers to '$TargetDomain'`n- Ensure you have DNS/network connectivity to the domain`n`nCause: $errMsg"
        [System.Windows.Forms.MessageBox]::Show($errorDetail, "Authentication Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } elseif ($errMsg -match "network|connectivity|reachable|NETERR|name resolution") {
        Write-Host "[-] ERROR: Network connectivity issue. Cannot reach the target domain." -ForegroundColor Red
        Write-Host "[*] Diagnostic tip: Check network connectivity and DNS resolution for '$TargetDomain'." -ForegroundColor Yellow
        Write-Host "[*] You may need to configure DNS servers before retrying." -ForegroundColor Yellow
        $errorDetail = "Failed to reach the target domain. Check the following:`n`n- Network connectivity to domain controllers`n- DNS resolution for '$TargetDomain'`n- Firewall rules allowing domain join traffic (TCP 445, 389, etc.)`n- Target DNS servers are configured (current servers may not resolve the new domain)`n`nCause: $errMsg"
        [System.Windows.Forms.MessageBox]::Show($errorDetail, "Network or Connectivity Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } else {
        Write-Host "[*] Review the error message above for details on how to resolve this issue." -ForegroundColor Cyan
        $errorDetail = "Failed to join domain: $errMsg`n`nPlease review the console output above for more details and troubleshooting steps."
        [System.Windows.Forms.MessageBox]::Show($errorDetail, "Migration Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}