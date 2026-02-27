<#
.SYNOPSIS
    Export ACLs from OUs for analysis.

.DESCRIPTION
    Retrieves security descriptors for all OUs to analyze permissions and delegation.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "ADMigration module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'Security'

# Ensure export directory exists
if (-not (Test-Path $ExportPath)) { New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null }

if (-not $SourceDomain) { $SourceDomain = Read-Host "Enter source domain name" }

Write-Log -Message "Starting ACL export from domain: $SourceDomain" -Level INFO

try {
    Invoke-Safely -ScriptBlock {
        # Get all OUs
        $OUs = Get-ADOrganizationalUnit -Filter * -Server $SourceDomain
        Write-Log -Message "Found $($OUs.Count) OUs to scan for ACLs" -Level INFO
        
        $ACLData = @()

        foreach ($ou in $OUs) {
            try {
                # Get-Acl needs the AD: drive path
                $acl = Get-Acl -Path "AD:\$($ou.DistinguishedName)"
                
                foreach ($access in $acl.Access) {
                    $ACLData += [PSCustomObject]@{
                        DistinguishedName     = $ou.DistinguishedName
                        IdentityReference     = $access.IdentityReference.ToString()
                        AccessControlType     = $access.AccessControlType.ToString()
                        ActiveDirectoryRights = $access.ActiveDirectoryRights.ToString()
                        IsInherited           = $access.IsInherited
                        InheritanceFlags      = $access.InheritanceFlags.ToString()
                        PropagationFlags      = $access.PropagationFlags.ToString()
                    }
                }
            } catch {
                Write-Log -Message "Failed to get ACL for $($ou.DistinguishedName): $_" -Level WARN
            }
        }
        
        $outputFile = Join-Path $ExportPath "ACLs_OUs_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $ACLData | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
        
        Write-Log -Message "Exported $($ACLData.Count) ACEs to $outputFile" -Level INFO
        Write-Host "ACL export complete: $outputFile"
        
    } -Operation "Export ACLs from $SourceDomain"
    
} catch {
    Write-Log -Message "Failed to export ACLs: $_" -Level ERROR
}