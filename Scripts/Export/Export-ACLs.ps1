<#
.SYNOPSIS
    Export Access Control Lists (ACLs) for OUs.

.DESCRIPTION
    Retrieves security descriptors for all Organizational Units in the source domain.
    Exports the Discretionary Access Control List (DACL) to CSV for analysis.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'Security'

# Ensure export directory exists
if (-not (Test-Path $ExportPath)) { New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null }

if (-not $SourceDomain) {
    $SourceDomain = (Get-ADDomain).DNSRoot
    Write-Log -Message "Source domain not specified, using current: $SourceDomain" -Level INFO
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = Join-Path $ExportPath "ACLs_OUs_$timestamp.csv"

Write-Log -Message "Starting ACL Export for $SourceDomain..." -Level INFO

try {
    Invoke-Safely -ScriptBlock {
        # Ensure AD Drive is mounted for Get-Acl
        if (-not (Get-PSDrive -Name AD -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name AD -PSProvider ActiveDirectory -Root "" -Server $SourceDomain | Out-Null
        }

        $OUs = Get-ADOrganizationalUnit -Filter * -Server $SourceDomain
        Write-Log -Message "Found $($OUs.Count) OUs to scan." -Level INFO

        $ACLData = [System.Collections.Generic.List[PSObject]]::new()
        $count = 0
        $total = $OUs.Count

        foreach ($ou in $OUs) {
            $count++
            if ($count % 10 -eq 0) { Write-Progress -Activity "Exporting ACLs" -Status "Processing OU $count of $total" -PercentComplete (($count / $total) * 100) }

            try {
                # Get-Acl requires the AD drive path
                $acl = Get-Acl -Path "AD:\$($ou.DistinguishedName)" -ErrorAction Stop
                
                foreach ($access in $acl.Access) {
                    $ACLData.Add([PSCustomObject]@{
                        DistinguishedName     = $ou.DistinguishedName
                        IdentityReference     = "$($access.IdentityReference)"
                        ActiveDirectoryRights = $access.ActiveDirectoryRights.ToString()
                        AccessControlType     = $access.AccessControlType.ToString()
                        IsInherited           = $access.IsInherited
                        InheritanceFlags      = $access.InheritanceFlags.ToString()
                        PropagationFlags      = $access.PropagationFlags.ToString()
                    })
                }
            } catch {
                Write-Log -Message "Failed to get ACL for $($ou.DistinguishedName): $_" -Level WARN
            }
        }

        $ACLData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Exported $($ACLData.Count) ACEs to $csvPath" -Level INFO

    } -Operation "Export OU ACLs"

    Write-Host "✅ ACL Export Complete: $csvPath" -ForegroundColor Green
} catch {
    Write-Log -Message "Failed to export ACLs: $_" -Level ERROR
    throw "ACL export failed. Check logs."
}