<#
.SYNOPSIS
    Rebuild GPO link structure.

.DESCRIPTION
    Uses XML report data to recreate link locations and link order in target domain.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

$config = Get-ADMigrationConfig
$ReportPath = Join-Path $config.ExportRoot 'GPO_Reports'
$MapPath = Join-Path $config.TransformRoot 'Mapping'

# 1. Load OU Map (Transform Data)
$mapFiles = Get-ChildItem -Path $MapPath -Filter "OU_Map_*.csv" | Sort-Object LastWriteTime -Descending
if (-not $mapFiles) { throw "Missing OU Map (Run Transform-OUMap.ps1)" }

$OUMap = @{}
Import-Csv $mapFiles[0].FullName | ForEach-Object { $OUMap[$_.SourceDN] = $_.TargetDN }
Write-Log -Message "Loaded OU Map with $($OUMap.Count) entries" -Level INFO

# 2. Load GPO XML Reports (Export Data)
$xmlFiles = Get-ChildItem -Path $ReportPath -Filter "*.xml"
if (-not $xmlFiles) { throw "Missing GPO Reports (Run Export-GPOReports.ps1)" }

if (-not $TargetDomain) { $TargetDomain = (Get-ADDomain).DNSRoot }

Write-Log -Message "Starting GPO Link Rebuild..." -Level INFO

Invoke-Safely -ScriptBlock {
    foreach ($xmlFile in $xmlFiles) {
        [xml]$xml = Get-Content $xmlFile.FullName
        $gpoName = $xml.GPO.Name
        
        # Find LinksTo (Where this GPO is linked)
        $links = $xml.GPO.LinksTo
        
        if ($links) {
            foreach ($link in $links) {
                $sourceSOM = $link.SOMPath # e.g. "ou=sales,dc=source,dc=local"
                $targetSOM = $null
                
                # Determine Target Location
                if ($OUMap.ContainsKey($sourceSOM)) {
                    $targetSOM = $OUMap[$sourceSOM]
                } elseif ($sourceSOM -match "DC=") {
                    # If linked to Domain Root, map to Target Domain Root
                    $targetSOM = $TargetDomain # Simplified, assumes domain root link
                }
                
                if ($targetSOM) {
                    try {
                        # Check if link exists
                        # Note: Get-GPO doesn't easily show links, usually we just try to add it
                        # New-GPLink throws if GPO doesn't exist, but we assume Import-GPOs ran
                        
                        New-GPLink -Name $gpoName -Target $targetSOM -LinkEnabled ($link.Enabled -eq 'true') -Enforced ($link.NoOverride -eq 'true') -Server $TargetDomain -ErrorAction Stop | Out-Null
                        Write-Log -Message "Linked '$gpoName' to '$targetSOM'" -Level INFO
                    } catch {
                        Write-Log -Message "Failed to link '$gpoName' to '$targetSOM': $_" -Level WARN
                    }
                } else {
                    Write-Log -Message "Skipping link for '$gpoName': Source path '$sourceSOM' not found in OU Map" -Level WARN
                }
            }
        }
        
        # Restore WMI Filter Link if present
        $wmiFilter = $xml.GPO.FilterData.Name
        if ($wmiFilter) {
            try {
                $gpo = Get-GPO -Name $gpoName -Server $TargetDomain
                $gpo.WmiFilter = $wmiFilter # This property sets the link
                Write-Log -Message "Linked WMI Filter '$wmiFilter' to '$gpoName'" -Level INFO
            } catch {
                Write-Log -Message "Failed to link WMI Filter '$wmiFilter' to '$gpoName'" -Level WARN
            }
        }

        # Restore Security Filtering (Permissions)
        $securityDescriptor = $xml.GPO.SecurityDescriptor
        if ($securityDescriptor) {
            $customFiltersApplied = $false
            foreach ($permission in $securityDescriptor.Permissions.Permission) {
                # We only care about re-applying GpoApply filters
                if ($permission.Type -eq 'GpoApply') {
                    $trusteeName = $permission.Trustee.Name
                    $trusteeType = $permission.Trustee.Type # User, Group, Computer
                    
                    # Skip well-known SIDs that are handled by default or are not filters
                    if ($trusteeName -in @('NT AUTHORITY\ENTERPRISE DOMAIN CONTROLLERS', 'NT AUTHORITY\Authenticated Users', 'NT AUTHORITY\SYSTEM')) {
                        continue
                    }

                    try {
                        Write-Log -Message "Applying GpoApply security filter for '$trusteeName' on GPO '$gpoName'" -Level INFO
                        Set-GPPermission -Name $gpoName -PermissionLevel GpoApply -TargetName $trusteeName -TargetType $trusteeType -Server $TargetDomain -ErrorAction Stop
                        $customFiltersApplied = $true
                    }
                    catch {
                        Write-Log -Message "Failed to set GpoApply for '$trusteeName' on GPO '$gpoName'. Principal may not exist in target domain. Details: $_" -Level WARN
                    }
                }
            }

            # If custom filters were added, remove the default 'Authenticated Users' permission
            # to ensure the policy is correctly restricted.
            if ($customFiltersApplied) {
                try {
                    Write-Log -Message "Custom filters applied. Removing default 'Authenticated Users' GpoApply permission from '$gpoName'." -Level INFO
                    Set-GPPermission -Name $gpoName -PermissionLevel None -TargetName 'Authenticated Users' -TargetType Group -Server $TargetDomain -ErrorAction Stop
                } catch {
                    # This might fail if it was already removed or never existed, which is fine.
                    Write-Log -Message "Could not remove GpoApply from 'Authenticated Users' on GPO '$gpoName'. It may have already been removed. Details: $_" -Level WARN
                }
            }
        }
    }
} -Operation "Rebuild GPO Links"