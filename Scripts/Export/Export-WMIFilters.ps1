<#
.SYNOPSIS
    Export WMI filters from source domain.

.DESCRIPTION
    Captures WMI filter names, queries, and descriptions for recreation in target domain.
    WMI filters are stored in AD and need to be rewritten for domain-specific references.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) {
    throw "ADMigration module manifest missing, cannot continue."
}
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) {
    Write-Log -Message "Required function Invoke-Safely not defined after module import" -Level ERROR
    throw "Invoke-Safely unavailable"
}
$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'WMI_Filters'

# Ensure export directory exists
if (-not (Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
    Write-Log -Message "Created export directory: $ExportPath" -Level INFO
}

# Prompt for source domain if not provided
if (-not $SourceDomain) {
    $SourceDomain = Read-Host "Enter source domain name (e.g., source.local)"
}

Write-Log -Message "Starting WMI filter export from domain: $SourceDomain" -Level INFO

try {
    Invoke-Safely -ScriptBlock {
        # Get all WMI filters from source domain
        # WMI filters are stored in: CN=SOM,CN=WMIPolicy,CN=System,DC=...
        $rootDSE = Get-ADRootDSE -Server $SourceDomain
        $wmipolicyPath = "CN=WMIPolicy,CN=System,$($rootDSE.defaultNamingContext)"
        
        $WMIFilters = Get-ADObject -SearchBase $wmipolicyPath -Filter "objectClass -eq 'msWMISom'" `
            -Properties msWMI-Name, msWMI-Parm1, msWMI-Parm2, Created, Modified `
            -Server $SourceDomain
        
        $WMIFilterData = @()
        if ($WMIFilters.Count -eq 0) {
            Write-Log -Message "No WMI filters found in source domain" -Level WARN
            Write-Host "[-] No WMI filters found"
        } else {
            foreach ($filter in $WMIFilters) {
                $WMIFilterData += [PSCustomObject]@{
                    FilterName        = $filter.'msWMI-Name'
                    FilterDescription = $filter.'msWMI-Parm1'
                    Query             = $filter.'msWMI-Parm2'
                    DistinguishedName = $filter.DistinguishedName
                    Created           = $filter.Created
                    Modified          = $filter.Modified
                    GUID              = $filter.ObjectGUID
                }
            }
            Write-Host "WMI filter export complete: $($WMIFilters.Count) filters exported"
        }
        
        # Always create the export file, even if empty, for validation consistency.
        $outputFile = Join-Path $ExportPath "WMI_Filters_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $WMIFilterData | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Exported $($WMIFilterData.Count) WMI filters to $outputFile" -Level INFO

        # Also export the raw WMI query data for transform phase if any were found
        if ($WMIFilterData.Count -gt 0) {
            $queryFile = Join-Path $ExportPath "WMI_Filters_Queries_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
            $WMIFilterData | ForEach-Object {
                Add-Content -Path $queryFile -Value "=== $($_.FilterName) ===" -Encoding UTF8
                Add-Content -Path $queryFile -Value ($_.Query -join "`r`n") -Encoding UTF8
                Add-Content -Path $queryFile -Value "`n" -Encoding UTF8
            }
        }
    } -Operation "Export WMI filters from $SourceDomain"
    
} catch {
    Write-Log -Message "Failed to export WMI filters: $_" -Level ERROR
    throw "WMI filter export failed. Check logs for details."
}
