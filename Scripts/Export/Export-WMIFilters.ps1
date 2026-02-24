<#
.SYNOPSIS
    Export WMI filters from source domain.

.DESCRIPTION
    Captures WMI filter names, queries, and descriptions for recreation in target domain.
#>

$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'WMI_Filters'

# ensure export folder exists
if (-not (Test-Path $ExportPath)) {
    New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
}

Invoke-Safely -Operation 'Export WMI filters' -ScriptBlock {
    $filters = Get-WmiFilter -All
    $out = $filters | Select-Object Name,Description,Query,Id
    $out | Export-Csv -Path (Join-Path $ExportPath 'WmiFilters.csv') -NoTypeInformation -Force
}

Write-Log -Message 'WMI filters exported' -Level INFO
