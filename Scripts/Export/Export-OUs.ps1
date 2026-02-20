<#
.SYNOPSIS
    Export OU structure from saicprint.local.

.DESCRIPTION
    Retrieves the full OU hierarchy and exports it to CSV for mapping and reconstruction.
#>

$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'OU_Structure'

# TODO: Add export logic