<#
.SYNOPSIS
    Export user and service account attributes.

.DESCRIPTION
    Used for account reconciliation between saicprint.local and crit.ad.
#>

$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'Security'

# TODO: Add account export logic