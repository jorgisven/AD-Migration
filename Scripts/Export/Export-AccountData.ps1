<#
.SYNOPSIS
    Export user and service account attributes.

.DESCRIPTION
    Used for account reconciliation between source and target domains.
#>

$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'Security'

# TODO: Add account export logic