function Get-ADMigrationConfig {
    [CmdletBinding()]
    param()

    # Use local Documents folder, not OneDrive (sensitive data stays local)
    $documents = Join-Path $env:USERPROFILE 'Documents'
    $root      = Join-Path $documents 'ADMigration'

    return @{
        Root          = $root
        LogRoot       = Join-Path $root 'Logs'
        ExportRoot    = Join-Path $root 'Export'
        TransformRoot = Join-Path $root 'Transform'
        ImportRoot    = Join-Path $root 'Import'
    }
}
