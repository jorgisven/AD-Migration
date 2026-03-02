function Initialize-ADMigration {
    [CmdletBinding()]
    param()

    $config = Get-ADMigrationConfig
    # create all configured directories if they don't exist
    foreach ($path in $config.Values) {
        if (-not (Test-Path $path)) {
            New-Item -Path $path -ItemType Directory -Force | Out-Null
        }
    }

    Write-Log -Message "ADMigration directories initialized under $($config.Root)" -Level INFO
}
