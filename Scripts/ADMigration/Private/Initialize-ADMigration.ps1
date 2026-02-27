function Initialize-ADMigration {
    $config = Get-ADMigrationConfig

    foreach ($path in @(
        $config.LogRoot,
        $config.ExportRoot,
        $config.TransformRoot,
        $config.ImportRoot
    )) {
        if (-not (Test-Path $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }

    Write-Log -Message "ADMigration module initialized"
}