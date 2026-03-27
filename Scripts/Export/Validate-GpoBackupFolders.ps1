<#
.SYNOPSIS
    Export-side entry point for GPO backup validation.

.DESCRIPTION
    This script forwards execution to Scripts/Import/Validate-GpoBackupFolders.ps1
    so operators can run validation from either workflow folder.

.PARAMETER BackupPath
    Optional path to the GPO_Backups folder to validate.

.EXAMPLE
    .\Validate-GpoBackupFolders.ps1 -BackupPath "C:\Users\Administrator\Documents\ADMigration\Export\GPO_Backups"
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$BackupPath
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$importScript = Join-Path (Split-Path -Parent $scriptRoot) 'Import\Validate-GpoBackupFolders.ps1'

if (-not (Test-Path $importScript)) {
    Write-Error "Could not locate import validator at: $importScript"
    exit 1
}

$invokeParams = @{}
if ($PSBoundParameters.ContainsKey('BackupPath')) {
    $invokeParams.BackupPath = $BackupPath
}

& $importScript @invokeParams
if ($LASTEXITCODE -ne $null) {
    exit $LASTEXITCODE
}
