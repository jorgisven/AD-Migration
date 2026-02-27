<#
.SYNOPSIS
    Restore GPOs into target domain.

.DESCRIPTION
    Restores GPO backups, applies rewrites, and validates settings.
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
$BackupPath = Join-Path $config.ExportRoot 'GPO_Backups'
$MigTablePath = Join-Path $config.TransformRoot 'GPO_MigrationTable.migtable'

if (-not (Test-Path $BackupPath)) {
    Write-Log -Message "No GPO backups found at $BackupPath. Run Export-GPOReports.ps1 first." -Level ERROR
    throw "GPO Backups missing"
}

if (-not $TargetDomain) {
    $TargetDomain = (Get-ADDomain).DNSRoot
}

Write-Log -Message "Starting GPO Import to $TargetDomain..." -Level INFO

if (Test-Path $MigTablePath) {
    Write-Log -Message "Using Migration Table: $MigTablePath" -Level INFO
}

Invoke-Safely -ScriptBlock {
    # Get all backups from the folder
    # Import-GPO requires the backup directory, not individual files
    $Backups = Get-GPO -Path $BackupPath -All
    
    Write-Log -Message "Found $($Backups.Count) GPO backups to import" -Level INFO

    foreach ($backup in $Backups) {
        $gpoName = $backup.DisplayName
        
        try {
            # Check if GPO exists to avoid errors or decide on overwrite strategy
            # Here we use CreateIfNeeded which imports settings into a new or existing GPO
            $params = @{
                BackupId       = $backup.Id
                Path           = $BackupPath
                TargetName     = $gpoName
                CreateIfNeeded = $true
                Domain         = $TargetDomain
                ErrorAction    = 'Stop'
            }

            if (Test-Path $MigTablePath) {
                $params['MigrationTable'] = $MigTablePath
            }

            Import-GPO @params | Out-Null
            
            Write-Log -Message "Imported GPO: $gpoName" -Level INFO
        } catch {
            Write-Log -Message "Failed to import GPO '$gpoName': $_" -Level ERROR
        }
    }
    
} -Operation "Import GPOs to $TargetDomain"