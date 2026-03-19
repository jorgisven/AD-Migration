<#
.SYNOPSIS
    Import Group Policy Objects (GPOs) from backups into the target domain.

.DESCRIPTION
    Reads GPO backups from the Export directory and imports them into the current (Target) domain.
    Supports using a Migration Table to map security principals and UNC paths.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain,

    [Parameter(Mandatory = $false)]
    [string]$MigrationTablePath,

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

# Check for GroupPolicy module
if (-not (Get-Module -Name GroupPolicy -ListAvailable)) {
    throw "GroupPolicy module (GPMC) is required but not installed."
}

$config = Get-ADMigrationConfig
$BackupPath = Join-Path $config.ExportRoot 'GPO_Backups'

# Default migration table path if not provided
if (-not $MigrationTablePath) {
    # Look for a .migtable file in the Transform/Mapping folder
    $MapPath = Join-Path $config.TransformRoot 'Mapping'
    $foundTable = Get-ChildItem -Path $MapPath -Filter "*.migtable" | Select-Object -First 1
    if ($foundTable) {
        $MigrationTablePath = $foundTable.FullName
        Write-Log -Message "Auto-detected migration table: $MigrationTablePath" -Level INFO
    }
}

# Validate Migration Table if present
if ($MigrationTablePath) {
    if (-not (Test-Path $MigrationTablePath)) {
        throw "Migration Table file not found at '$MigrationTablePath'"
    }
    
    # Check for unmapped entries which cause Import-GPO to fail
    try {
        [xml]$migTableXml = Get-Content $MigrationTablePath
        $unmapped = $migTableXml.MigrationTable.Mapping | Where-Object { 
            ($_.Destination.Path -eq "" -or $_.Destination.Path -eq $null) 
        }
        if ($unmapped) {
            throw "The Migration Table contains $($unmapped.Count) unmapped entries (empty Destination). Please review and edit '$MigrationTablePath' before importing."
        }
        Write-Log -Message "Using validated Migration Table: $MigrationTablePath" -Level INFO
    } catch {
        throw "Migration Table Validation Failed: $_"
    }
} else {
    Write-Log -Message "No Migration Table specified or detected. GPOs will be imported with original security principals and paths." -Level WARN
}

if (-not (Test-Path $BackupPath)) {
    Write-Log -Message "GPO Backup directory not found at $BackupPath. Ensure Export-GPOs (or manual backup) was run." -Level ERROR
    throw "Missing GPO Backups"
}

if (-not $TargetDomain) { $TargetDomain = (Get-ADDomain).DNSRoot }

Write-Log -Message "Starting GPO Import to $TargetDomain..." -Level INFO

Invoke-Safely -ScriptBlock {
    # Get list of backups from the manifest
    $manifestPath = Join-Path $BackupPath "manifest.xml"
    if (-not (Test-Path $manifestPath)) { throw "Backup manifest.xml not found in $BackupPath" }
    
    [xml]$manifest = Get-Content $manifestPath
    $backups = $manifest.Backups.BackupInst
    
    foreach ($b in $backups) {
        $gpoName = $b.GPODisplayName
        $backupId = $b.ID
        
        # Check if GPO exists to ensure idempotency
        $existing = Get-GPO -Name $gpoName -Server $TargetDomain -ErrorAction SilentlyContinue
        if ($existing -and -not $Force) {
            Write-Log -Message "GPO '$gpoName' already exists. Skipping (Use -Force to overwrite)." -Level WARN
            continue
        }

        Write-Log -Message "Importing GPO: '$gpoName' (ID: $backupId)" -Level INFO
        
        $params = @{
            BackupId       = $backupId
            Path           = $BackupPath
            TargetName     = $gpoName
            Domain         = $TargetDomain
            CreateIfNeeded = $true
            ErrorAction    = 'Stop'
        }
        
        if ($MigrationTablePath) {
            $params.MigrationTable = $MigrationTablePath
        }
        
        try {
            if ($PSCmdlet.ShouldProcess($gpoName, "Import GPO")) {
                Import-GPO @params | Out-Null
                Write-Log -Message "Successfully imported '$gpoName'" -Level INFO
            }
        } catch {
            Write-Log -Message "Failed to import '$gpoName': $_" -Level ERROR
        }
    }
} -Operation "Import GPOs"