<#
.SYNOPSIS
    Package exported data and scripts for transfer to target environment.

.DESCRIPTION
    Creates a ZIP file containing the 'Export' data and the 'Scripts' folder.
    This allows you to easily move the migration state to a disconnected target network.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

# Add GUI support
Add-Type -AssemblyName System.Windows.Forms

$config = Get-ADMigrationConfig
$ExportRoot = $config.ExportRoot
$TransformRoot = $config.TransformRoot
$RepoRoot = Split-Path -Parent (Split-Path -Parent $ScriptRoot) # Root of the git repo

# Define Output Path (e.g. Desktop)
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$PackageName = "ADMigration_Package_$(Get-Date -Format 'yyyyMMdd_HHmmss').zip"
$ZipPath = Join-Path $DesktopPath $PackageName

Write-Log -Message "Starting Migration Package generation..." -Level INFO

try {
    # Create a temporary staging folder
    $TempDir = Join-Path $env:TEMP "ADMigration_Staging_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null

    # 1. Copy Export Data
    Write-Log -Message "Copying Export data from $ExportRoot..." -Level INFO
    Copy-Item -Path $ExportRoot -Destination $TempDir -Recurse -Container

    # 2. Copy Transform Data (If available)
    if (Test-Path $TransformRoot) {
        Write-Log -Message "Copying Transform data from $TransformRoot..." -Level INFO
        Copy-Item -Path $TransformRoot -Destination $TempDir -Recurse -Container
    }

    # 3. Copy Scripts (So you can run Import on the other side)
    $StageScripts = Join-Path $TempDir "Scripts"
    $SourceScripts = Join-Path $RepoRoot "Scripts"
    Write-Log -Message "Copying Scripts from $SourceScripts..." -Level INFO
    Copy-Item -Path $SourceScripts -Destination $StageScripts -Recurse -Container

    # 4. Zip it up
    Write-Log -Message "Compressing package to $ZipPath..." -Level INFO
    Compress-Archive -Path "$TempDir\*" -DestinationPath $ZipPath

    # Cleanup
    Remove-Item -Path $TempDir -Recurse -Force

    $Instructions = @"
SUCCESS: Migration Package Created
Location: $ZipPath

NEXT STEPS:
1. Transfer this ZIP file to your destination machine (Analyst Workstation or Target Domain).
2. Extract the archive.

IF ANALYZING (Analyst Workstation):
   - Run the interactive transform wizard (No Admin required):
     Scripts\Transform\Run-AllTransforms.ps1

IF IMPORTING (Target Domain):
   - Run the import orchestrator AS ADMINISTRATOR:
     Scripts\Import\Run-AllImports.ps1
"@

    Write-Host "✅ Migration Package Created: $ZipPath" -ForegroundColor Green
    Write-Host $Instructions -ForegroundColor Cyan

    # GUI Pop-up
    [System.Windows.Forms.MessageBox]::Show($Instructions, "Migration Package Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

} catch {
    Write-Log -Message "Failed to create migration package: $_" -Level ERROR
    Write-Host "❌ Packaging failed. Check logs." -ForegroundColor Red
}
