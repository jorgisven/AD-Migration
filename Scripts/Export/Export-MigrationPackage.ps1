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

    # 1. Copy Start-MigrationGUI.ps1 to the root of the staging area
    $GuiScriptPath = Join-Path $RepoRoot "Start-MigrationGUI.ps1"
    if (Test-Path $GuiScriptPath) {
        Write-Log -Message "Copying Start-MigrationGUI.ps1..." -Level INFO
        Copy-Item -Path $GuiScriptPath -Destination $TempDir
    } else {
        Write-Log -Message "Start-MigrationGUI.ps1 not found at $GuiScriptPath. It will be missing from the package." -Level WARN
    }

    # 2. Copy Export Data
    Write-Log -Message "Copying Export data from $ExportRoot..." -Level INFO
    Copy-Item -Path $ExportRoot -Destination $TempDir -Recurse -Container

    # 3. Copy Transform Data (If available)
    if (Test-Path $TransformRoot) {
        Write-Log -Message "Copying Transform data from $TransformRoot..." -Level INFO
        Copy-Item -Path $TransformRoot -Destination $TempDir -Recurse -Container
    }

    # 4. Copy Scripts (So you can run Import on the other side)
    $StageScripts = Join-Path $TempDir "Scripts"
    $SourceScripts = Join-Path $RepoRoot "Scripts"
    Write-Log -Message "Copying Scripts from $SourceScripts..." -Level INFO
    if (-not (Test-Path $StageScripts)) {
        New-Item -ItemType Directory -Path $StageScripts -Force | Out-Null
    }
    Copy-Item -Path (Join-Path $SourceScripts '*') -Destination $StageScripts -Recurse -Container

    # 4b. Write package metadata for verification on target host.
    $gpoImportScript = Join-Path $SourceScripts 'Import\Import-GPOs.ps1'
    $gpoImportVersion = ''
    $gpoImportHash = ''
    if (Test-Path $gpoImportScript) {
        $versionLine = Select-String -Path $gpoImportScript -Pattern "Import-GPOs script version" -SimpleMatch | Select-Object -First 1
        if ($versionLine) {
            $gpoImportVersion = [string]$versionLine.Line
        }
        try {
            $gpoImportHash = (Get-FileHash -LiteralPath $gpoImportScript -Algorithm SHA256 -ErrorAction Stop).Hash
        } catch {
            $gpoImportHash = ''
        }
    }

    $metadata = [ordered]@{
        PackageCreatedAt = (Get-Date).ToString('o')
        RepoRoot = $RepoRoot
        SourceScriptsPath = $SourceScripts
        ImportGPOsPath = $gpoImportScript
        ImportGPOsVersionLine = $gpoImportVersion
        ImportGPOsSHA256 = $gpoImportHash
    }
    $metadataPath = Join-Path $TempDir 'PACKAGE_METADATA.json'
    $metadata | ConvertTo-Json -Depth 4 | Set-Content -Path $metadataPath -Encoding UTF8

    # 5. Copy Documentation
    $DocsDir = Join-Path $RepoRoot "Docs"
    if (Test-Path $DocsDir) {
        Write-Log -Message "Copying Documentation from $DocsDir..." -Level INFO
        Copy-Item -Path $DocsDir -Destination $TempDir -Recurse -Container
    }
    
    $RootMdFiles = Get-ChildItem -Path $RepoRoot -Filter "*.md" -File
    if ($RootMdFiles) {
        Write-Log -Message "Copying root markdown files..." -Level INFO
        $RootMdFiles | Copy-Item -Destination $TempDir
    }

    # 6. Zip it up
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
3. Run `Start-MigrationGUI.ps1` from the root of the extracted folder.
4. Use the GUI buttons to proceed with the **Transform** or **Import** phase.
"@

    Write-Host "✅ Migration Package Created: $ZipPath" -ForegroundColor Green
    Write-Host $Instructions -ForegroundColor Cyan

    # GUI Pop-up
    [System.Windows.Forms.MessageBox]::Show($Instructions, "Migration Package Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

} catch {
    Write-Log -Message "Failed to create migration package: $_" -Level ERROR
    Write-Host "❌ Packaging failed. Check logs." -ForegroundColor Red
}
