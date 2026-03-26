<#
.SYNOPSIS
    Automates the ingestion of exported GPOs into Microsoft Policy Analyzer.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$BackupPath = Join-Path $config.ExportRoot 'GPO_Backups'

if (-not (Test-Path $BackupPath)) {
    throw "GPO Backup directory not found. Please run the Export phase first."
}

# This script is a launcher/wrapper for Microsoft Policy Analyzer.
# It prompts for PolicyAnalyzer.exe and opens the exported GPO backup directory.
# Ask the user where Policy Analyzer is located
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "Select PolicyAnalyzer.exe to run GPO conflict review (Download from https://aka.ms/sct)"
$dialog.Filter = "PolicyAnalyzer.exe|PolicyAnalyzer.exe"

if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $analyzerExe = $dialog.FileName
    
    Write-Host "Launching Microsoft Policy Analyzer with exported GPOs..." -ForegroundColor Cyan
    Write-Host "Look for settings highlighted in YELLOW to identify conflicts!" -ForegroundColor Yellow
    
    # Passing the Backup directory as an argument automatically loads all GPOs into the tool
    Start-Process -FilePath $analyzerExe -ArgumentList "`"$BackupPath`"" -Wait
} else {
    Write-Host "GPO conflict review skipped (Policy Analyzer was not selected)." -ForegroundColor DarkGray
}