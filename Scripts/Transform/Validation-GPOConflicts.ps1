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

$runState = 'Skipped'
$runReason = 'Not started.'
$exitCode = $null

Write-Log -Message "Starting optional GPO conflict review launcher (Policy Analyzer)." -Level INFO

# This script is a launcher/wrapper for Microsoft Policy Analyzer.
# It prompts for PolicyAnalyzer.exe and opens the exported GPO backup directory.
# Ask the user where Policy Analyzer is located
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "Select PolicyAnalyzer.exe to run GPO conflict review (Download from https://aka.ms/sct)"
$dialog.Filter = "PolicyAnalyzer.exe|PolicyAnalyzer.exe"

$dialogResult = $dialog.ShowDialog()
if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $analyzerExe = $dialog.FileName

    if (-not (Test-Path $analyzerExe)) {
        $runState = 'Failed'
        $runReason = "Selected PolicyAnalyzer.exe path does not exist: $analyzerExe"
        Write-Host "GPO conflict review failed: selected PolicyAnalyzer.exe path was not found." -ForegroundColor Red
        Write-Log -Message $runReason -Level ERROR
    } else {
        $runState = 'Attempted'
        $runReason = "Launching Policy Analyzer from '$analyzerExe' with backup path '$BackupPath'."
        Write-Log -Message "Policy Analyzer run state: $runState. $runReason" -Level INFO

        Write-Host "Launching Microsoft Policy Analyzer with exported GPOs..." -ForegroundColor Cyan
        Write-Host "Look for settings highlighted in YELLOW to identify conflicts!" -ForegroundColor Yellow

        try {
            # Passing the Backup directory as an argument automatically loads all GPOs into the tool
            $process = Start-Process -FilePath $analyzerExe -ArgumentList "`"$BackupPath`"" -Wait -PassThru -ErrorAction Stop
            $exitCode = $process.ExitCode
            if ($exitCode -eq 0) {
                $runState = 'Succeeded'
                $runReason = "Policy Analyzer exited normally (ExitCode=0)."
            } else {
                $runState = 'Failed'
                $runReason = "Policy Analyzer exited with non-zero code $exitCode."
            }
        } catch {
            $runState = 'Failed'
            $runReason = "Policy Analyzer launch failed: $($_.Exception.Message)"
        }
    }
} else {
    Write-Host "GPO conflict review skipped (Policy Analyzer was not selected)." -ForegroundColor DarkGray
    $runState = 'Skipped'
    $runReason = 'User did not select PolicyAnalyzer.exe.'
}

$finalMessage = if ($null -ne $exitCode) {
    "Policy Analyzer run state: $runState. Reason: $runReason ExitCode=$exitCode"
} else {
    "Policy Analyzer run state: $runState. Reason: $runReason"
}

if ($runState -eq 'Failed') {
    Write-Log -Message $finalMessage -Level WARN
} else {
    Write-Log -Message $finalMessage -Level INFO
}