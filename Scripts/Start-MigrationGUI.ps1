<#
.SYNOPSIS
    GUI Front-end for AD Migration.

.DESCRIPTION
    Launches a Windows Forms interface to orchestrate the export and transform phases.
    Prompts for Source/Target domains, runs exports, generates mappings, and opens the output folder.
#>

# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Set up paths
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path $ScriptDir 'ADMigration') 'ADMigration.psd1'

# Load Module
if (Test-Path $ModulePath) {
    Import-Module $ModulePath -Force
} else {
    [System.Windows.Forms.MessageBox]::Show("ADMigration module not found at $ModulePath", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

$config = Get-ADMigrationConfig

# Create Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Migration Assistant"
$form.Size = New-Object System.Drawing.Size(500, 380)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $true

# Header
$lblHeader = New-Object System.Windows.Forms.Label
$lblHeader.Location = New-Object System.Drawing.Point(20, 15)
$lblHeader.Size = New-Object System.Drawing.Size(450, 30)
$lblHeader.Text = "Export & Transform Wizard"
$lblHeader.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblHeader)

# Source Domain Group
$grpSource = New-Object System.Windows.Forms.GroupBox
$grpSource.Location = New-Object System.Drawing.Point(20, 50)
$grpSource.Size = New-Object System.Drawing.Size(440, 80)
$grpSource.Text = "Source Domain"
$form.Controls.Add($grpSource)

$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Location = New-Object System.Drawing.Point(15, 25)
$lblSource.Size = New-Object System.Drawing.Size(400, 20)
$lblSource.Text = "FQDN (e.g. source.local):"
$grpSource.Controls.Add($lblSource)

$txtSource = New-Object System.Windows.Forms.TextBox
$txtSource.Location = New-Object System.Drawing.Point(15, 45)
$txtSource.Size = New-Object System.Drawing.Size(410, 20)
if ($env:USERDNSDOMAIN) { $txtSource.Text = $env:USERDNSDOMAIN }
$grpSource.Controls.Add($txtSource)

# Target Domain Group
$grpTarget = New-Object System.Windows.Forms.GroupBox
$grpTarget.Location = New-Object System.Drawing.Point(20, 140)
$grpTarget.Size = New-Object System.Drawing.Size(440, 80)
$grpTarget.Text = "Target Domain"
$form.Controls.Add($grpTarget)

$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Location = New-Object System.Drawing.Point(15, 25)
$lblTarget.Size = New-Object System.Drawing.Size(400, 20)
$lblTarget.Text = "FQDN (e.g. target.local):"
$grpTarget.Controls.Add($lblTarget)

$txtTarget = New-Object System.Windows.Forms.TextBox
$txtTarget.Location = New-Object System.Drawing.Point(15, 45)
$txtTarget.Size = New-Object System.Drawing.Size(410, 20)
$grpTarget.Controls.Add($txtTarget)

# Action Buttons
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Location = New-Object System.Drawing.Point(20, 240)
$btnRun.Size = New-Object System.Drawing.Size(150, 40)
$btnRun.Text = "Run Export && Transform"
$btnRun.BackColor = [System.Drawing.Color]::LightSkyBlue
$form.Controls.Add($btnRun)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Location = New-Object System.Drawing.Point(180, 240)
$btnOpen.Size = New-Object System.Drawing.Size(150, 40)
$btnOpen.Text = "Open Export Folder"
$form.Controls.Add($btnOpen)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Point(340, 240)
$btnClose.Size = New-Object System.Drawing.Size(120, 40)
$btnClose.Text = "Close"
$form.Controls.Add($btnClose)

# Status Bar
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Location = New-Object System.Drawing.Point(20, 300)
$lblStatus.Size = New-Object System.Drawing.Size(440, 30)
$lblStatus.Text = "Ready"
$lblStatus.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblStatus)

# Logic
$btnOpen.Add_Click({
    if (Test-Path $config.ExportRoot) {
        Invoke-Item $config.ExportRoot
    } else {
        [System.Windows.Forms.MessageBox]::Show("Export directory does not exist yet.", "Info", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

$btnClose.Add_Click({ $form.Close() })

$btnRun.Add_Click({
    $source = $txtSource.Text.Trim()
    $target = $txtTarget.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($source)) {
        [System.Windows.Forms.MessageBox]::Show("Source Domain is required.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    if ([string]::IsNullOrWhiteSpace($target)) {
        [System.Windows.Forms.MessageBox]::Show("Target Domain is required for transformation.", "Validation", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $btnRun.Enabled = $false
    $lblStatus.Text = "Running exports... (Check Console)"
    $form.Refresh()

    try {
        # 1. Run Exports
        $exportScript = Join-Path $ScriptDir "Export\Run-AllExports.ps1"
        if (Test-Path $exportScript) {
            & $exportScript -SourceDomain $source
        } else {
            throw "Export script missing: $exportScript"
        }

        $lblStatus.Text = "Running transforms..."
        $form.Refresh()

        # 2. Run Transforms
        # Migration Table
        $migTableScript = Join-Path $ScriptDir "Transform\Transform-GenerateMigrationTable.ps1"
        if (Test-Path $migTableScript) {
            & $migTableScript -SourceDomain $source -TargetDomain $target
        }

        # OU Map
        # Convert FQDN to DN (e.g. target.local -> DC=target,DC=local)
        $targetDN = "DC=" + ($target -replace '\.', ',DC=')
        $ouMapScript = Join-Path $ScriptDir "Transform\Transform-OUMap.ps1"
        if (Test-Path $ouMapScript) {
            & $ouMapScript -TargetBaseDN $targetDN
        }

        $lblStatus.Text = "Completed."
        [System.Windows.Forms.MessageBox]::Show("Process Complete!`n`nExports saved to: $($config.ExportRoot)`nTransforms saved to: $($config.TransformRoot)", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Auto-open folder as requested
        Invoke-Item $config.ExportRoot

    } catch {
        $lblStatus.Text = "Error occurred."
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $btnRun.Enabled = $true
    }
})

# Show
$form.ShowDialog() | Out-Null
$form.Dispose()