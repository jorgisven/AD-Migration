<#
.SYNOPSIS
    Main graphical user interface for the AD Migration framework.
.DESCRIPTION
    Provides a central launchpad to start the Export, Transform, or Import phases of the migration.
    This GUI guides the user to the correct orchestration script for each step.
#>

# Set working directory to the script's location to ensure relative paths work
try {
    $ScriptDir = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
    Set-Location -Path $ScriptDir
} catch {
    Write-Warning "Could not set working directory. Relative paths may fail."
}

# --- GUI Setup ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# --- Prerequisite Check: Active Directory Module ---
if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
    $msg = "The 'ActiveDirectory' PowerShell module is required but not found.`n`nPlease install the Remote Server Administration Tools (RSAT) for Active Directory.`n`nOn Windows 10/11, you can do this via:`nSettings > Apps > Optional features > Add a feature`n(Search for 'RSAT: Active Directory Domain Services...')`n`nOr by running this PowerShell command as Administrator:`nAdd-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'"
    [System.Windows.Forms.MessageBox]::Show($msg, "Prerequisite Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# --- Prerequisite Check: Group Policy Module ---
if (-not (Get-Module -Name GroupPolicy -ListAvailable)) {
    $msg = "The 'GroupPolicy' PowerShell module is required but not found.`n`nPlease install the RSAT Group Policy Management Tools.`n`nOn Windows 10/11, you can do this via:`nSettings > Apps > Optional features > Add a feature`n(Search for 'RSAT: Group Policy Management Tools')`n`nOr by running this PowerShell command as Administrator:`nAdd-WindowsCapability -Online -Name 'Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0'"
    [System.Windows.Forms.MessageBox]::Show($msg, "Prerequisite Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# --- Check for running from temp location ---
$tempPath = [System.IO.Path]::GetTempPath()
if ($ScriptDir.StartsWith($tempPath, [System.StringComparison]::OrdinalIgnoreCase)) {
    $msg = "It appears you are running this script from a temporary folder, likely inside a ZIP file.`n`nPlease extract the entire archive to a permanent location before running the migration assistant."
    [System.Windows.Forms.MessageBox]::Show($msg, "Run from Extracted Folder", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    exit
}


# --- Form Creation ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Migration Orchestrator"
$form.Size = New-Object System.Drawing.Size(400, 420)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# --- Controls ---

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "AD Migration Framework"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(20, 20)
$lblTitle.AutoSize = $true
$form.Controls.Add($lblTitle)

# --- Export Button ---
$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "1. Export from Source Domain"
$btnExport.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnExport.Location = New-Object System.Drawing.Point(40, 80)
$btnExport.Size = New-Object System.Drawing.Size(300, 50)
$form.Controls.Add($btnExport)

$lblExport = New-Object System.Windows.Forms.Label
$lblExport.Text = "Extracts OUs, GPOs, accounts, and more from the source domain into a portable package."
$lblExport.Location = New-Object System.Drawing.Point(40, 135)
$lblExport.Size = New-Object System.Drawing.Size(300, 40)
$form.Controls.Add($lblExport)

# --- Transform Button ---
$btnTransform = New-Object System.Windows.Forms.Button
$btnTransform.Text = "2. Transform Migration Data"
$btnTransform.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnTransform.Location = New-Object System.Drawing.Point(40, 185)
$btnTransform.Size = New-Object System.Drawing.Size(300, 50)
$form.Controls.Add($btnTransform)

$lblTransform = New-Object System.Windows.Forms.Label
$lblTransform.Text = "Interactive wizard to map OUs, rewrite GPOs, and prepare data for the target domain. (Offline)"
$lblTransform.Location = New-Object System.Drawing.Point(40, 240)
$lblTransform.Size = New-Object System.Drawing.Size(300, 40)
$form.Controls.Add($lblTransform)

# --- Import Button ---
$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Text = "3. Import to Target Domain"
$btnImport.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnImport.Location = New-Object System.Drawing.Point(40, 290)
$btnImport.Size = New-Object System.Drawing.Size(300, 50)
$form.Controls.Add($btnImport)

$lblImport = New-Object System.Windows.Forms.Label
$lblImport.Text = "Rebuilds the OU structure, GPOs, and other objects in the target domain. (Requires Admin)"
$lblImport.Location = New-Object System.Drawing.Point(40, 345)
$lblImport.Size = New-Object System.Drawing.Size(300, 40)
$form.Controls.Add($lblImport)


# --- Event Handlers ---

$btnExport.Add_Click({
    $form.Hide()
    try {
        $sourceDomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the FQDN of the SOURCE domain:", "Export Phase", "source.local")
        if ([string]::IsNullOrWhiteSpace($sourceDomain)) {
            [System.Windows.Forms.MessageBox]::Show("Export cancelled. Source domain cannot be empty.", "Cancelled", "OK", "Warning")
            return
        }
        
        $exportScript = ".\Scripts\Export\Run-AllExports.ps1"
        if (-not (Test-Path $exportScript)) { throw "Export script not found at $exportScript" }
        
        # Launch in a new window to show progress
        Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "& '$exportScript' -SourceDomain '$sourceDomain'; Write-Host 'Press any key to exit...'; [System.Console]::ReadKey() | Out-Null"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error", "OK", "Error")
    } finally {
        $form.Show()
    }
})

$btnTransform.Add_Click({
    $form.Hide()
    try {
        $transformScript = ".\Scripts\Transform\Run-AllTransforms.ps1"
        if (-not (Test-Path $transformScript)) { throw "Transform script not found at $transformScript" }
        
        Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "& '$transformScript'; Write-Host 'Press any key to exit...'; [System.Console]::ReadKey() | Out-Null"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error", "OK", "Error")
    } finally {
        $form.Show()
    }
})

$btnImport.Add_Click({
    $form.Hide()
    try {
        $importScript = ".\Scripts\Import\Run-AllImports.ps1"
        if (-not (Test-Path $importScript)) { throw "Import orchestrator script not found at $importScript." }
        
        Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", "& '$importScript'; Write-Host 'Press any key to exit...'; [System.Console]::ReadKey() | Out-Null"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $_", "Error", "OK", "Error")
    } finally {
        $form.Show()
    }
})


# --- Show Form ---
$form.ShowDialog() | Out-Null