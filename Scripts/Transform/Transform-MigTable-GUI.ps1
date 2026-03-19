<#
.SYNOPSIS
    GUI Tool for editing the GPO Migration Table (.migtable).

.DESCRIPTION
    Provides a DataGridView interface to review and modify Security Principal 
    and UNC Path mappings for GPO transformations.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$MapPath = Join-Path $config.TransformRoot 'Mapping'
$MigTablePath = Join-Path $MapPath "MigrationTable.migtable"

if (-not (Test-Path $MigTablePath)) {
    [System.Windows.Forms.MessageBox]::Show("Migration Table not found. Run the Transform-GenerateMigrationTable step first.", "Error", "OK", "Error")
    exit
}

# Load Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
Add-Type -AssemblyName System.Xml

[xml]$script:migXml = Get-Content $MigTablePath

# --- GUI SETUP ---

$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Migration - Migration Table Editor"
$form.Size = New-Object System.Drawing.Size(1000, 600)
$form.StartPosition = "CenterScreen"

# Top Panel (Search)
$panelTop = New-Object System.Windows.Forms.Panel
$panelTop.Height = 40
$panelTop.Dock = "Top"
$form.Controls.Add($panelTop)

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = "Search:"
$lblSearch.AutoSize = $true
$lblSearch.Location = New-Object System.Drawing.Point(10, 12)
$panelTop.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(60, 10)
$txtSearch.Width = 300
$panelTop.Controls.Add($txtSearch)

# Bottom Panel (Buttons)
$panelBottom = New-Object System.Windows.Forms.Panel
$panelBottom.Height = 50
$panelBottom.Dock = "Bottom"
$form.Controls.Add($panelBottom)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save Migration Table"
$btnSave.Width = 150
$btnSave.Location = New-Object System.Drawing.Point(810, 10)
$btnSave.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$panelBottom.Controls.Add($btnSave)

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Fill in the 'Destination' column for any blank entries. These map source items to your new target domain."
$lblInfo.AutoSize = $true
$lblInfo.Location = New-Object System.Drawing.Point(10, 15)
$panelBottom.Controls.Add($lblInfo)

# DataGrid Setup
$dt = New-Object System.Data.DataTable
$dt.Columns.Add("Type") | Out-Null
$dt.Columns.Add("Source") | Out-Null
$dt.Columns.Add("Destination") | Out-Null

$mappings = $script:migXml.GetElementsByTagName("Mapping")
foreach ($m in $mappings) {
    $dr = $dt.NewRow()
    $sourceType = $m.Source.GetAttribute("xsi:type")
    if (-not $sourceType) { $sourceType = $m.Source.GetAttribute("type", "http://www.w3.org/2001/XMLSchema-instance") }
    
    if ($sourceType -eq "GPMTrustee") {
        $dr["Type"] = "Security Principal"
        $dr["Source"] = $m.Source.GetAttribute("Name")
        $dr["Destination"] = $m.Destination.GetAttribute("Name")
    } else {
        $dr["Type"] = "UNC Path"
        $dr["Source"] = $m.Source.GetAttribute("Path")
        $dr["Destination"] = $m.Destination.GetAttribute("Path")
    }
    $dt.Rows.Add($dr)
}

$dgv = New-Object System.Windows.Forms.DataGridView
$dgv.Dock = "Fill"
$dgv.DataSource = $dt
$dgv.AllowUserToAddRows = $false
$dgv.AllowUserToDeleteRows = $false
$dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$form.Controls.Add($dgv)
$dgv.BringToFront()

# Lock non-editable columns
$dgv.Columns["Type"].ReadOnly = $true
$dgv.Columns["Type"].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$dgv.Columns["Type"].FillWeight = 20

$dgv.Columns["Source"].ReadOnly = $true
$dgv.Columns["Source"].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
$dgv.Columns["Source"].FillWeight = 40

$dgv.Columns["Destination"].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
$dgv.Columns["Destination"].FillWeight = 40

# Events
$txtSearch.Add_TextChanged({
    $searchText = $txtSearch.Text.Replace("'", "''").Replace("[", "[[]").Replace("*", "[*]").Replace("%", "[%]")
    if ([string]::IsNullOrWhiteSpace($searchText)) { $dt.DefaultView.RowFilter = "" } 
    else { $dt.DefaultView.RowFilter = "Type LIKE '*$searchText*' OR Source LIKE '*$searchText*' OR Destination LIKE '*$searchText*'" }
})

$btnSave.Add_Click({
    foreach ($row in $dt.Rows) {
        foreach ($m in $mappings) {
            $mType = if ($m.Source.GetAttribute("xsi:type")) { $m.Source.GetAttribute("xsi:type") } else { $m.Source.GetAttribute("type", "http://www.w3.org/2001/XMLSchema-instance") }
            if ($row["Type"] -eq "Security Principal" -and $mType -eq "GPMTrustee" -and $m.Source.GetAttribute("Name") -eq $row["Source"]) { $m.Destination.SetAttribute("Name", $row["Destination"]); break }
            if ($row["Type"] -eq "UNC Path" -and $mType -eq "GPMPath" -and $m.Source.GetAttribute("Path") -eq $row["Source"]) { $m.Destination.SetAttribute("Path", $row["Destination"]); break }
        }
    }
    $script:migXml.Save($MigTablePath)
    [System.Windows.Forms.MessageBox]::Show("Migration Table saved successfully.", "Save Complete", "OK", "Information")
    $form.Close()
})

$form.ShowDialog() | Out-Null