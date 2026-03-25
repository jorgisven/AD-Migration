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

$btnSysvol = New-Object System.Windows.Forms.Button
$btnSysvol.Text = "Auto-Map SYSVOL"
$btnSysvol.Width = 130
$btnSysvol.Location = New-Object System.Drawing.Point(670, 10)
$btnSysvol.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$panelBottom.Controls.Add($btnSysvol)

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

# Highlight SYSVOL paths for visibility
# Highlight SYSVOL paths for visibility and empty Destinations for validation errors
$dgv.Add_CellFormatting({
    param($sender, $e)
    if ($e.RowIndex -lt 0) { return }

    if ($sender.Columns[$e.ColumnIndex].Name -eq "Source") {
        $typeVal = $sender.Rows[$e.RowIndex].Cells["Type"].Value
        $srcVal = $e.Value
        if ($null -ne $srcVal -and $typeVal -eq "UNC Path" -and $srcVal -match "(?i)\\sysvol") {
            $e.CellStyle.BackColor = [System.Drawing.Color]::LightCyan
            $e.CellStyle.ForeColor = [System.Drawing.Color]::DarkBlue
        }
    }
    if ($sender.Columns[$e.ColumnIndex].Name -eq "Destination") {
        if ([string]::IsNullOrWhiteSpace($e.Value)) {
            $e.CellStyle.BackColor = [System.Drawing.Color]::Salmon
        }
    }
})

# Lock non-editable columns
if ($dgv.Columns.Contains("Type")) {
    $dgv.Columns["Type"].ReadOnly = $true
    $dgv.Columns["Type"].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
    $dgv.Columns["Type"].FillWeight = 20
}

if ($dgv.Columns.Contains("Source")) {
    $dgv.Columns["Source"].ReadOnly = $true
    $dgv.Columns["Source"].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
    $dgv.Columns["Source"].FillWeight = 40
}

if ($dgv.Columns.Contains("Destination")) {
    $dgv.Columns["Destination"].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
    $dgv.Columns["Destination"].FillWeight = 40
}

# Events
$txtSearch.Add_TextChanged({
    $searchText = $txtSearch.Text.Replace("'", "''").Replace("[", "[[]").Replace("*", "[*]").Replace("%", "[%]")
    if ([string]::IsNullOrWhiteSpace($searchText)) { $dt.DefaultView.RowFilter = "" } 
    else { $dt.DefaultView.RowFilter = "Type LIKE '*$searchText*' OR Source LIKE '*$searchText*' OR Destination LIKE '*$searchText*'" }
})

$btnSysvol.Add_Click({
    $targetDomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the FQDN of the TARGET domain (e.g., target.local) to automatically translate SYSVOL paths:", "Auto-Map SYSVOL", "")
    if ([string]::IsNullOrWhiteSpace($targetDomain)) { return }

    $updated = 0
    foreach ($row in $dt.Rows) {
        if ($row["Type"] -eq "UNC Path" -and $row["Source"] -match "(?i)\\sysvol") {
            $src = $row["Source"]
            $parts = $src -split "\\"
            # Ensure it is a valid UNC path containing SYSVOL in the share position
            if ($parts.Count -ge 4 -and $parts[3] -match "(?i)sysvol") {
                $oldDomain = $parts[2]
                $newDest = $src -ireplace [regex]::Escape($oldDomain), $targetDomain
                $row["Destination"] = $newDest
                $updated++
            }
        }
    }
    [System.Windows.Forms.MessageBox]::Show("Successfully updated $updated SYSVOL paths to point to '$targetDomain'.", "Auto-Map Complete", "OK", "Information") | Out-Null
})

$btnSave.Add_Click({
    $dgv.EndEdit()
    foreach ($row in $dt.Rows) {
        $destVal = if ($row["Destination"] -is [System.DBNull]) { "" } else { [string]$row["Destination"] }
        foreach ($m in $mappings) {
            $mType = if ($m.Source.GetAttribute("xsi:type")) { $m.Source.GetAttribute("xsi:type") } else { $m.Source.GetAttribute("type", "http://www.w3.org/2001/XMLSchema-instance") }
            if ($row["Type"] -eq "Security Principal" -and $mType -eq "GPMTrustee" -and $m.Source.GetAttribute("Name") -eq $row["Source"]) { $m.Destination.SetAttribute("Name", $row["Destination"]); break }
            if ($row["Type"] -eq "UNC Path" -and $mType -eq "GPMPath" -and $m.Source.GetAttribute("Path") -eq $row["Source"]) { $m.Destination.SetAttribute("Path", $row["Destination"]); break }
            if ($row["Type"] -eq "Security Principal" -and $mType -eq "GPMTrustee" -and $m.Source.GetAttribute("Name") -eq $row["Source"]) { $m.Destination.SetAttribute("Name", $destVal); break }
            if ($row["Type"] -eq "UNC Path" -and $mType -eq "GPMPath" -and $m.Source.GetAttribute("Path") -eq $row["Source"]) { $m.Destination.SetAttribute("Path", $destVal); break }
        }
    }
    $script:migXml.Save($MigTablePath)
    [System.Windows.Forms.MessageBox]::Show("Migration Table saved successfully.", "Save Complete", "OK", "Information")
    $form.Close()
})

$form.ShowDialog() | Out-Null