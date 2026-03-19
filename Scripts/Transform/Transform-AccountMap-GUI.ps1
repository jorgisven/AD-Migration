<#
.SYNOPSIS
    GUI Tool for editing Account Mapping CSVs.

.DESCRIPTION
    Provides a tabbed DataGridView interface to review and modify User, Computer, 
    and Group account mappings without requiring Excel.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$MapPath = Join-Path $config.TransformRoot 'Mapping'

# Load Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

# --- GUI SETUP ---

$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Migration - Account Mapper"
$form.Size = New-Object System.Drawing.Size(1100, 600)
$form.StartPosition = "CenterScreen"

$tabCtrl = New-Object System.Windows.Forms.TabControl
$tabCtrl.Dock = "Fill"
$form.Controls.Add($tabCtrl)

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

$btnSearchClear = New-Object System.Windows.Forms.Button
$btnSearchClear.Text = "Clear"
$btnSearchClear.Location = New-Object System.Drawing.Point(370, 9)
$btnSearchClear.Width = 60
$panelTop.Controls.Add($btnSearchClear)

# Bottom Panel (Buttons)
$panelBottom = New-Object System.Windows.Forms.Panel
$panelBottom.Height = 50
$panelBottom.Dock = "Bottom"
$form.Controls.Add($panelBottom)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save All Maps"
$btnSave.Width = 150
$btnSave.Location = New-Object System.Drawing.Point(910, 10)
$btnSave.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
$panelBottom.Controls.Add($btnSave)

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Set Action to 'Create' for new objects, 'Merge' to link to existing targets, or 'Skip' to ignore."
$lblInfo.AutoSize = $true
$lblInfo.Location = New-Object System.Drawing.Point(10, 15)
$panelBottom.Controls.Add($lblInfo)

$tabCtrl.BringToFront()

# --- FUNCTIONS ---

$script:grids = @{}

function Load-CsvToGrid ($fileName, $tabName) {
    $csvPath = Join-Path $MapPath $fileName
    if (-not (Test-Path $csvPath)) { return }

    $data = Import-Csv $csvPath
    if (-not $data) { return }

    # Create DataTable for DataGridView binding
    $dt = New-Object System.Data.DataTable
    $props = $data[0].psobject.properties.name
    
    foreach ($prop in $props) {
        $dt.Columns.Add($prop) | Out-Null
    }
    
    foreach ($row in $data) {
        $dr = $dt.NewRow()
        foreach ($prop in $props) {
            $dr[$prop] = $row.$prop
        }
        $dt.Rows.Add($dr)
    }

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = "$tabName ($($dt.Rows.Count))"
    $tab.Tag = $tabName

    $dgv = New-Object System.Windows.Forms.DataGridView
    $dgv.Dock = "Fill"
    $dgv.DataSource = $dt
    $dgv.AllowUserToAddRows = $false # Force users to use "Skip" instead of deleting rows
    $dgv.AllowUserToDeleteRows = $false
    $dgv.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $tab.Controls.Add($dgv)
    
    # Replace 'Action' column with a ComboBox
    if ($dgv.Columns.Contains("Action")) {
        $colIndex = $dgv.Columns["Action"].Index
        $dgv.Columns.Remove("Action")
        
        $cmbCol = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
        $cmbCol.Name = "Action"
        $cmbCol.DataPropertyName = "Action"
        $cmbCol.HeaderText = "Action"
        $cmbCol.Items.AddRange("Create", "Merge", "Skip")
        $cmbCol.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
        $dgv.Columns.Insert($colIndex, $cmbCol)
    }

    # Lock source columns
    foreach ($col in $dgv.Columns) {
        if ($col.Name -match "^Source") {
            $col.ReadOnly = $true
            $col.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
            $col.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGray
        }
    }
    
    # Ensure TargetOU_DN takes up remaining space
    if ($dgv.Columns.Contains("TargetOU_DN")) {
        $dgv.Columns["TargetOU_DN"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    }

    $tabCtrl.TabPages.Add($tab)
    $script:grids[$csvPath] = $dgv
}

$btnSave.Add_Click({
    $dgvCount = 0
    foreach ($key in $script:grids.Keys) {
        $dgv = $script:grids[$key]
        $dt = $dgv.DataSource
        
        $exportData = @()
        foreach ($row in $dt.Rows) {
            if ($row.RowState -eq [System.Data.DataRowState]::Deleted) { continue }
            $obj = New-Object PSObject
            foreach ($col in $dt.Columns) {
                $obj | Add-Member -MemberType NoteProperty -Name $col.ColumnName -Value $row[$col.ColumnName]
            }
            $exportData += $obj
        }
        
        $exportData | Export-Csv -Path $key -NoTypeInformation -Encoding UTF8
        $dgvCount++
    }
    
    [System.Windows.Forms.MessageBox]::Show("Successfully saved $dgvCount mapping files.", "Save Complete", "OK", "Information")
    $form.Close()
})

# Load the mappings
Load-CsvToGrid "User_Account_Map.csv" "Users"
Load-CsvToGrid "Computer_Account_Map.csv" "Computers"
Load-CsvToGrid "Group_Account_Map.csv" "Groups"

if ($tabCtrl.TabPages.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("No account mapping files found. Run the Transform-AccountMapping step first.", "Error", "OK", "Error")
    exit
}

# --- SEARCH LOGIC ---
$txtSearch.Add_TextChanged({
    # Escape characters that have special meaning in DataView RowFilters
    $searchText = $txtSearch.Text.Replace("'", "''").Replace("[", "[[]").Replace("*", "[*]").Replace("%", "[%]")
    
    foreach ($key in $script:grids.Keys) {
        $dgv = $script:grids[$key]
        $dt = $dgv.DataSource
        $tab = $dgv.Parent
        $baseName = $tab.Tag
        $totalRows = $dt.Rows.Count
        
        if ([string]::IsNullOrWhiteSpace($searchText)) {
            $dt.DefaultView.RowFilter = ""
            $tab.Text = "$baseName ($totalRows)"
            continue
        }

        $filterParts = @()
        foreach ($col in $dt.Columns) {
            $filterParts += "[$($col.ColumnName)] LIKE '*$searchText*'"
        }
        $dt.DefaultView.RowFilter = $filterParts -join " OR "
        
        $filteredRows = $dt.DefaultView.Count
        $tab.Text = "$baseName ($filteredRows / $totalRows)"
    }
})

$btnSearchClear.Add_Click({ $txtSearch.Text = "" })

# --- SHOW ---
$form.ShowDialog() | Out-Null