<#
.SYNOPSIS
    GUI Tool for editing Account Mapping CSVs.

.DESCRIPTION
    Provides a tabbed DataGridView interface to review and modify User, Computer, 
    and Group account mappings without requiring Excel.
#>

# PSScriptAnalyzer false-positive suppression for WinForms event scriptblocks.
#pragma warning disable PSAvoidAssignmentToAutomaticVariable

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$MapPath = Join-Path $config.TransformRoot 'Mapping'
$SourceSecurityPath = Join-Path $config.ExportRoot 'Security'

# Load Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop } catch { }

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

$btnLoadTargetOUs = New-Object System.Windows.Forms.Button
$btnLoadTargetOUs.Text = "Load Target OUs"
$btnLoadTargetOUs.Location = New-Object System.Drawing.Point(440, 9)
$btnLoadTargetOUs.Width = 130
$panelTop.Controls.Add($btnLoadTargetOUs)

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
$script:saveErrorCount = 0
$script:targetOUs = @()
$script:explicitKeepGroups = @{}
$script:protectedDefaultGroupNames = @(
    'Domain Admins',
    'Domain Controllers',
    'Schema Admins',
    'Enterprise Admins',
    'Group Policy Creator Owners',
    'Read-only Domain Controllers',
    'Cloneable Domain Controllers',
    'Protected Users',
    'Key Admins',
    'Enterprise Key Admins'
)

$latestMembers = Get-ChildItem -Path $SourceSecurityPath -Filter "GroupMembers_*.csv" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latestMembers) {
    try {
        Import-Csv $latestMembers.FullName | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_.GroupSam)) {
                $script:explicitKeepGroups[[string]$_.GroupSam] = $true
            }
        }
    } catch {
        Write-Host "[!] WARNING: Failed to load explicit privileged membership keeps from $($latestMembers.FullName): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Get-UserInputText {
    param(
        [string]$Prompt,
        [string]$Title,
        [string]$DefaultText = ""
    )

    $interactionType = [Type]::GetType('Microsoft.VisualBasic.Interaction, Microsoft.VisualBasic', $false)
    if ($null -ne $interactionType) {
        return [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, $DefaultText)
    }

    $dialogForm = New-Object System.Windows.Forms.Form
    $dialogForm.Text = $Title
    $dialogForm.Size = New-Object System.Drawing.Size(560, 165)
    $dialogForm.StartPosition = 'CenterParent'
    $dialogForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dialogForm.MaximizeBox = $false
    $dialogForm.MinimizeBox = $false

    $lblPrompt = New-Object System.Windows.Forms.Label
    $lblPrompt.Text = $Prompt
    $lblPrompt.AutoSize = $false
    $lblPrompt.Size = New-Object System.Drawing.Size(520, 36)
    $lblPrompt.Location = New-Object System.Drawing.Point(12, 10)
    $dialogForm.Controls.Add($lblPrompt)

    $txtValue = New-Object System.Windows.Forms.TextBox
    $txtValue.Text = $DefaultText
    $txtValue.Size = New-Object System.Drawing.Size(520, 22)
    $txtValue.Location = New-Object System.Drawing.Point(12, 50)
    $dialogForm.Controls.Add($txtValue)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Size = New-Object System.Drawing.Size(90, 28)
    $btnOk.Location = New-Object System.Drawing.Point(350, 85)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dialogForm.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Size = New-Object System.Drawing.Size(90, 28)
    $btnCancel.Location = New-Object System.Drawing.Point(442, 85)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dialogForm.Controls.Add($btnCancel)

    $dialogForm.AcceptButton = $btnOk
    $dialogForm.CancelButton = $btnCancel

    $dialogResult = $dialogForm.ShowDialog()
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        return $txtValue.Text
    }
    return ''
}

function Set-TargetOUColumnEditor {
    param([System.Windows.Forms.DataGridView]$Grid)

    if ($null -eq $Grid -or -not $Grid.Columns.Contains('TargetOU_DN')) { return }

    $dt = $Grid.DataSource
    $targetValues = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($ou in $script:targetOUs) {
        if (-not [string]::IsNullOrWhiteSpace($ou)) { [void]$targetValues.Add([string]$ou) }
    }

    if ($null -ne $dt -and $dt.Columns.Contains('TargetOU_DN')) {
        foreach ($r in $dt.Rows) {
            $existing = [string]$r['TargetOU_DN']
            if (-not [string]::IsNullOrWhiteSpace($existing)) { [void]$targetValues.Add($existing) }
        }
    }

    if ($targetValues.Count -eq 0) { return }

    $colIndex = $Grid.Columns['TargetOU_DN'].Index
    $Grid.Columns.Remove('TargetOU_DN')

    $ouCol = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
    $ouCol.Name = 'TargetOU_DN'
    $ouCol.DataPropertyName = 'TargetOU_DN'
    $ouCol.HeaderText = 'TargetOU_DN'
    $ouCol.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $ouCol.MinimumWidth = 250
    $ouCol.DropDownWidth = 500
    $ouCol.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    $ouCol.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightYellow
    $ouCol.Items.AddRange([string[]]($targetValues | Sort-Object))

    $Grid.Columns.Insert($colIndex, $ouCol)
}

function Get-TargetOUs {
    if (-not (Get-Command Get-ADOrganizationalUnit -ErrorAction SilentlyContinue)) {
        [System.Windows.Forms.MessageBox]::Show("Active Directory cmdlets are not available on this system. Install RSAT/ActiveDirectory module to load target OUs.", "AD Cmdlets Missing", "OK", "Warning") | Out-Null
        return
    }

    $targetDomain = Get-UserInputText -Prompt "Enter the target domain FQDN or DC hostname to query OUs (e.g., target.local):" -Title "Load Target OUs" -DefaultText ""
    if ([string]::IsNullOrWhiteSpace($targetDomain)) { return }

    try {
        $rootDse = Get-ADRootDSE -Server $targetDomain -ErrorAction Stop
        $baseDn = $rootDse.defaultNamingContext
        $script:targetOUs = @(Get-ADOrganizationalUnit -Filter * -SearchBase $baseDn -Server $targetDomain -ErrorAction Stop |
            Select-Object -ExpandProperty DistinguishedName |
            Where-Object { $_ -notmatch "(?:^|,)OU=Domain Controllers," } |
            Sort-Object -Unique)

        foreach ($gridPath in $script:grids.Keys) {
            Set-TargetOUColumnEditor -Grid $script:grids[$gridPath]
        }

        [System.Windows.Forms.MessageBox]::Show("Loaded $($script:targetOUs.Count) target OUs from '$targetDomain'. TargetOU_DN now uses a dropdown list.", "Target OUs Loaded", "OK", "Information") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to load target OUs from '$targetDomain'.`n`n$($_.Exception.Message)", "Load Target OUs Failed", "OK", "Error") | Out-Null
    }
}

function Import-CsvToGrid ($fileName, $tabName) {
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
    
    # Load OU Map to check for skipped OUs
    $ouMapPath = Join-Path $MapPath 'OU_Map_Draft.csv'
    $ouMap = @{}
    if (Test-Path $ouMapPath) {
        Import-Csv $ouMapPath | ForEach-Object {
            if ($_.Action -eq 'Skip' -and -not [string]::IsNullOrWhiteSpace($_.SourceDN)) {
                $ouMap[$_.SourceDN] = $_.TargetDN
            }
        }
    }

    foreach ($row in $data) {
        $dr = $dt.NewRow()
        foreach ($prop in $props) {
            $dr[$prop] = $row.$prop
        }

        # Ensure built-in principals are clearly marked and skipped by default.
        $sourceDN = if ($dt.Columns.Contains('SourceDN')) { [string]$dr['SourceDN'] } else { '' }
        $sourceSam = if ($dt.Columns.Contains('SourceSam')) { [string]$dr['SourceSam'] } else { '' }
        $sourceName = if ($dt.Columns.Contains('SourceName')) { [string]$dr['SourceName'] } else { '' }
        $actionVal = if ($dt.Columns.Contains('Action')) { [string]$dr['Action'] } else { '' }
        $isBuiltinPrincipal = $false
        $isProtectedDefaultGroup = $false
        $hasExplicitKeep = $script:explicitKeepGroups.ContainsKey($sourceSam) -or $script:explicitKeepGroups.ContainsKey($sourceName)
        if ($sourceDN -match '(?:^|,)CN=Builtin,') { $isBuiltinPrincipal = $true }
        if ($sourceSam -eq 'System Managed Accounts Group' -or $sourceName -eq 'System Managed Accounts Group') { $isBuiltinPrincipal = $true }
        if ($sourceSam -eq 'DefaultAccount' -or $sourceName -eq 'DefaultAccount') { $isBuiltinPrincipal = $true }
        if ($script:protectedDefaultGroupNames -contains $sourceSam -or $script:protectedDefaultGroupNames -contains $sourceName) { $isProtectedDefaultGroup = $true }
        $shouldForceSafeDefault = [string]::IsNullOrWhiteSpace($actionVal) -or $actionVal -eq 'Create'

        if ($hasExplicitKeep -and ($isBuiltinPrincipal -or $isProtectedDefaultGroup)) {
            if ($isBuiltinPrincipal -and $dt.Columns.Contains('TargetOU_DN')) { $dr['TargetOU_DN'] = 'BUILTIN (Do not move)' }
            if ($isProtectedDefaultGroup -and $dt.Columns.Contains('TargetOU_DN')) { $dr['TargetOU_DN'] = 'SYSTEM DEFAULT GROUP (Do not move)' }
            if ($dt.Columns.Contains('Action')) { $dr['Action'] = 'Merge' }
            if ($dt.Columns.Contains('Notes') -and [string]::IsNullOrWhiteSpace([string]$dr['Notes'])) {
                $dr['Notes'] = 'Privileged membership was explicitly preserved during export. Action set to Merge.'
            }
            $principalName = if (-not [string]::IsNullOrWhiteSpace($sourceSam)) { $sourceSam } else { $sourceName }
            Write-Log -Message "Account Mapper override applied for '$principalName': explicit privileged membership keep detected; Action forced to Merge." -Level WARN
        } elseif ($isBuiltinPrincipal -and $shouldForceSafeDefault) {
            if ($dt.Columns.Contains('TargetOU_DN')) { $dr['TargetOU_DN'] = 'BUILTIN (Do not move)' }
            if ($dt.Columns.Contains('Action')) { $dr['Action'] = 'Skip' }
            if ($dt.Columns.Contains('Notes') -and [string]::IsNullOrWhiteSpace([string]$dr['Notes'])) {
                $dr['Notes'] = 'Built-in principal. Action set to Skip.'
            }
        } elseif ($isProtectedDefaultGroup -and $shouldForceSafeDefault) {
            if ($dt.Columns.Contains('TargetOU_DN')) { $dr['TargetOU_DN'] = 'SYSTEM DEFAULT GROUP (Do not move)' }
            if ($dt.Columns.Contains('Action')) { $dr['Action'] = 'Skip' }
            if ($dt.Columns.Contains('Notes') -and [string]::IsNullOrWhiteSpace([string]$dr['Notes'])) {
                $dr['Notes'] = 'Protected default domain group. Action set to Skip.'
            }
        }

        # If the row's SourceDN parent OU was skipped, set TargetOU_DN to warning
        if ($dt.Columns.Contains('SourceDN') -and $dt.Columns.Contains('TargetOU_DN')) {
            $parentDN = $dr['SourceDN'] -replace '^[^,]+,', ''
            if ($ouMap.ContainsKey($parentDN)) {
                $dr['TargetOU_DN'] = 'WARNING: Source OU was skipped during mapping to target domain. Please specify a valid Target OU or set Action to Skip.'
            }
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

    # --- COPY/PASTE SUPPORT ---
    $dgv.Add_KeyDown({
        param($grid, $e)
        if ($e.Control -and $e.KeyCode -eq 'C') {
            # Copy selected cell(s) to clipboard
            $sb = New-Object System.Text.StringBuilder
            $selRows = $grid.SelectedCells | Sort-Object RowIndex, ColumnIndex | Group-Object RowIndex
            foreach ($rowGroup in $selRows) {
                $rowText = ($rowGroup.Group | Sort-Object ColumnIndex | ForEach-Object { $_.Value }) -join "\t"
                $sb.AppendLine($rowText) | Out-Null
            }
            [Windows.Forms.Clipboard]::SetText($sb.ToString().TrimEnd())
            $e.SuppressKeyPress = $true
        } elseif ($e.Control -and $e.KeyCode -eq 'V') {
            # Paste clipboard text into selected cells (row-wise, col-wise)
            $clipText = [Windows.Forms.Clipboard]::GetText()
            if (-not [string]::IsNullOrWhiteSpace($clipText) -and $grid.SelectedCells.Count -gt 0) {
                $startCell = $grid.SelectedCells | Sort-Object RowIndex, ColumnIndex | Select-Object -First 1
                $rows = $clipText -split "\r?\n"
                for ($i = 0; $i -lt $rows.Count; $i++) {
                    $cols = $rows[$i] -split "\t"
                    $rowIdx = $startCell.RowIndex + $i
                    if ($rowIdx -ge $grid.RowCount) { break }
                    for ($j = 0; $j -lt $cols.Count; $j++) {
                        $colIdx = $startCell.ColumnIndex + $j
                        if ($colIdx -ge $grid.ColumnCount) { break }
                        $cell = $grid.Rows[$rowIdx].Cells[$colIdx]
                        if (-not $cell.ReadOnly) { $cell.Value = $cols[$j] }
                    }
                }
                $e.SuppressKeyPress = $true
            }
        }
    })
    
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
        $dgv.Columns["TargetOU_DN"].MinimumWidth = 250
        $dgv.Columns["TargetOU_DN"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
    }

    # If target OUs have been loaded, present TargetOU_DN as a dropdown.
    Set-TargetOUColumnEditor -Grid $dgv

    # --- HIGHLIGHT INVALID ACCOUNTS ---
    $invalidRows = @()
    if ($dt.Columns.Contains('Action') -and $dt.Columns.Contains('TargetOU_DN')) {
        for ($i = 0; $i -lt $dt.Rows.Count; $i++) {
            $row = $dt.Rows[$i]
            $action = $row['Action']
            $ou = $row['TargetOU_DN']
            if ($action -eq 'Create' -and ([string]::IsNullOrWhiteSpace($ou) -or $ou -eq 'SKIPPED' -or $ou -eq 'UNSPECIFIED' -or $ou -like 'WARNING:*')) {
                $invalidRows += $i
            }
        }
        foreach ($idx in $invalidRows) {
            foreach ($cell in $dgv.Rows[$idx].Cells) {
                $cell.Style.BackColor = [System.Drawing.Color]::Salmon
            }
        }
        if ($invalidRows.Count -gt 0) {
            $tab.Text = "[!] $tabName ($($dt.Rows.Count))"
        }
    }

    $tabCtrl.TabPages.Add($tab)
    $script:grids[$csvPath] = $dgv
}

$btnSave.Add_Click({
    $hasInvalid = $false
    $dgvCount = 0
    $firstErrorTab = $null
    foreach ($key in $script:grids.Keys) {
        $dgv = $script:grids[$key]
        # Force any pending edits to commit
        $dgv.EndEdit()
        $dgv.CurrentCell = $null
        $dt = $dgv.DataSource
        $tab = $dgv.Parent
        $baseName = $tab.Tag
        $invalidRows = @()
        if ($dt.Columns.Contains('Action') -and $dt.Columns.Contains('TargetOU_DN')) {
            for ($i = 0; $i -lt $dt.Rows.Count; $i++) {
                $row = $dt.Rows[$i]
                $action = $row['Action']
                $ou = $row['TargetOU_DN']
                if ($action -eq 'Create' -and ([string]::IsNullOrWhiteSpace($ou) -or $ou -eq 'SKIPPED' -or $ou -eq 'UNSPECIFIED' -or $ou -like 'WARNING:*')) {
                    $invalidRows += $i
                }
            }
        }
        if ($invalidRows.Count -gt 0) {
            $hasInvalid = $true
            if ($null -eq $firstErrorTab) { $firstErrorTab = $tab }
            $tab.Text = "[!] $baseName ($($dt.Rows.Count))"
            # Highlight invalid rows in red
            foreach ($idx in $invalidRows) {
                foreach ($cell in $dgv.Rows[$idx].Cells) {
                    $cell.Style.BackColor = [System.Drawing.Color]::Salmon
                }
            }
        } else {
            $tab.Text = "$baseName ($($dt.Rows.Count))"
        }
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
    if ($hasInvalid) {
        if ($firstErrorTab) { $tabCtrl.SelectedTab = $firstErrorTab }
        [System.Windows.Forms.MessageBox]::Show("WARNING: Some accounts are still mapped to invalid, skipped, or unspecified OUs. These rows are highlighted in red. Validation will fail until all issues are resolved.`n`nPlease check all tabs for completion.", "Unresolved Account Placement Issues", "OK", "Warning")
    } else {
        [System.Windows.Forms.MessageBox]::Show("Successfully saved $dgvCount mapping files.", "Save Complete", "OK", "Information")
        $form.Close()
    }
})

# Load the mappings
Import-CsvToGrid "User_Account_Map.csv" "Users"
Import-CsvToGrid "Computer_Account_Map.csv" "Computers"
Import-CsvToGrid "Group_Account_Map.csv" "Groups"

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
        
        $hasError = $tab.Text -match "^\[!\]"
        $prefix = if ($hasError) { "[!] " } else { "" }

        if ([string]::IsNullOrWhiteSpace($searchText)) {
            $dt.DefaultView.RowFilter = ""
            $tab.Text = "$prefix$baseName ($totalRows)"
            continue
        }

        $filterParts = @()
        foreach ($col in $dt.Columns) {
            $filterParts += "[$($col.ColumnName)] LIKE '*$searchText*'"
        }
        $dt.DefaultView.RowFilter = $filterParts -join " OR "
        
        $filteredRows = $dt.DefaultView.Count
        $tab.Text = "$prefix$baseName ($filteredRows / $totalRows)"
    }
})

$btnSearchClear.Add_Click({ $txtSearch.Text = "" })
$btnLoadTargetOUs.Add_Click({ Get-TargetOUs })

# --- SHOW ---
$form.ShowDialog() | Out-Null