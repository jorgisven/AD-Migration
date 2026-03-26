<#
.SYNOPSIS
    GUI Tool for Visual OU Mapping.

.DESCRIPTION
    Provides a dual-pane interface to map Source OUs to a Target Domain structure via Drag & Drop.
    Outputs the standard 'OU_Map_Draft.csv' required by the Import phase.

.PARAMETER TargetDomain
    The root DN of the target domain (e.g., "DC=target,DC=local").
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$SourceOUPath = Join-Path $config.ExportRoot 'OU_Structure'
$MapPath = Join-Path $config.TransformRoot 'Mapping'

# Ensure transform directory exists
if (-not (Test-Path $MapPath)) { New-Item -ItemType Directory -Path $MapPath -Force | Out-Null }

# Load Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- DATA LOADING ---

if (-not $TargetDomain) {
    $TargetDomain = (Get-ADDomain).DistinguishedName
}

# Find latest export
$latestExport = Get-ChildItem -Path $SourceOUPath -Filter "OU_Structure_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $latestExport) {
    [System.Windows.Forms.MessageBox]::Show("No OU Export found. Run Export-OUs.ps1 first.", "Error", "OK", "Error")
    exit
}

$SourceOUs = Import-Csv $latestExport.FullName | Sort-Object { $_.DistinguishedName.Length }

$EmptyOUs = @()
$emptyListFile = Join-Path $SourceOUPath "EmptyOUs.txt"
if (Test-Path $emptyListFile) {
    $EmptyOUs = @(Get-Content $emptyListFile)
}

# --- GPO LINKED OUs LOADING ---
$GpoLinkedOUs = @()
$ReportPath = Join-Path $config.ExportRoot 'GPO_Reports'
if (Test-Path $ReportPath) {
    $xmlFiles = Get-ChildItem -Path $ReportPath -Filter "*.xml"
    foreach ($xmlFile in $xmlFiles) {
        try {
            [xml]$xml = Get-Content $xmlFile.FullName
            $gpoName = $xml.GPO.Name
            $links = $xml.GPO.LinksTo
            if ($links) {
                foreach ($link in $links) {
                    $GpoLinkedOUs += [PSCustomObject]@{
                        GPO_Name = $gpoName
                        SOMPath  = ([string]$link.SOMPath).Trim() -replace '\s*,\s*', ','
                    }
                }
            }
        } catch { }
    }
}

# --- GUI SETUP ---

$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Migration - OU Mapper"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"

# --- GPO INHERITANCE WARNING ---
[System.Windows.Forms.MessageBox]::Show(
    "IMPORTANT: If you change the OU structure or move OUs during migration, inherited GPOs from the source domain will NOT be automatically re-linked in the target domain.`nYou must manually re-link GPOs at the appropriate OUs in the target domain to restore intended inheritance.",
    "GPO Inheritance Notice",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
)

# Layout
$splitContainer = New-Object System.Windows.Forms.SplitContainer
$splitContainer.Dock = "Fill"
$splitContainer.SplitterDistance = 400
$form.Controls.Add($splitContainer)

# Left Pane (Source)
$lblSource = New-Object System.Windows.Forms.Label
$lblSource.Text = "Source Domain (Read-Only)"
$lblSource.Dock = "Top"
$splitContainer.Panel1.Controls.Add($lblSource)

# Search Source
$pnlSourceSearch = New-Object System.Windows.Forms.Panel
$pnlSourceSearch.Height = 25
$pnlSourceSearch.Dock = "Top"
$splitContainer.Panel1.Controls.Add($pnlSourceSearch)

$btnSourceSearch = New-Object System.Windows.Forms.Button; $btnSourceSearch.Text = "Find"; $btnSourceSearch.Width = 50; $btnSourceSearch.Dock = "Right"; $pnlSourceSearch.Controls.Add($btnSourceSearch)
$txtSourceSearch = New-Object System.Windows.Forms.TextBox; $txtSourceSearch.Dock = "Fill"; $pnlSourceSearch.Controls.Add($txtSourceSearch)
$txtSourceSearch.BringToFront()

$treeSource = New-Object System.Windows.Forms.TreeView
$treeSource.Dock = "Fill"
$treeSource.AllowDrop = $false # Source is drag source only
$splitContainer.Panel1.Controls.Add($treeSource)
$treeSource.BringToFront()

# Right Pane (Target)
$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "Target Domain (Drag here to Map)"
$lblTarget.Dock = "Top"
$splitContainer.Panel2.Controls.Add($lblTarget)

# Search Target
$pnlTargetSearch = New-Object System.Windows.Forms.Panel
$pnlTargetSearch.Height = 25
$pnlTargetSearch.Dock = "Top"
$splitContainer.Panel2.Controls.Add($pnlTargetSearch)

$btnTargetSearch = New-Object System.Windows.Forms.Button; $btnTargetSearch.Text = "Find"; $btnTargetSearch.Width = 50; $btnTargetSearch.Dock = "Right"; $pnlTargetSearch.Controls.Add($btnTargetSearch)
$txtTargetSearch = New-Object System.Windows.Forms.TextBox; $txtTargetSearch.Dock = "Fill"; $pnlTargetSearch.Controls.Add($txtTargetSearch)
$txtTargetSearch.BringToFront()

$treeTarget = New-Object System.Windows.Forms.TreeView
$treeTarget.Dock = "Fill"
$treeTarget.AllowDrop = $true
$treeTarget.LabelEdit = $true # Allow renaming
$splitContainer.Panel2.Controls.Add($treeTarget)
$treeTarget.BringToFront()

# Bottom Panel (Buttons)
$panelBottom = New-Object System.Windows.Forms.Panel
$panelBottom.Height = 50
$panelBottom.Dock = "Bottom"
$form.Controls.Add($panelBottom)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save Mapping CSV"
$btnSave.Width = 150
$btnSave.Location = New-Object System.Drawing.Point(820, 10)
$panelBottom.Controls.Add($btnSave)

$btnUndo = New-Object System.Windows.Forms.Button
$btnUndo.Text = "Undo"
$btnUndo.Width = 80
$btnUndo.Location = New-Object System.Drawing.Point(730, 10)
$btnUndo.Enabled = $false
$panelBottom.Controls.Add($btnUndo)

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Drag OUs from Left to Right. Right-click Target nodes to Rename/Delete."
$lblInfo.AutoSize = $true
$lblInfo.Location = New-Object System.Drawing.Point(10, 15)
$panelBottom.Controls.Add($lblInfo)

# --- FUNCTIONS ---

# State Management for Undo/Copy
$script:UndoStack = New-Object System.Collections.Generic.List[object]
$script:MaxUndo = 20
$script:ClipboardNode = $null

function Save-TreeState {
    $snapshot = @()
    foreach ($n in $treeTarget.Nodes) {
        $snapshot += $n.Clone()
    }
    $script:UndoStack.Add($snapshot)
    if ($script:UndoStack.Count -gt $script:MaxUndo) {
        $script:UndoStack.RemoveAt(0)
    }
    $btnUndo.Enabled = $true
}

function Undo-Action {
    if ($script:UndoStack.Count -gt 0) {
        $lastIndex = $script:UndoStack.Count - 1
        $snapshot = $script:UndoStack[$lastIndex]
        $script:UndoStack.RemoveAt($lastIndex)
        
        $treeTarget.BeginUpdate()
        $treeTarget.Nodes.Clear()
        foreach ($n in $snapshot) {
            $treeTarget.Nodes.Add($n) | Out-Null
        }
        $treeTarget.ExpandAll()
        $treeTarget.EndUpdate()
        
        $btnUndo.Enabled = ($script:UndoStack.Count -gt 0)
    }
}

function Copy-Action {
    if ($treeTarget.SelectedNode) {
        $script:ClipboardNode = $treeTarget.SelectedNode.Clone()
    }
}

function Invoke-PasteAction {
    if ($script:ClipboardNode -and $treeTarget.SelectedNode) {
        $targetNode = $treeTarget.SelectedNode
        $existing = $targetNode.Nodes | Where-Object { $_.Text -eq $script:ClipboardNode.Text }
        if ($existing) {
            [System.Windows.Forms.MessageBox]::Show("An OU with the name '$($script:ClipboardNode.Text)' already exists under '$($targetNode.Text)'.", "Conflict", "OK", "Warning")
            return
        }
        Save-TreeState
        $targetNode.Nodes.Add($script:ClipboardNode.Clone()) | Out-Null
        $targetNode.Expand()
    }
}

function Remove-Action {
    $node = $treeTarget.SelectedNode
    if ($node) {
        if ($null -eq $node.Parent) {
            [System.Windows.Forms.MessageBox]::Show("You cannot delete the root domain node.", "Action Denied", "OK", "Warning")
            return
        }
        if ($node.Tag -and $node.Tag.SourceDN -eq 'PRE-EXISTING') {
            [System.Windows.Forms.MessageBox]::Show("This OU already exists in the live target domain and serves as an anchor point for your migration.`n`nIt cannot be removed from the mapping view.", "Cannot Delete Pre-Existing OU", "OK", "Warning")
            return
        }
        Save-TreeState
        $node.Remove()
    }
}

function Add-NodeToTree ($tree, $dn, $tagData) {
    # Parse DN to find hierarchy
    # Simple parser: assumes standard comma separation. 
    # For robust DN parsing, specific regex is needed, but this suffices for visualization.
    
    # We only care about OUs and DCs.
    $parts = @($dn -split "(?<!\\)," | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrEmpty($_) })
    # Reverse to build from root up
    [array]::Reverse($parts)
    
    $currentNode = $null
    $nodesCollection = $tree.Nodes

    for ($i = 0; $i -lt $parts.Count; $i++) {
        $part = $parts[$i]
        # Check if node exists at this level
        $found = $nodesCollection | Where-Object { $_.Text -eq $part } | Select-Object -First 1
        
        if ($found) {
            $currentNode = $found
        } else {
            $currentNode = New-Object System.Windows.Forms.TreeNode
            $currentNode.Text = $part
            $currentNode.Name = $part # Use Name for lookup
            $currentNode.ImageIndex = 0
            
            # If this is the leaf node (the full DN), attach the data
            if ($i -eq ($parts.Count - 1) -and $tagData) {
                $currentNode.Tag = $tagData
                $currentNode.ForeColor = if ($tagData.SourceDN -eq 'PRE-EXISTING') { [System.Drawing.Color]::Gray } else { [System.Drawing.Color]::DarkBlue }
            }

            $nodesCollection.Add($currentNode) | Out-Null
        }
        $nodesCollection = $currentNode.Nodes
    }
    return $currentNode
}

function Search-Tree ($tree, $text) {
    if ([string]::IsNullOrWhiteSpace($text)) { return }
    
    $foundNodes = @()
    
    # Recursive search
    function Find-Nodes ($nodes) {
        foreach ($node in $nodes) {
            # Reset color
            $node.BackColor = [System.Drawing.Color]::Empty
            
            if ($node.Text -like "*$text*") {
                $foundNodes += $node
            }
            
            if ($node.Nodes.Count -gt 0) {
                Find-Nodes $node.Nodes
            }
        }
    }
    
    $tree.BeginUpdate()
    Find-Nodes $tree.Nodes
    
    if ($foundNodes.Count -gt 0) {
        foreach ($node in $foundNodes) {
            $node.EnsureVisible()
            $node.BackColor = [System.Drawing.Color]::Yellow
        }
        $tree.SelectedNode = $foundNodes[0]
        $tree.Focus()
    } else {
        [System.Windows.Forms.MessageBox]::Show("No matches found.", "Search", "OK", "Information")
    }
    $tree.EndUpdate()
}

$btnUndo.Add_Click({ Undo-Action })

# Populate Source Tree
$treeSource.BeginUpdate()
foreach ($ou in $SourceOUs) {
    $tag = @{
        SourceDN = $ou.DistinguishedName
        Description = $ou.Description
        SourceOU = $ou.OU
    }
    Add-NodeToTree -tree $treeSource -dn $ou.DistinguishedName -tagData $tag | Out-Null
}
$treeSource.ExpandAll()
if ($treeSource.Nodes.Count -gt 0) { $treeSource.Nodes[0].EnsureVisible() }
$treeSource.EndUpdate()

# Initialize Target Tree
$treeTarget.BeginUpdate()

# --- Pre-load from Live Target Domain (Optional) ---
$msg = "Do you want to load the existing OU structure from the live target domain '$TargetDomain'?`n`nThis is useful if you are mapping into a pre-existing structure."
$result = [System.Windows.Forms.MessageBox]::Show($msg, "Load Target Structure", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

if ($result -eq 'Yes') {
    try {
        # The -Server parameter requires an FQDN (e.g., target.local), not a DN (e.g., DC=target,DC=local). Convert it.
        $TargetServer = (($TargetDomain -replace 'DC=','' -split '(?<!\\),') -join '.')
        Write-Host "Attempting to connect to server '$TargetServer' to load existing OUs..." -ForegroundColor Yellow
        $TargetOUs = Get-ADOrganizationalUnit -Filter * -Server $TargetServer -Properties Description | Sort-Object { $_.DistinguishedName.Length }
        Write-Host "Found $($TargetOUs.Count) existing OUs in target domain." -ForegroundColor Green
        
        # Add the domain root first
        Add-NodeToTree -tree $treeTarget -dn $TargetDomain -tagData $null | Out-Null
        
        foreach ($ou in $TargetOUs) {
            # Create a tag for these so we know they are pre-existing, not from a source mapping
            $tag = @{
                SourceDN    = "PRE-EXISTING"
                Description = $ou.Description
                SourceOU    = $ou.Name
            }
            Add-NodeToTree -tree $treeTarget -dn $ou.DistinguishedName -tagData $tag | Out-Null
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to connect to target domain '$TargetServer'.`n`nError: $($_.Exception.Message)`n`nContinuing with an empty target structure.", "Connection Error", "OK", "Error")
        # If it fails, just add the root node
        Add-NodeToTree -tree $treeTarget -dn $TargetDomain -tagData $null | Out-Null
    }
} else {
    # User said no, just add the root node
    Add-NodeToTree -tree $treeTarget -dn $TargetDomain -tagData $null | Out-Null
}

# Check for existing draft and merge it on top of the current view
$draftFile = Join-Path $MapPath "OU_Map_Draft.csv"
if (Test-Path $draftFile) {
    Write-Host "Loading existing draft file: $draftFile" -ForegroundColor Cyan
    $Draft = Import-Csv $draftFile
    foreach ($row in $Draft) {
        if ($row.Action -ne 'Skip' -and -not [string]::IsNullOrWhiteSpace($row.TargetDN)) {
            $tag = @{
                SourceDN    = $row.SourceDN
                Description = $row.Description
                SourceOU    = $row.SourceOU
            }
            Add-NodeToTree -tree $treeTarget -dn $row.TargetDN -tagData $tag | Out-Null
        }
    }
}

$treeTarget.ExpandAll()
$treeTarget.EndUpdate()

# --- EVENTS ---

# Drag Start (Source & Target)
$DragHandler = {
    param($eventSender, $e)
    if ($e.Item) {
        $eventSender.DoDragDrop($e.Item, [System.Windows.Forms.DragDropEffects]::Move)
    }
}
$treeSource.Add_ItemDrag($DragHandler)
$treeTarget.Add_ItemDrag($DragHandler)

# Drag Enter (Target)
$treeTarget.Add_DragEnter({
    param($eventSender, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.TreeNode])) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
    }
})

# Drag Drop (Target)
$treeTarget.Add_DragDrop({
    param($eventSender, $e)
    
    $targetPoint = $eventSender.PointToClient([System.Drawing.Point]::new($e.X, $e.Y))
    $targetNode = $eventSender.GetNodeAt($targetPoint)
    $draggedNode = $e.Data.GetData([System.Windows.Forms.TreeNode])

        if ($targetNode -and $targetNode.Text -match "^OU=Domain Controllers$") {
            [System.Windows.Forms.MessageBox]::Show("The 'Domain Controllers' OU is managed by Active Directory and cannot be used as a migration target.`n`nPlease select a different OU.", "Action Denied", "OK", "Warning")
            return
        }

        if ($draggedNode -and $draggedNode.Text -match "^OU=Domain Controllers$") {
            [System.Windows.Forms.MessageBox]::Show("The source 'Domain Controllers' OU cannot be migrated.`n`nDomain Controllers must be manually promoted in the target domain.", "Action Denied", "OK", "Warning")
            return
        }

    # Handle MERGE action (dragging from source onto an existing target node)
    if ($targetNode -and $draggedNode -and ($draggedNode.TreeView -eq $treeSource)) {
        # A node with a Tag is either pre-existing or already mapped.
        if ($targetNode.Tag) {
            $mergeForm = New-Object System.Windows.Forms.Form
            $mergeForm.Text = "Merge or Add as Child?"
            $mergeForm.Size = New-Object System.Drawing.Size(460, 200)
            $mergeForm.StartPosition = "CenterParent"
            $mergeForm.FormBorderStyle = "FixedDialog"
            $mergeForm.MaximizeBox = $false
            $mergeForm.MinimizeBox = $false

            $lblMergeMsg = New-Object System.Windows.Forms.Label
            $lblMergeMsg.Text = "You are dropping onto an existing target OU.`n`nSource: '$($draggedNode.Text)'`nTarget: '$($targetNode.Text)'`n`nWould you like to Merge the source into the target, or create it as a new Child OU?"
            $lblMergeMsg.Location = New-Object System.Drawing.Point(20, 20)
            $lblMergeMsg.Size = New-Object System.Drawing.Size(400, 70)
            $mergeForm.Controls.Add($lblMergeMsg)

            $btnMerge = New-Object System.Windows.Forms.Button; $btnMerge.Text = "Merge"; $btnMerge.Location = New-Object System.Drawing.Point(40, 110); $btnMerge.Size = New-Object System.Drawing.Size(110, 30); $btnMerge.DialogResult = [System.Windows.Forms.DialogResult]::Yes; $mergeForm.Controls.Add($btnMerge)
            $btnChild = New-Object System.Windows.Forms.Button; $btnChild.Text = "Create Child OU"; $btnChild.Location = New-Object System.Drawing.Point(160, 110); $btnChild.Size = New-Object System.Drawing.Size(120, 30); $btnChild.DialogResult = [System.Windows.Forms.DialogResult]::No; $mergeForm.Controls.Add($btnChild)
            $btnCancelMerge = New-Object System.Windows.Forms.Button; $btnCancelMerge.Text = "Cancel"; $btnCancelMerge.Location = New-Object System.Drawing.Point(290, 110); $btnCancelMerge.Size = New-Object System.Drawing.Size(100, 30); $btnCancelMerge.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $mergeForm.Controls.Add($btnCancelMerge)

            $mergeForm.AcceptButton = $btnMerge
            $mergeForm.CancelButton = $btnCancelMerge

            $result = $mergeForm.ShowDialog()
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Save-TreeState
                # Perform the merge: copy the tag from source to target.
                $targetNode.Tag = $draggedNode.Tag
                $targetNode.ForeColor = [System.Drawing.Color]::DarkBlue # Mark as mapped
                return # The drop action is complete.
            } elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
                return # Cancel the drop entirely
            }
            # If user says No (Create Child OU), we fall through to the default "add as child" behavior.
        }
    }

    if ($targetNode -and $draggedNode) {
        # Prevent dropping into itself or children
        $parent = $targetNode
        while ($parent) {
            if ($parent -eq $draggedNode) { return }
            $parent = $parent.Parent
        }

        # Check for duplicate name in destination
        $existing = $targetNode.Nodes | Where-Object { $_.Text -eq $draggedNode.Text }
        if ($existing) {
            [System.Windows.Forms.MessageBox]::Show("An OU with the name '$($draggedNode.Text)' already exists under '$($targetNode.Text)'.", "Conflict", "OK", "Warning")
            return
        }

        Save-TreeState

        # Clone node to move it (or copy if from source)
        $newNode = $draggedNode.Clone()
        
        # If dragging from Source, we keep the original in Source (Copy effect)
        # If dragging from Target, we remove the original (Move effect)
        if ($draggedNode.TreeView -eq $treeTarget) {
            $draggedNode.Remove()
        }

        $targetNode.Nodes.Add($newNode) | Out-Null
        $targetNode.Expand()
    }
})

# Context Menu for Target
$ctxMenu = New-Object System.Windows.Forms.ContextMenu
$itemDelete = $ctxMenu.MenuItems.Add("Delete")
$itemDelete.Add_Click({ Remove-Action })
$itemRename = $ctxMenu.MenuItems.Add("Rename")
$itemRename.Add_Click({
    if ($treeTarget.SelectedNode) {
        if ($treeTarget.SelectedNode.Tag -and $treeTarget.SelectedNode.Tag.SourceDN -eq 'PRE-EXISTING') {
            [System.Windows.Forms.MessageBox]::Show("You cannot rename a pre-existing OU in this mapping tool.`n`nIf you need to reorganize or rename the existing target infrastructure, please do so directly in Active Directory (RSAT).", "Cannot Rename Pre-Existing OU", "OK", "Warning")
            return
        }
        $treeTarget.SelectedNode.BeginEdit()
    }
})
$itemNew = $ctxMenu.MenuItems.Add("New OU")
$itemNew.Add_Click({
    if ($treeTarget.SelectedNode) {
        Save-TreeState
        $newNode = New-Object System.Windows.Forms.TreeNode
        $newNode.Text = "OU=NewOU"
        $treeTarget.SelectedNode.Nodes.Add($newNode) | Out-Null
        $treeTarget.SelectedNode.Expand()
        $newNode.BeginEdit()
    }
})

$ctxMenu.MenuItems.Add("-") | Out-Null
$itemCopy = $ctxMenu.MenuItems.Add("Copy (Ctrl+C)")
$itemCopy.Add_Click({ Copy-Action })
$itemPaste = $ctxMenu.MenuItems.Add("Paste (Ctrl+V)")
$itemPaste.Add_Click({ Invoke-PasteAction })

$treeTarget.ContextMenu = $ctxMenu

# Rename Validation
$treeTarget.Add_BeforeLabelEdit({
    param($eventSender, $e)
    if ($e.Node.Tag -and $e.Node.Tag.SourceDN -eq 'PRE-EXISTING') {
        $e.CancelEdit = $true
        [System.Windows.Forms.MessageBox]::Show("You cannot rename a pre-existing OU in this mapping tool.`n`nIf you need to reorganize or rename the existing target infrastructure, please do so directly in Active Directory (RSAT).", "Cannot Rename Pre-Existing OU", "OK", "Warning")
        return
    }
    Save-TreeState
})
$treeTarget.Add_AfterLabelEdit({
    param($eventSender, $e)
    if ($e.Label) {
        # Check invalid chars for AD (Basic check)
        if ($e.Label -match "[\\/,#+<>;=""]") {
            $e.CancelEdit = $true
            [System.Windows.Forms.MessageBox]::Show("Name contains invalid characters.", "Error", "OK", "Error")
            return
        }
        # Check siblings for duplicates
        $siblings = if ($e.Node.Parent) { $e.Node.Parent.Nodes } else { $treeTarget.Nodes }
        
        foreach ($sib in $siblings) {
            if ($sib -ne $e.Node -and $sib.Text -eq $e.Label) {
                $e.CancelEdit = $true
                [System.Windows.Forms.MessageBox]::Show("A sibling with this name already exists.", "Conflict", "OK", "Warning")
                return
            }
        }
    }
})

# Save Logic
$btnSave.Add_Click({
    $MappingData = [System.Collections.Generic.List[PSObject]]::new()

    # Recursive function to walk tree and build DNs
    function Get-TreeData ($nodes) {
        foreach ($node in $nodes) {
            # Build DN by walking up parents
            $dnParts = @($node.Text)
            $p = $node.Parent
            while ($p) {
                $dnParts += $p.Text
                $p = $p.Parent
            }
            $fullTargetDN = ($dnParts -join ",")

            # Only export OUs (nodes starting with OU= or just assume all non-root nodes are OUs)
            # We assume the root is DC=... and everything under it is an OU structure
            
            if ($node.Text -notmatch "^(?:DC|CN)=" -and ($node.Text -match "^OU=" -or $node.Parent)) {
                # If it has a tag, it came from source
                if ($node.Tag) {
                    $MappingData.Add([PSCustomObject]@{
                        Action      = "Migrate"
                        SourceOU    = $node.Tag.SourceOU
                        TargetOU    = ($node.Text -replace "^OU=","")
                        TargetDN    = $fullTargetDN
                        SourceDN    = $node.Tag.SourceDN
                        Description = $node.Tag.Description
                    })
                } else {
                    # New node created in GUI
                    $MappingData.Add([PSCustomObject]@{
                        Action      = "Migrate"
                        SourceOU    = ""
                        TargetOU    = ($node.Text -replace "^OU=","")
                        TargetDN    = $fullTargetDN
                        SourceDN    = "" # No source
                        Description = "Created via GUI Mapper"
                    })
                }
            }

            if ($node.Nodes.Count -gt 0) {
                Get-TreeData $node.Nodes
            }
        }
    }

    Get-TreeData $treeTarget.Nodes

    # Find unmapped source OUs
    $mappedSourceDNs = $MappingData | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SourceDN) } | Select-Object -ExpandProperty SourceDN

    # Auto-skip the Domain Controllers OU so it doesn't trigger unmapped warnings
    $dcOUs = $SourceOUs | Where-Object { $_.DistinguishedName -match "(?:^|,)OU=Domain Controllers," -and $_.DistinguishedName -notin $mappedSourceDNs }
    foreach ($dc in $dcOUs) {
        $MappingData.Add([PSCustomObject]@{
            Action      = "Skip"
            SourceOU    = $dc.OU
            TargetOU    = ""
            TargetDN    = ""
            SourceDN    = $dc.DistinguishedName
            Description = "Built-in Domain Controllers OU (Auto-Skipped)"
        })
    }
    
    # Refresh mapped list before checking for unmapped OUs
    $mappedSourceDNs = $MappingData | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SourceDN) } | Select-Object -ExpandProperty SourceDN
    $unmappedOUs = $SourceOUs | Where-Object { $_.DistinguishedName -notin $mappedSourceDNs }

    # Update Source Tree Colors
    function Update-SourceColors ($nodes) {
        foreach ($node in $nodes) {
            if ($node.Tag) {
                if ($node.Tag.SourceDN -notin $mappedSourceDNs) {
                    $node.ForeColor = [System.Drawing.Color]::Red
                    # Expand parent
                    $p = $node.Parent
                    while ($p) { $p.Expand(); $p = $p.Parent }
                } else {
                    $node.ForeColor = [System.Drawing.Color]::Gray
                }
            }
            if ($node.Nodes.Count -gt 0) {
                Update-SourceColors $node.Nodes
            }
        }
    }
    
    $treeSource.BeginUpdate()
    Update-SourceColors $treeSource.Nodes
    $treeSource.EndUpdate()
    $treeSource.Refresh() # Force WinForms to repaint the TreeView before the message box blocks the thread

    if ($unmappedOUs.Count -gt 0) {
        $nonEmptyUnmapped = $unmappedOUs | Where-Object { $_.DistinguishedName -notin $EmptyOUs }
        
        if ($nonEmptyUnmapped.Count -gt 0) {
            $msg = "CRITICAL WARNING: $($nonEmptyUnmapped.Count) unmapped OUs are NOT EMPTY!`nIf you do not map these, any users or computers inside them will be orphaned or fail to migrate.`n`nUnmapped OUs are highlighted in RED in the Source tree.`n`nAre you absolutely SURE you want to save and skip them?"
        } else {
            $msg = "Not all OUs were migrated ($($unmappedOUs.Count) unmapped).`nThe unmapped OUs are empty, so skipping them is safe.`nUnmapped OUs have been highlighted in RED in the Source tree.`n`nSave anyway?"
        }
        
        $result = [System.Windows.Forms.MessageBox]::Show($msg, "Unmapped OUs Detected", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -eq 'No') {
            return
        }

        # Add unmapped to MappingData so they aren't lost
        foreach ($ou in $unmappedOUs) {
            $MappingData.Add([PSCustomObject]@{
                Action      = "Skip"
                SourceOU    = $ou.OU
                TargetOU    = ""
                TargetDN    = ""
                SourceDN    = $ou.DistinguishedName
                Description = $ou.Description
            })
        }

        # --- GPO LINKED OUs WARNING ---
        if ($GpoLinkedOUs.Count -gt 0) {
            $unmappedDnsMap = @{}
            foreach ($ou in $unmappedOUs) {
                $norm = ($ou.DistinguishedName).Trim() -replace '\s*,\s*', ','
                $unmappedDnsMap[$norm.ToLowerInvariant()] = $ou.DistinguishedName
            }
            
            $linked = @()
            foreach ($gpoLink in $GpoLinkedOUs) {
                $somNorm = $gpoLink.SOMPath.ToLowerInvariant()
                if ($unmappedDnsMap.ContainsKey($somNorm)) {
                    $linked += [PSCustomObject]@{
                        OU_DistinguishedName = $unmappedDnsMap[$somNorm]
                        GPO_Name = $gpoLink.GPO_Name
                    }
                }
            }

            if ($linked.Count -gt 0) {
                $ouGroups = $linked | Group-Object OU_DistinguishedName
                
                $sb = New-Object System.Text.StringBuilder
                $sb.AppendLine("WARNING: The following unmigrated OUs have GPOs linked:`n") | Out-Null
                foreach ($group in $ouGroups) {
                    $sb.AppendLine("OU: $($group.Name)") | Out-Null
                    $sb.AppendLine("  GPOs: " + ($group.Group | Select-Object -ExpandProperty GPO_Name -Unique -join ", ")) | Out-Null
                    $sb.AppendLine() | Out-Null
                }
                $sb.AppendLine("These GPO links will NOT be present in the target domain unless you map these OUs or re-link the GPOs manually later.") | Out-Null
                $sb.AppendLine("Continue saving?") | Out-Null
                
                $gpoForm = New-Object System.Windows.Forms.Form
                $gpoForm.Text = "Unmigrated OUs with GPO Links"
                $gpoForm.Size = New-Object System.Drawing.Size(650, 450)
                $gpoForm.StartPosition = "CenterParent"
                
                $txtReport = New-Object System.Windows.Forms.TextBox
                $txtReport.Multiline = $true
                $txtReport.ScrollBars = "Vertical"
                $txtReport.ReadOnly = $true
                $txtReport.Font = New-Object System.Drawing.Font("Consolas", 10)
                $txtReport.Text = $sb.ToString()
                $txtReport.Dock = "Fill"
                
                $pnlBottom = New-Object System.Windows.Forms.Panel; $pnlBottom.Height = 50; $pnlBottom.Dock = "Bottom"
                $btnYes = New-Object System.Windows.Forms.Button; $btnYes.Text = "Yes, Continue"; $btnYes.DialogResult = "Yes"; $btnYes.Location = New-Object System.Drawing.Point(410, 10); $btnYes.Size = New-Object System.Drawing.Size(100, 30); $pnlBottom.Controls.Add($btnYes)
                $btnNo = New-Object System.Windows.Forms.Button; $btnNo.Text = "No, Cancel"; $btnNo.DialogResult = "No"; $btnNo.Location = New-Object System.Drawing.Point(520, 10); $btnNo.Size = New-Object System.Drawing.Size(100, 30); $pnlBottom.Controls.Add($btnNo)
                
                $gpoForm.Controls.Add($txtReport); $gpoForm.Controls.Add($pnlBottom); $gpoForm.CancelButton = $btnNo
                
                if ($gpoForm.ShowDialog() -ne [System.Windows.Forms.DialogResult]::Yes) {
                    return
                }
            }
        }
    }

    # Validation: Check for duplicates
    $duplicates = $MappingData | Group-Object TargetDN | Where-Object { $_.Count -gt 1 }
    if ($duplicates) {
        $msg = "Duplicate Target DNs detected. Please resolve before saving:`n" + ($duplicates.Name -join "`n")
        [System.Windows.Forms.MessageBox]::Show($msg, "Validation Error", "OK", "Error")
        return
    }

    # Sort by length to ensure parents created first
    $MappingData = $MappingData | Sort-Object { $_.TargetDN.Length }

    $outFile = Join-Path $MapPath "OU_Map_Draft.csv"
    $MappingData | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

    [System.Windows.Forms.MessageBox]::Show("Mapping saved to:`n$outFile", "Success", "OK", "Information")
})

# Search Handlers
$btnSourceSearch.Add_Click({ Search-Tree $treeSource $txtSourceSearch.Text })
$txtSourceSearch.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { Search-Tree $treeSource $txtSourceSearch.Text; $_.SuppressKeyPress = $true } })

$btnTargetSearch.Add_Click({ Search-Tree $treeTarget $txtTargetSearch.Text })
$txtTargetSearch.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { Search-Tree $treeTarget $txtTargetSearch.Text; $_.SuppressKeyPress = $true } })

$treeTarget.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq 'Z') {
        Undo-Action; $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq 'C') {
        Copy-Action; $_.SuppressKeyPress = $true
    }
    elseif ($_.Control -and $_.KeyCode -eq 'V') {
        Invoke-PasteAction; $_.SuppressKeyPress = $true
    }
    elseif ($_.KeyCode -eq 'Delete') {
        Remove-Action; $_.SuppressKeyPress = $true
    }
})

# --- SHOW ---

$form.Add_Shown({
    $form.Activate()
})

$form.ShowDialog() | Out-Null
