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

# --- GUI SETUP ---

$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Migration - OU Mapper"
$form.Size = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition = "CenterScreen"

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

$treeSource = New-Object System.Windows.Forms.TreeView
$treeSource.Dock = "Fill"
$treeSource.AllowDrop = $false # Source is drag source only
$splitContainer.Panel1.Controls.Add($treeSource)

# Right Pane (Target)
$lblTarget = New-Object System.Windows.Forms.Label
$lblTarget.Text = "Target Domain (Drag here to Map)"
$lblTarget.Dock = "Top"
$splitContainer.Panel2.Controls.Add($lblTarget)

$treeTarget = New-Object System.Windows.Forms.TreeView
$treeTarget.Dock = "Fill"
$treeTarget.AllowDrop = $true
$treeTarget.LabelEdit = $true # Allow renaming
$splitContainer.Panel2.Controls.Add($treeTarget)

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

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Drag OUs from Left to Right. Right-click Target nodes to Rename/Delete."
$lblInfo.AutoSize = $true
$lblInfo.Location = New-Object System.Drawing.Point(10, 15)
$panelBottom.Controls.Add($lblInfo)

# --- FUNCTIONS ---

function Add-NodeToTree ($tree, $dn, $tagData) {
    # Parse DN to find hierarchy
    # Simple parser: assumes standard comma separation. 
    # For robust DN parsing, specific regex is needed, but this suffices for visualization.
    
    # We only care about OUs and DCs.
    $parts = $dn -split "(?<!\\)," | ForEach-Object { $_.Trim() }
    # Reverse to build from root up
    [array]::Reverse($parts)
    
    $currentNode = $null
    $nodesCollection = $tree.Nodes

    foreach ($part in $parts) {
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
            if ($part -eq $parts[-1] -and $tagData) {
                $currentNode.Tag = $tagData
                $currentNode.ForeColor = "DarkBlue"
            }

            $nodesCollection.Add($currentNode) | Out-Null
        }
        $nodesCollection = $currentNode.Nodes
    }
    return $currentNode
}

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

# Check for existing draft
$draftFile = Join-Path $MapPath "OU_Map_Draft.csv"
if (Test-Path $draftFile) {
    $Draft = Import-Csv $draftFile
    foreach ($row in $Draft) {
        if ($row.Action -ne 'Skip') {
            $tag = @{
                SourceDN = $row.SourceDN
                Description = $row.Description
                SourceOU = $row.SourceOU
            }
            Add-NodeToTree -tree $treeTarget -dn $row.TargetDN -tagData $tag | Out-Null
        }
    }
} else {
    # Default Root
    Add-NodeToTree -tree $treeTarget -dn $TargetDomain -tagData $null | Out-Null
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
$itemDelete.Add_Click({
    if ($treeTarget.SelectedNode) {
        $treeTarget.SelectedNode.Remove()
    }
})
$itemRename = $ctxMenu.MenuItems.Add("Rename")
$itemRename.Add_Click({
    if ($treeTarget.SelectedNode) {
        $treeTarget.SelectedNode.BeginEdit()
    }
})
$itemNew = $ctxMenu.MenuItems.Add("New OU")
$itemNew.Add_Click({
    if ($treeTarget.SelectedNode) {
        $newNode = New-Object System.Windows.Forms.TreeNode
        $newNode.Text = "NewOU"
        $newNode.Name = "OU=NewOU"
        $treeTarget.SelectedNode.Nodes.Add($newNode) | Out-Null
        $treeTarget.SelectedNode.Expand()
        $newNode.BeginEdit()
    }
})

$treeTarget.ContextMenu = $ctxMenu

# Rename Validation
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
            
            if ($node.Text -match "^OU=" -or $node.Parent) {
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

# --- SHOW ---

$form.Add_Shown({
    $form.Activate()
})

$form.ShowDialog() | Out-Null
