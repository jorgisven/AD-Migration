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
        $TargetServer = ($TargetDomain -replace 'DC=','' -split '(?<!\\),' | Join-Object -Separator '.')
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
        if ($row.Action -ne 'Skip') {
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

    # Handle MERGE action (dragging from source onto an existing target node)
    if ($targetNode -and $draggedNode -and ($draggedNode.TreeView -eq $treeSource)) {
        # A node with a Tag is either pre-existing or already mapped.
        if ($targetNode.Tag) {
            $msg = "You are dropping onto an existing target OU.`n`nSource: '$($draggedNode.Text)'`nTarget: '$($targetNode.Text)'`n`nWould you like to:`n`n[Yes] Map the source OU to this target OU (Merge).`n[No] Create the source OU as a new child of the target OU.`n[Cancel] Abort the operation."
            $result = [System.Windows.Forms.MessageBox]::Show($msg, "Merge or Add as Child?", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
            
            if ($result -eq 'Yes') {
                # Perform the merge: copy the tag from source to target.
                $targetNode.Tag = $draggedNode.Tag
                $targetNode.ForeColor = [System.Drawing.Color]::DarkBlue # Mark as mapped
                return # The drop action is complete.
            } elseif ($result -eq 'Cancel') {
                return # Cancel the drop entirely
            }
            # If user says No, we fall through to the default "add as child" behavior.
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
        $newNode.Text = "OU=NewOU"
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

# Search Handlers
$btnSourceSearch.Add_Click({ Search-Tree $treeSource $txtSourceSearch.Text })
$txtSourceSearch.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { Search-Tree $treeSource $txtSourceSearch.Text; $_.SuppressKeyPress = $true } })

$btnTargetSearch.Add_Click({ Search-Tree $treeTarget $txtTargetSearch.Text })
$txtTargetSearch.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { Search-Tree $treeTarget $txtTargetSearch.Text; $_.SuppressKeyPress = $true } })

# --- SHOW ---

$form.Add_Shown({
    $form.Activate()
})

$form.ShowDialog() | Out-Null
