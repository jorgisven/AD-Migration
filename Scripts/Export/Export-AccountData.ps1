<#
.SYNOPSIS
    Export user and service account attributes.

.DESCRIPTION
    Used for account reconciliation between source and target domains.
    Exports user and computer attributes needed for migration planning.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
# build path to the module manifest (go up two levels to repo root, then into Scripts\ADMigration)
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) {
    throw "ADMigration module manifest missing, cannot continue."
}
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
# sanity-check that key helpers are available
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) {
    Write-Log -Message "Required function Invoke-Safely not defined after module import" -Level ERROR
    throw "Invoke-Safely unavailable"
}
$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'Security'

# Ensure export directory exists
if (-not (Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
    Write-Log -Message "Created export directory: $ExportPath" -Level INFO
}

# Prompt for source domain if not provided
if (-not $SourceDomain) {
    $SourceDomain = Read-Host "Enter source domain name (e.g., source.local)"
}

Write-Log -Message "Starting account data export from domain: $SourceDomain" -Level INFO

# GUI Prompt for Export Scope (Filter)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

$SearchBase = $null
$scopeResult = [System.Windows.Forms.MessageBox]::Show("Do you want to export objects from the ENTIRE domain?`n`nYes = Entire Domain (Default)`nNo = Filter by specific OU (SearchBase)", "Export Scope", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

if ($scopeResult -eq 'No') {
    $inputOU = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Distinguished Name (DN) of the OU to export from:`n(e.g. OU=Sales,DC=source,DC=local)", "Filter by OU", "")
    if (-not [string]::IsNullOrWhiteSpace($inputOU)) {
        $SearchBase = $inputOU
        Write-Log -Message "Export scoped to OU: $SearchBase" -Level INFO
    } else {
        Write-Log -Message "No OU provided, defaulting to entire domain." -Level WARN
    }
} else {
    Write-Log -Message "Exporting from entire domain." -Level INFO
}

try {
    Invoke-Safely -ScriptBlock {
        # Export user accounts
        Write-Log -Message "Exporting user accounts from $SourceDomain" -Level INFO
        $userProps = 'DisplayName', 'GivenName', 'Surname', 'Enabled', 'LastLogonDate', 'PasswordLastSet', 'WhenCreated', 'WhenChanged', 'AccountExpirationDate', 'UserPrincipalName', 'SamAccountName', 'DistinguishedName', 'Description'
        
        $userParams = @{
            Filter = "*"
            Server = $SourceDomain
            Properties = $userProps
        }
        if ($SearchBase) { $userParams['SearchBase'] = $SearchBase }

        $Users = Get-ADUser @userParams | `
            Select-Object @{Name = 'SamAccountName'; Expression = { $_.SamAccountName }},
                          @{Name = 'SID'; Expression = { $_.SID.Value }},
                          @{Name = 'UserPrincipalName'; Expression = { $_.UserPrincipalName }},
                          @{Name = 'DisplayName'; Expression = { $_.DisplayName }},
                          @{Name = 'GivenName'; Expression = { $_.GivenName }},
                          @{Name = 'Surname'; Expression = { $_.Surname }},
                          @{Name = 'Description'; Expression = { $_.Description }},
                          @{Name = 'DistinguishedName'; Expression = { $_.DistinguishedName }},
                          @{Name = 'Enabled'; Expression = { $_.Enabled }},
                          @{Name = 'LastLogonDate'; Expression = { $_.LastLogonDate }},
                          @{Name = 'PasswordLastSet'; Expression = { $_.PasswordLastSet }},
                          @{Name = 'WhenCreated'; Expression = { $_.WhenCreated }},
                          @{Name = 'WhenChanged'; Expression = { $_.WhenChanged }},
                          @{Name = 'AccountExpirationDate'; Expression = { $_.AccountExpirationDate }} | `
            Sort-Object SamAccountName
        
        $userFile = Join-Path $ExportPath "Users_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Users | Export-Csv -Path $userFile -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Exported $($Users.Count) user accounts to $userFile" -Level INFO
        
        # Export service accounts (users with specific naming patterns)
        Write-Log -Message "Exporting service accounts from $SourceDomain" -Level INFO
        
        $svcParams = @{
            Filter = { (SamAccountName -like 'svc_*') -or (SamAccountName -like '*service*') }
            Server = $SourceDomain
            Properties = 'DisplayName', 'UserPrincipalName', 'Enabled', 'userAccountControl', 'Description'
        }
        if ($SearchBase) { $svcParams['SearchBase'] = $SearchBase }

        $ServiceAccounts = Get-ADUser @svcParams | `
            Select-Object @{Name = 'SamAccountName'; Expression = { $_.SamAccountName }},
                          @{Name = 'SID'; Expression = { $_.SID.Value }},
                          @{Name = 'UserPrincipalName'; Expression = { $_.UserPrincipalName }},
                          @{Name = 'DisplayName'; Expression = { $_.DisplayName }},
                          @{Name = 'Description'; Expression = { $_.Description }},
                          @{Name = 'DistinguishedName'; Expression = { $_.DistinguishedName }},
                          @{Name = 'Enabled'; Expression = { $_.Enabled }},
                          @{Name = 'PasswordNotRequired'; Expression = { ($_.userAccountControl -band 32) -eq 32 }},
                          @{Name = 'AccountNeverExpires'; Expression = { ($_.userAccountControl -band 65536) -eq 65536 }} | `
            Sort-Object SamAccountName
        
        if ($ServiceAccounts.Count -gt 0) {
            $svcFile = Join-Path $ExportPath "ServiceAccounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $ServiceAccounts | Export-Csv -Path $svcFile -NoTypeInformation -Encoding UTF8
            Write-Log -Message "Exported $($ServiceAccounts.Count) service accounts to $svcFile" -Level INFO
        }
        
        # Export computer accounts
        Write-Log -Message "Exporting computer accounts from $SourceDomain" -Level INFO
        $compProps = 'OperatingSystem', 'OperatingSystemVersion', 'Enabled', 'LastLogonDate', 'WhenCreated', 'WhenChanged', 'Description'
        # Filter out Domain Controllers (PrimaryGroupID 516)
        
        $compParams = @{
            Filter = "PrimaryGroupID -ne 516"
            Server = $SourceDomain
            Properties = $compProps
        }
        if ($SearchBase) { $compParams['SearchBase'] = $SearchBase }

        $Computers = Get-ADComputer @compParams | `
            Select-Object @{Name = 'ComputerName'; Expression = { $_.Name }},
                          @{Name = 'SID'; Expression = { $_.SID.Value }},
                          @{Name = 'SamAccountName'; Expression = { $_.SamAccountName }},
                          @{Name = 'Description'; Expression = { $_.Description }},
                          @{Name = 'DistinguishedName'; Expression = { $_.DistinguishedName }},
                          @{Name = 'OperatingSystem'; Expression = { $_.OperatingSystem }},
                          @{Name = 'OperatingSystemVersion'; Expression = { $_.OperatingSystemVersion }},
                          @{Name = 'Enabled'; Expression = { $_.Enabled }},
                          @{Name = 'LastLogonDate'; Expression = { $_.LastLogonDate }},
                          @{Name = 'WhenCreated'; Expression = { $_.WhenCreated }},
                          @{Name = 'WhenChanged'; Expression = { $_.WhenChanged }} | `
            Sort-Object ComputerName
        
        $computerFile = Join-Path $ExportPath "Computers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Computers | Export-Csv -Path $computerFile -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Exported $($Computers.Count) computer accounts to $computerFile" -Level INFO
        
        # Export groups
        Write-Log -Message "Exporting groups from $SourceDomain" -Level INFO
        
        $groupParams = @{
            Filter = "*"
            Server = $SourceDomain
            Properties = 'Description'
        }
        if ($SearchBase) { $groupParams['SearchBase'] = $SearchBase }

        $Groups = Get-ADGroup @groupParams | `
            Select-Object @{Name = 'Name'; Expression = { $_.Name }},
                          @{Name = 'SamAccountName'; Expression = { $_.SamAccountName }},
                          @{Name = 'SID'; Expression = { $_.SID.Value }},
                          @{Name = 'Description'; Expression = { $_.Description }},
                          @{Name = 'DistinguishedName'; Expression = { $_.DistinguishedName }},
                          @{Name = 'GroupCategory'; Expression = { $_.GroupCategory }},
                          @{Name = 'GroupScope'; Expression = { $_.GroupScope }} | `
            Sort-Object Name
            
        $groupFile = Join-Path $ExportPath "Groups_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Groups | Export-Csv -Path $groupFile -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Exported $($Groups.Count) groups to $groupFile" -Level INFO

        # --- Handle High-Privilege Group Memberships ---

        # Define default restricted groups
        $DefaultRestrictedGroups = @(
            "Domain Admins", "Enterprise Admins", "Schema Admins", "Administrators",
            "Account Operators", "Backup Operators", "Print Operators", "Server Operators",
            "Group Policy Creator Owners"
        )
        # This will be the final list of groups to skip
        $RestrictedGroups = $DefaultRestrictedGroups

        # Find which of these groups actually exist in the source domain
        $FoundPrivilegedGroups = $Groups | Where-Object { $_.Name -in $DefaultRestrictedGroups } | Select-Object -ExpandProperty Name

        # Pre-scan privileged groups for members to filter out empty ones and prepare for GUI
        $PrivilegedGroupData = @{}
        foreach ($pGroup in $FoundPrivilegedGroups) {
            $gObj = $Groups | Where-Object { $_.Name -eq $pGroup } | Select-Object -First 1
            try {
                $members = Get-ADGroupMember -Identity $gObj.DistinguishedName -Server $SourceDomain -ErrorAction SilentlyContinue
                if ($members) {
                    $PrivilegedGroupData[$pGroup] = @{
                        GroupObject = $gObj
                        Members = $members
                    }
                }
            } catch {}
        }

        $GroupsWithMembers = $PrivilegedGroupData.Keys | Sort-Object
        $KeptPrivilegedMemberships = @()

        if ($GroupsWithMembers.Count -gt 0) {
            # Build and show a custom form
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "High-Privilege Group Memberships"
            $form.Size = New-Object System.Drawing.Size(500, 510)
            $form.StartPosition = "CenterScreen"

            $label = New-Object System.Windows.Forms.Label
            $label.Text = "The following high-privilege groups have members. Select which memberships to KEEP in the export. For accounts that already exist in the target domain, these group additions are additive only—existing permissions will not be removed, only elevated if selected here."
            $label.Location = New-Object System.Drawing.Point(10, 10)
            $label.Size = New-Object System.Drawing.Size(460, 30)
            $form.Controls.Add($label)

            $noteLabel = New-Object System.Windows.Forms.Label

            $noteLabel.Text = "Note: If an account is listed here that you plan to exclude from migration, group membership changes made here will have no effect. Only accounts that are actually migrated will have their group memberships and permissions processed. Review your account mapping to ensure consistency."
            $noteLabel.Location = New-Object System.Drawing.Point(10, 40)
            $noteLabel.Size = New-Object System.Drawing.Size(460, 40)
            $noteLabel.ForeColor = [System.Drawing.Color]::DarkRed
            $noteLabel.Font = New-Object System.Drawing.Font($label.Font, [System.Drawing.FontStyle]::Italic)
            $form.Controls.Add($noteLabel)

            $builtinLabel = New-Object System.Windows.Forms.Label
            $builtinLabel.Text = "Best Practice: Built-in accounts (e.g., Administrator, Guest) should usually NOT have their group memberships imported. These accounts exist by default in every domain, and their memberships are managed by the system. Only include them if you have a specific, documented need."
            $builtinLabel.Location = New-Object System.Drawing.Point(10, 80)
            $builtinLabel.Size = New-Object System.Drawing.Size(460, 45)
            $builtinLabel.ForeColor = [System.Drawing.Color]::DarkBlue
            $builtinLabel.Font = New-Object System.Drawing.Font($label.Font, [System.Drawing.FontStyle]::Italic)
            $form.Controls.Add($builtinLabel)

            $treeView = New-Object System.Windows.Forms.TreeView
            $treeView.Location = New-Object System.Drawing.Point(10, 130)
            $treeView.Size = New-Object System.Drawing.Size(460, 265)
            $treeView.CheckBoxes = $true
            $form.Controls.Add($treeView)
            
            foreach ($grpName in $GroupsWithMembers) {
                $node = $treeView.Nodes.Add($grpName)
                foreach ($member in $PrivilegedGroupData[$grpName].Members) {
                    # Filter out default nested groups to avoid confusion (e.g. Domain Admins inside Administrators)
                    # These links exist by default in the target and unchecking them here won't break that default inheritance.
                    if ($grpName -eq "Administrators" -and $member.SamAccountName -in @("Domain Admins", "Enterprise Admins")) {
                        continue
                    }
                    $mNode = $node.Nodes.Add("$($member.Name) ($($member.SamAccountName))")
                    $mNode.Tag = $member
                }
            }
            
            # Cascade check events
            $treeView.Add_AfterCheck({
                param($sourceControl, $e)
                if ($e.Action -ne 'Unknown') { 
                    foreach ($child in $e.Node.Nodes) {
                        $child.Checked = $e.Node.Checked
                    }
                }
            })

            $btnKeep = New-Object System.Windows.Forms.Button; $btnKeep.Text = "Export Selected"; $btnKeep.DialogResult = [System.Windows.Forms.DialogResult]::Yes; $btnKeep.Location = New-Object System.Drawing.Point(10, 360); $btnKeep.Size = New-Object System.Drawing.Size(220, 40); $form.Controls.Add($btnKeep)
            $btnRemove = New-Object System.Windows.Forms.Button; $btnRemove.Text = "Skip All (Secure)"; $btnRemove.DialogResult = [System.Windows.Forms.DialogResult]::No; $btnRemove.Location = New-Object System.Drawing.Point(250, 360); $btnRemove.Size = New-Object System.Drawing.Size(220, 40); $form.Controls.Add($btnRemove)

            $result = $form.ShowDialog()

            if ($result -eq 'Yes') { # Export Selected
                foreach ($gNode in $treeView.Nodes) {
                     $gName = $gNode.Text
                     $gObj = $PrivilegedGroupData[$gName].GroupObject
                     
                     foreach ($mNode in $gNode.Nodes) {
                         if ($mNode.Checked) {
                             $mem = $mNode.Tag
                             $KeptPrivilegedMemberships += [PSCustomObject]@{
                                GroupSam   = $gObj.SamAccountName
                                MemberSam  = $mem.SamAccountName
                                MemberType = $mem.objectClass
                             }
                         }
                     }
                 }
                 Write-Log -Message "User selected specific high-privilege memberships to keep." -Level WARN
            } else {
                Write-Log -Message "User chose to REMOVE all high-privilege group memberships." -Level INFO
            }
        }

        # Always skip Domain Users as it's auto-assigned
        $RestrictedGroups += "Domain Users"

        # Export Group Memberships (excluding high-privilege groups)
        Write-Log -Message "Exporting group memberships..." -Level INFO
        
        $Memberships = @()
        
        # Add the kept privileged ones first
        $Memberships += $KeptPrivilegedMemberships
        
        foreach ($grp in $Groups) {
            if ($grp.Name -in $RestrictedGroups) {
                Write-Log -Message "Skipping membership export for restricted/default group: $($grp.Name)" -Level INFO
                continue
            }
            
            try {
                # Get members (Direct membership only to preserve structure)
                $members = Get-ADGroupMember -Identity $grp.DistinguishedName -Server $SourceDomain -ErrorAction SilentlyContinue
                foreach ($member in $members) {
                    $Memberships += [PSCustomObject]@{
                        GroupSam   = $grp.SamAccountName
                        MemberSam  = $member.SamAccountName
                        MemberType = $member.objectClass
                    }
                }
            } catch {
                Write-Log -Message "Failed to retrieve members for group $($grp.Name): $_" -Level WARN
            }
        }

        $memberFile = Join-Path $ExportPath "GroupMembers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Memberships | Export-Csv -Path $memberFile -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Exported $($Memberships.Count) membership entries to $memberFile" -Level INFO

        # Summary
        $summary = @"
=== Account Export Summary ===
Domain: $SourceDomain
Exported: $(Get-Date)

Users: $($Users.Count)
Service Accounts: $($ServiceAccounts.Count)
Computers: $($Computers.Count)
Groups: $($Groups.Count)
Memberships: $($Memberships.Count)

Total Objects: $($Users.Count + $ServiceAccounts.Count + $Computers.Count + $Groups.Count + $Memberships.Count)
"@
        
        Add-Content -Path (Join-Path $ExportPath 'EXPORT_SUMMARY.txt') -Value $summary -Encoding UTF8
        
        Write-Host "Account data export complete:"
        Write-Host "  - Users: $($Users.Count)"
        Write-Host "  - Service Accounts: $($ServiceAccounts.Count)"
        Write-Host "  - Computers: $($Computers.Count)"
        Write-Host "  - Groups: $($Groups.Count)"
        Write-Host "  - Memberships: $($Memberships.Count)"
        
    } -Operation "Export account data from $SourceDomain"
    
} catch {
    Write-Log -Message "Failed to export account data: $_" -Level ERROR
    throw "Account data export failed. Check logs for details."
}
