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

        if ($FoundPrivilegedGroups.Count -gt 0) {
            # Build and show a custom form
            $form = New-Object System.Windows.Forms.Form
            $form.Text = "High-Privilege Group Memberships"
            $form.Size = New-Object System.Drawing.Size(500, 300)
            $form.StartPosition = "CenterScreen"

            $label = New-Object System.Windows.Forms.Label; $label.Text = "The following high-privilege groups were found. How should their memberships be handled?"; $label.Location = New-Object System.Drawing.Point(10, 10); $label.Size = New-Object System.Drawing.Size(460, 30); $form.Controls.Add($label)
            $listBox = New-Object System.Windows.Forms.ListBox; $listBox.Location = New-Object System.Drawing.Point(10, 45); $listBox.Size = New-Object System.Drawing.Size(460, 120); $form.Controls.Add($listBox)
            $FoundPrivilegedGroups | ForEach-Object { [void]$listBox.Items.Add($_) }

            $btnKeep = New-Object System.Windows.Forms.Button; $btnKeep.Text = "Keep Privileges"; $btnKeep.DialogResult = [System.Windows.Forms.DialogResult]::Yes; $btnKeep.Location = New-Object System.Drawing.Point(10, 180); $btnKeep.Size = New-Object System.Drawing.Size(150, 40); $form.Controls.Add($btnKeep)
            $btnRemove = New-Object System.Windows.Forms.Button; $btnRemove.Text = "Remove Privileges"; $btnRemove.DialogResult = [System.Windows.Forms.DialogResult]::No; $btnRemove.Location = New-Object System.Drawing.Point(170, 180); $btnRemove.Size = New-Object System.Drawing.Size(150, 40); $form.Controls.Add($btnRemove)
            $btnCustomize = New-Object System.Windows.Forms.Button; $btnCustomize.Text = "Customize..."; $btnCustomize.DialogResult = [System.Windows.Forms.DialogResult]::Retry; $btnCustomize.Location = New-Object System.Drawing.Point(330, 180); $btnCustomize.Size = New-Object System.Drawing.Size(150, 40); $form.Controls.Add($btnCustomize)

            $result = $form.ShowDialog()

            if ($result -eq 'Yes') { # Keep All
                Write-Log -Message "User chose to KEEP all high-privilege group memberships." -Level WARN
                $RestrictedGroups = @() # Empty the list
            }
            elseif ($result -eq 'No') { # Remove All
                Write-Log -Message "User chose to REMOVE all high-privilege group memberships (Default)." -Level INFO
                # $RestrictedGroups is already set to the default, so do nothing.
            }
            elseif ($result -eq 'Retry') { # Customize
                $customForm = New-Object System.Windows.Forms.Form; $customForm.Text = "Customize Privileges"; $customForm.Size = New-Object System.Drawing.Size(400, 450); $customForm.StartPosition = "CenterScreen"
                $customLabel = New-Object System.Windows.Forms.Label; $customLabel.Text = "Check the groups whose memberships you want to REMOVE from the export."; $customLabel.Location = New-Object System.Drawing.Point(10, 10); $customLabel.Size = New-Object System.Drawing.Size(360, 30); $customForm.Controls.Add($customLabel)
                $checkedListBox = New-Object System.Windows.Forms.CheckedListBox; $checkedListBox.Location = New-Object System.Drawing.Point(10, 45); $checkedListBox.Size = New-Object System.Drawing.Size(360, 300); $customForm.Controls.Add($checkedListBox)
                $FoundPrivilegedGroups | ForEach-Object { [void]$checkedListBox.Items.Add($_, $true) }
                $btnOk = New-Object System.Windows.Forms.Button; $btnOk.Text = "OK"; $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK; $btnOk.Location = New-Object System.Drawing.Point(290, 360); $customForm.Controls.Add($btnOk); $customForm.AcceptButton = $btnOk

                if ($customForm.ShowDialog() -eq 'OK') {
                    $RestrictedGroups = $checkedListBox.CheckedItems
                    Write-Log -Message "User customized privilege removal. Groups to be skipped: $($RestrictedGroups -join ', ')" -Level INFO
                } else {
                    Write-Log -Message "User cancelled customization. Defaulting to REMOVE all high-privilege group memberships." -Level INFO
                }
            }
        }

        # Always skip Domain Users as it's auto-assigned
        $RestrictedGroups += "Domain Users"

        # Export Group Memberships (excluding high-privilege groups)
        Write-Log -Message "Exporting group memberships..." -Level INFO
        
        $Memberships = @()
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
    Write-Host "Account data export failed. Check logs for details."
}
