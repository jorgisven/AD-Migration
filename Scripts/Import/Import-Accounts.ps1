<#
.SYNOPSIS
    Imports mapped accounts (Users, Computers, Groups) into the target domain.

.DESCRIPTION
    Reads the Account Mapping CSVs and creates the designated objects in the new OUs.
    Restores attributes from the original Export data and assigns a default password to users.
    Finally, restores group memberships for mapped accounts.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

$config = Get-ADMigrationConfig
$ExportSecurityPath = Join-Path $config.ExportRoot 'Security'
$MapPath = Join-Path $config.TransformRoot 'Mapping'

if (-not $TargetDomain) { $TargetDomain = (Get-ADDomain).DNSRoot }

Write-Log -Message "Starting Account Import to $TargetDomain..." -Level INFO

# --- 1. Load Mappings ---
$userMapFile = Join-Path $MapPath "User_Account_Map.csv"
$compMapFile = Join-Path $MapPath "Computer_Account_Map.csv"
$groupMapFile = Join-Path $MapPath "Group_Account_Map.csv"

if (-not (Test-Path $userMapFile)) { throw "Account Mapping files not found. Run Transform-AccountMapping first." }

$UserMap = Import-Csv $userMapFile
$CompMap = Import-Csv $compMapFile
$GroupMap = Import-Csv $groupMapFile

# Build Translation Dictionary for Memberships (SourceSam -> TargetSam)
$IdentityMap = @{}
foreach ($u in $UserMap) { if ($u.Action -ne 'Skip') { $IdentityMap[$u.SourceSam] = $u.TargetSam } }
foreach ($c in $CompMap) { if ($c.Action -ne 'Skip') { $IdentityMap[$c.SourceName] = $c.TargetName } }
foreach ($g in $GroupMap) { if ($g.Action -ne 'Skip') { $IdentityMap[$g.SourceSam] = $g.TargetSam } }

# --- 2. Load Exported Attributes ---
function Get-LatestExport ($prefix) {
    $file = Get-ChildItem -Path $ExportSecurityPath -Filter "${prefix}_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($file) { return Import-Csv $file.FullName }
    return @()
}

$exportedUsers = Get-LatestExport "Users" | Group-Object -AsHashTable -Property SamAccountName
$exportedGroups = Get-LatestExport "Groups" | Group-Object -AsHashTable -Property SamAccountName
$exportedMembers = Get-LatestExport "GroupMembers"

# --- 3. Prompt for Default Password ---
$usersToCreate = $UserMap | Where-Object { $_.Action -eq 'Create' }
$securePassword = $null

if ($usersToCreate.Count -gt 0) {
    Write-Host "`nThere are $($usersToCreate.Count) new users to create." -ForegroundColor Cyan
    Write-Host "Please enter a default initial password. (Users will be forced to change this at first logon)." -ForegroundColor Yellow
    $securePassword = Read-Host "Default Password" -AsSecureString
    if (-not $securePassword) { throw "A default password is required to create new user accounts." }
}

# --- 4. Import Users ---
Invoke-Safely -ScriptBlock {
    Write-Host "`n--- Importing Users ---" -ForegroundColor Cyan
    $createdUserCount = 0
    $skippedExistingUserCount = 0
    foreach ($u in $usersToCreate) {
        $srcUser = $exportedUsers[$u.SourceSam]
        $targetSam = $u.TargetSam
        $targetOU = $u.TargetOU_DN
        $upn = "$targetSam@$TargetDomain"

        $existingUser = Get-ADUser -Identity $targetSam -Server $TargetDomain -ErrorAction SilentlyContinue
        if ($existingUser) {
            $skippedExistingUserCount++
            Write-Log -Message "Idempotency skip: User '$targetSam' already exists in target domain. Creation skipped." -Level WARN
            continue
        }

        $userParams = @{
            Name                  = $targetSam
            SamAccountName        = $targetSam
            UserPrincipalName     = $upn
            Path                  = $targetOU
            AccountPassword       = $securePassword
            ChangePasswordAtLogon = $true
            Enabled               = $true
            Server                = $TargetDomain
        }
        
        # Add rich attributes if they existed in the export
        if ($srcUser) {
            if ($srcUser.GivenName) { $userParams.GivenName = $srcUser.GivenName }
            if ($srcUser.Surname) { $userParams.Surname = $srcUser.Surname }
            if ($srcUser.DisplayName) { $userParams.DisplayName = $srcUser.DisplayName }
            if ($srcUser.EmailAddress) { $userParams.EmailAddress = $srcUser.EmailAddress }
            if ($srcUser.Description) { $userParams.Description = $srcUser.Description }
        }

        try {
            if ($PSCmdlet.ShouldProcess($targetSam, "Create User in $targetOU")) {
                New-ADUser @userParams -ErrorAction Stop
                Write-Log -Message "Created User: $targetSam" -Level INFO
                $createdUserCount++
            }
        } catch {
            Write-Log -Message "Failed to create user '$targetSam': $($_.Exception.Message)" -Level ERROR
        }
    }

    Write-Log -Message "User import summary: created=$createdUserCount; skippedExisting=$skippedExistingUserCount." -Level INFO
    if ($skippedExistingUserCount -gt 0) {
        Write-Host "[!] Users: Skipped $skippedExistingUserCount existing account(s). See log for details." -ForegroundColor Yellow
    }
} -Operation "Import Users"

# --- 5. Import Computers ---
Invoke-Safely -ScriptBlock {
    Write-Host "`n--- Importing Computers ---" -ForegroundColor Cyan
    $compsToCreate = $CompMap | Where-Object { $_.Action -eq 'Create' }
    $createdComputerCount = 0
    $skippedExistingComputerCount = 0
    foreach ($c in $compsToCreate) {
        $targetName = $c.TargetName
        $targetOU = $c.TargetOU_DN

        $existingComputer = Get-ADComputer -Identity $targetName -Server $TargetDomain -ErrorAction SilentlyContinue
        if ($existingComputer) {
            $skippedExistingComputerCount++
            Write-Log -Message "Idempotency skip: Computer '$targetName' already exists in target domain. Creation skipped." -Level WARN
            continue
        }

        $compParams = @{
            Name           = $targetName
            SamAccountName = "$targetName$"
            Path           = $targetOU
            Enabled        = $true
            Server         = $TargetDomain
        }
        if ($c.Description) { $compParams.Description = $c.Description }

        try {
            if ($PSCmdlet.ShouldProcess($targetName, "Create Computer in $targetOU")) {
                New-ADComputer @compParams -ErrorAction Stop
                Write-Log -Message "Created Computer: $targetName" -Level INFO
                $createdComputerCount++
            }
        } catch {
            Write-Log -Message "Failed to create computer '$targetName': $($_.Exception.Message)" -Level ERROR
        }
    }

    Write-Log -Message "Computer import summary: created=$createdComputerCount; skippedExisting=$skippedExistingComputerCount." -Level INFO
    if ($skippedExistingComputerCount -gt 0) {
        Write-Host "[!] Computers: Skipped $skippedExistingComputerCount existing account(s). See log for details." -ForegroundColor Yellow
    }
} -Operation "Import Computers"

# --- 6. Import Groups ---
Invoke-Safely -ScriptBlock {
    Write-Host "`n--- Importing Groups ---" -ForegroundColor Cyan
    $groupsToCreate = $GroupMap | Where-Object { $_.Action -eq 'Create' }
    $createdGroupCount = 0
    $skippedExistingGroupCount = 0
    foreach ($g in $groupsToCreate) {
        $srcGroup = $exportedGroups[$g.SourceSam]
        $targetSam = $g.TargetSam
        $targetOU = $g.TargetOU_DN

        $existingGroup = Get-ADGroup -Identity $targetSam -Server $TargetDomain -ErrorAction SilentlyContinue
        if ($existingGroup) {
            $skippedExistingGroupCount++
            Write-Log -Message "Idempotency skip: Group '$targetSam' already exists in target domain. Creation skipped." -Level WARN
            continue
        }

        $grpParams = @{
            Name           = $targetSam
            SamAccountName = $targetSam
            Path           = $targetOU
            Server         = $TargetDomain
        }
        
        # Default to Security/Global if not found in export
        if ($srcGroup -and $srcGroup.GroupCategory) { $grpParams.GroupCategory = $srcGroup.GroupCategory } else { $grpParams.GroupCategory = 'Security' }
        if ($srcGroup -and $srcGroup.GroupScope) { $grpParams.GroupScope = $srcGroup.GroupScope } else { $grpParams.GroupScope = 'Global' }
        if ($srcGroup -and $srcGroup.Description) { $grpParams.Description = $srcGroup.Description }

        try {
            if ($PSCmdlet.ShouldProcess($targetSam, "Create Group in $targetOU")) {
                New-ADGroup @grpParams -ErrorAction Stop
                Write-Log -Message "Created Group: $targetSam" -Level INFO
                $createdGroupCount++
            }
        } catch {
            Write-Log -Message "Failed to create group '$targetSam': $($_.Exception.Message)" -Level ERROR
        }
    }

    Write-Log -Message "Group import summary: created=$createdGroupCount; skippedExisting=$skippedExistingGroupCount." -Level INFO
    if ($skippedExistingGroupCount -gt 0) {
        Write-Host "[!] Groups: Skipped $skippedExistingGroupCount existing group(s). See log for details." -ForegroundColor Yellow
    }
} -Operation "Import Groups"

# --- 7. Restore Memberships ---
Invoke-Safely -ScriptBlock {
    if ($exportedMembers) {
        Write-Host "`n--- Restoring Group Memberships ---" -ForegroundColor Cyan
        $addedCount = 0
        $skippedCount = 0
        
        foreach ($m in $exportedMembers) {
            # Depending on how Export-AccountData exports it, we look for Group/Member names
            $srcGroup = if ($m.GroupSam) { $m.GroupSam }
                        elseif ($m.GroupSamAccountName) { $m.GroupSamAccountName }
                        else { $m.GroupName }
            $srcMember = if ($m.MemberSam) { $m.MemberSam }
                         elseif ($m.MemberSamAccountName) { $m.MemberSamAccountName }
                         else { $m.MemberName }

            if ([string]::IsNullOrWhiteSpace($srcGroup) -or [string]::IsNullOrWhiteSpace($srcMember)) {
                $skippedCount++
                Write-Log -Message "Skipping membership row with missing source values (Group='$srcGroup', Member='$srcMember')." -Level WARN
                continue
            }
            
            # Strip $ from computer SAMs for dictionary lookup
            if ($srcMember -match '\$$') { $srcMember = $srcMember -replace '\$$','' }

            # Translate both Group and Member to Target names
            $tgtGroup = $IdentityMap[$srcGroup]
            $tgtMember = $IdentityMap[$srcMember]

            if ($tgtGroup -and $tgtMember) {
                try {
                    # Check if they are actually mapped to be migrated/merged
                    if ($PSCmdlet.ShouldProcess("Add $tgtMember to $tgtGroup", "Group Membership")) {
                        # Use SilentlyContinue to ignore "Already a member" errors gracefully
                        Add-ADGroupMember -Identity $tgtGroup -Members $tgtMember -Server $TargetDomain -ErrorAction SilentlyContinue
                        $addedCount++
                    }
                } catch {
                    Write-Log -Message "Failed to add '$tgtMember' to '$tgtGroup': $($_.Exception.Message)" -Level ERROR
                }
            } else {
                $skippedCount++
                Write-Log -Message "Skipping membership '$srcMember' -> '$srcGroup' because one or both mappings were not found (TargetMember='$tgtMember', TargetGroup='$tgtGroup')." -Level WARN
            }
        }
        Write-Log -Message "Processed $addedCount membership links; skipped $skippedCount." -Level INFO
        Write-Host "Processed $addedCount membership links; skipped $skippedCount." -ForegroundColor Green
    } else {
        Write-Log -Message "No GroupMembers export found. Skipping membership restore." -Level WARN
    }
} -Operation "Import Memberships"

Write-Host "`n=== Account Import Complete ===" -ForegroundColor Green