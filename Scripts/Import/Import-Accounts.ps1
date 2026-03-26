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
foreach ($c in $CompMap) { if ($c.Action -ne 'Skip') { $IdentityMap[$c.SourceName] = "$($c.TargetName)$" } }
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

function ConvertTo-PlainText {
    param(
        [Parameter(Mandatory = $true)]
        [Security.SecureString]$SecureString
    )

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function Get-UserNameTokens {
    param(
        [string[]]$Values
    )

    $tokens = @()
    foreach ($v in $Values) {
        if ([string]::IsNullOrWhiteSpace($v)) { continue }
        $pieces = $v -split '[^A-Za-z0-9]+'
        foreach ($p in $pieces) {
            if ($p.Length -ge 3) {
                $tokens += $p.ToLowerInvariant()
            }
        }
    }

    return @($tokens | Select-Object -Unique)
}

function Test-DefaultPasswordCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [Security.SecureString]$Password,

        [string]$SamAccountName,
        [string]$GivenName,
        [string]$Surname,
        [string]$DisplayName,

        [int]$MinPasswordLength = 7,
        [bool]$ComplexityEnabled = $true
    )

    $reasons = @()
    $passwordText = ConvertTo-PlainText -SecureString $Password

    if ($passwordText.Length -lt $MinPasswordLength) {
        $reasons += "Length below MinPasswordLength ($MinPasswordLength)."
    }

    if ($ComplexityEnabled) {
        $categoryCount = 0
        if ($passwordText -cmatch '[A-Z]') { $categoryCount++ }
        if ($passwordText -cmatch '[a-z]') { $categoryCount++ }
        if ($passwordText -match '\d') { $categoryCount++ }
        if ($passwordText -match '[^A-Za-z0-9]') { $categoryCount++ }

        if ($categoryCount -lt 3) {
            $reasons += "Complexity categories below required minimum (3 of 4)."
        }

        $passwordLower = $passwordText.ToLowerInvariant()
        $tokenInputs = @($SamAccountName, $GivenName, $Surname, $DisplayName)
        $tokens = Get-UserNameTokens -Values $tokenInputs
        foreach ($token in $tokens) {
            if ($passwordLower.Contains($token)) {
                $reasons += "Contains identity token '$token'."
            }
        }
    }

    return [PSCustomObject]@{
        IsValid = ($reasons.Count -eq 0)
        Reasons = $reasons
    }
}

# --- 3. Prompt for Default Password ---
$usersToCreate = $UserMap | Where-Object { $_.Action -eq 'Create' }
$securePassword = $null

if ($usersToCreate.Count -gt 0) {
    $effectivePolicy = $null
    $minPasswordLength = 7
    $complexityEnabled = $true
    try {
        $effectivePolicy = Get-ADDefaultDomainPasswordPolicy -Server $TargetDomain -ErrorAction Stop
        if ($effectivePolicy.MinPasswordLength -gt 0) { $minPasswordLength = [int]$effectivePolicy.MinPasswordLength }
        $complexityEnabled = [bool]$effectivePolicy.ComplexityEnabled
        Write-Log -Message "Resolved domain password policy for '$TargetDomain': ComplexityEnabled=$complexityEnabled, MinPasswordLength=$minPasswordLength." -Level INFO
    } catch {
        Write-Log -Message "Could not resolve domain password policy for '$TargetDomain'. Using fallback assumptions (ComplexityEnabled=True, MinPasswordLength=7). Details: $($_.Exception.Message)" -Level WARN
    }

    $passwordAccepted = $false
    $passwordAttempt = 0

    Write-Host "`nThere are $($usersToCreate.Count) new users to create." -ForegroundColor Cyan
    Write-Host "Please enter a default initial password. (Users will be forced to change this at first logon)." -ForegroundColor Yellow

    while (-not $passwordAccepted) {
        $passwordAttempt++
        Write-Log -Message "Default password validation attempt #$passwordAttempt started." -Level INFO

        $candidateSecure = Read-Host "Default Password" -AsSecureString
        if (-not $candidateSecure) {
            throw "A default password is required to create new user accounts."
        }

        $candidateText = ConvertTo-PlainText -SecureString $candidateSecure
        if ([string]::IsNullOrWhiteSpace($candidateText)) {
            Write-Log -Message "Default password validation attempt #$passwordAttempt failed: password was empty." -Level WARN
            continue
        }

        $problemUsers = @()
        foreach ($u in $usersToCreate) {
            $srcUser = $exportedUsers[$u.SourceSam]
            if ($srcUser -is [array]) { $srcUser = $srcUser | Select-Object -First 1 }

            $givenName = if ($srcUser -and $srcUser.GivenName) { [string]$srcUser.GivenName } else { '' }
            $surname = if ($srcUser -and $srcUser.Surname) { [string]$srcUser.Surname } else { '' }
            $displayName = if ($srcUser -and $srcUser.DisplayName) { [string]$srcUser.DisplayName } else { '' }
            $sam = if ([string]::IsNullOrWhiteSpace($u.TargetSam)) { [string]$u.SourceSam } else { [string]$u.TargetSam }

            $check = Test-DefaultPasswordCandidate -Password $candidateSecure -SamAccountName $sam -GivenName $givenName -Surname $surname -DisplayName $displayName -MinPasswordLength $minPasswordLength -ComplexityEnabled $complexityEnabled
            if (-not $check.IsValid) {
                $problemUsers += [PSCustomObject]@{
                    Sam     = $sam
                    Reasons = ($check.Reasons -join ' ')
                }
            }
        }

        if ($problemUsers.Count -eq 0) {
            $passwordAccepted = $true
            $securePassword = $candidateSecure
            Write-Log -Message "Default password validation attempt #$passwordAttempt succeeded for all pending users." -Level INFO
            break
        }

        $sample = @($problemUsers | Select-Object -First 5)
        Write-Log -Message "Default password validation attempt #$passwordAttempt flagged $($problemUsers.Count) user(s) as likely to fail policy checks." -Level WARN
        foreach ($p in $sample) {
            Write-Log -Message "Password preflight warning for '$($p.Sam)': $($p.Reasons)" -Level WARN
        }

        Write-Host "[!] Password preflight warning: likely policy failure for $($problemUsers.Count) user(s)." -ForegroundColor Yellow
        $decision = Read-Host "Enter R to retry password, C to continue anyway, or A to abort"
        switch (($decision | ForEach-Object { $_.Trim().ToUpperInvariant() })) {
            'R' {
                Write-Log -Message "User chose to retry default password after attempt #$passwordAttempt." -Level WARN
                continue
            }
            'C' {
                $passwordAccepted = $true
                $securePassword = $candidateSecure
                Write-Log -Message "User selected Continue Anyway for default password after attempt #$passwordAttempt." -Level WARN
                break
            }
            default {
                Write-Log -Message "User aborted during default password preflight after attempt #$passwordAttempt." -Level WARN
                throw "User cancelled operation during default password preflight."
            }
        }
    }

}

$script:AccountStats = [ordered]@{
    UsersCreated = 0; UsersSkipped = 0; UsersFailed = 0
    CompsCreated = 0; CompsSkipped = 0; CompsFailed = 0
    GroupsCreated = 0; GroupsSkipped = 0; GroupsFailed = 0
    MembersAdded = 0; MembersSkipped = 0; MembersFailed = 0
}

# --- 4. Import Users ---
Invoke-Safely -ScriptBlock {
    Write-Host "`n--- Importing Users ---" -ForegroundColor Cyan
    foreach ($u in $usersToCreate) {
        $srcUser = $exportedUsers[$u.SourceSam]
        $targetSam = $u.TargetSam
        $targetOU = $u.TargetOU_DN
        $upn = "$targetSam@$TargetDomain"

        $existingUser = Get-ADUser -Filter { SamAccountName -eq $targetSam } -Server $TargetDomain -ErrorAction SilentlyContinue
        if ($existingUser) {
            $script:AccountStats.UsersSkipped++
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
            if ($PSCmdlet.ShouldProcess($targetSam, "Create User in $targetOU") -and -not $WhatIfPreference) {
                New-ADUser @userParams -ErrorAction Stop
                Write-Log -Message "Created User: $targetSam" -Level INFO
                $script:AccountStats.UsersCreated++
            }
        } catch {
            $script:AccountStats.UsersFailed++
            Write-Log -Message "Failed to create user '$targetSam': $($_.Exception.Message)" -Level ERROR
        }
    }

    Write-Log -Message "User import summary: created=$($script:AccountStats.UsersCreated); skippedExisting=$($script:AccountStats.UsersSkipped); failed=$($script:AccountStats.UsersFailed)." -Level INFO
    if ($script:AccountStats.UsersSkipped -gt 0) {
        Write-Host "[!] Users: Skipped $($script:AccountStats.UsersSkipped) existing account(s). See log for details." -ForegroundColor Yellow
    }
} -Operation "Import Users"

# --- 5. Import Computers ---
Invoke-Safely -ScriptBlock {
    Write-Host "`n--- Importing Computers ---" -ForegroundColor Cyan
    $compsToCreate = $CompMap | Where-Object { $_.Action -eq 'Create' }
    foreach ($c in $compsToCreate) {
        $targetName = $c.TargetName
        $targetOU = $c.TargetOU_DN

        $existingComputer = Get-ADComputer -Filter { Name -eq $targetName } -Server $TargetDomain -ErrorAction SilentlyContinue
        if ($existingComputer) {
            $script:AccountStats.CompsSkipped++
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
            if ($PSCmdlet.ShouldProcess($targetName, "Create Computer in $targetOU") -and -not $WhatIfPreference) {
                New-ADComputer @compParams -ErrorAction Stop
                Write-Log -Message "Created Computer: $targetName" -Level INFO
                $script:AccountStats.CompsCreated++
            }
        } catch {
            $script:AccountStats.CompsFailed++
            Write-Log -Message "Failed to create computer '$targetName': $($_.Exception.Message)" -Level ERROR
        }
    }

    Write-Log -Message "Computer import summary: created=$($script:AccountStats.CompsCreated); skippedExisting=$($script:AccountStats.CompsSkipped); failed=$($script:AccountStats.CompsFailed)." -Level INFO
    if ($script:AccountStats.CompsSkipped -gt 0) {
        Write-Host "[!] Computers: Skipped $($script:AccountStats.CompsSkipped) existing account(s). See log for details." -ForegroundColor Yellow
    }
} -Operation "Import Computers"

# --- 6. Import Groups ---
Invoke-Safely -ScriptBlock {
    Write-Host "`n--- Importing Groups ---" -ForegroundColor Cyan
    $groupsToCreate = $GroupMap | Where-Object { $_.Action -eq 'Create' }
    foreach ($g in $groupsToCreate) {
        $srcGroup = $exportedGroups[$g.SourceSam]
        $targetSam = $g.TargetSam
        $targetOU = $g.TargetOU_DN

        $existingGroup = Get-ADGroup -Filter { SamAccountName -eq $targetSam } -Server $TargetDomain -ErrorAction SilentlyContinue
        if ($existingGroup) {
            $script:AccountStats.GroupsSkipped++
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
            if ($PSCmdlet.ShouldProcess($targetSam, "Create Group in $targetOU") -and -not $WhatIfPreference) {
                New-ADGroup @grpParams -ErrorAction Stop
                Write-Log -Message "Created Group: $targetSam" -Level INFO
                $script:AccountStats.GroupsCreated++
            }
        } catch {
            $script:AccountStats.GroupsFailed++
            Write-Log -Message "Failed to create group '$targetSam': $($_.Exception.Message)" -Level ERROR
        }
    }

    Write-Log -Message "Group import summary: created=$($script:AccountStats.GroupsCreated); skippedExisting=$($script:AccountStats.GroupsSkipped); failed=$($script:AccountStats.GroupsFailed)." -Level INFO
    if ($script:AccountStats.GroupsSkipped -gt 0) {
        Write-Host "[!] Groups: Skipped $($script:AccountStats.GroupsSkipped) existing group(s). See log for details." -ForegroundColor Yellow
    }
} -Operation "Import Groups"

# --- 7. Restore Memberships ---
Invoke-Safely -ScriptBlock {
    if ($exportedMembers) {
        Write-Host "`n--- Restoring Group Memberships ---" -ForegroundColor Cyan
        
        foreach ($m in $exportedMembers) {
            # Depending on how Export-AccountData exports it, we look for Group/Member names
            $srcGroup = if ($m.GroupSam) { $m.GroupSam }
                        elseif ($m.GroupSamAccountName) { $m.GroupSamAccountName }
                        else { $m.GroupName }
            $srcMember = if ($m.MemberSam) { $m.MemberSam }
                         elseif ($m.MemberSamAccountName) { $m.MemberSamAccountName }
                         else { $m.MemberName }

            if ([string]::IsNullOrWhiteSpace($srcGroup) -or [string]::IsNullOrWhiteSpace($srcMember)) {
                $script:AccountStats.MembersSkipped++
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
                    if ($PSCmdlet.ShouldProcess("Add $tgtMember to $tgtGroup", "Group Membership") -and -not $WhatIfPreference) {
                        # Use SilentlyContinue to ignore "Already a member" errors gracefully
                        Add-ADGroupMember -Identity $tgtGroup -Members $tgtMember -Server $TargetDomain -ErrorAction SilentlyContinue
                        $script:AccountStats.MembersAdded++
                    }
                } catch {
                    $script:AccountStats.MembersFailed++
                    Write-Log -Message "Failed to add '$tgtMember' to '$tgtGroup': $($_.Exception.Message)" -Level ERROR
                }
            } else {
                $script:AccountStats.MembersSkipped++
                Write-Log -Message "Skipping membership '$srcMember' -> '$srcGroup' because one or both mappings were not found (TargetMember='$tgtMember', TargetGroup='$tgtGroup')." -Level WARN
            }
        }
        Write-Log -Message "Processed $($script:AccountStats.MembersAdded) membership links; skipped $($script:AccountStats.MembersSkipped); failed $($script:AccountStats.MembersFailed)." -Level INFO
        Write-Host "Processed $($script:AccountStats.MembersAdded) membership links; skipped $($script:AccountStats.MembersSkipped); failed $($script:AccountStats.MembersFailed)." -ForegroundColor Green
    } else {
        Write-Log -Message "No GroupMembers export found. Skipping membership restore." -Level WARN
    }
} -Operation "Import Memberships"

$warningCount = $script:AccountStats.UsersFailed + $script:AccountStats.CompsFailed + $script:AccountStats.GroupsFailed + $script:AccountStats.MembersFailed
$summary = "Account Import summary: Users (Created=$($script:AccountStats.UsersCreated), Skipped=$($script:AccountStats.UsersSkipped), Failed=$($script:AccountStats.UsersFailed)), " +
           "Computers (Created=$($script:AccountStats.CompsCreated), Skipped=$($script:AccountStats.CompsSkipped), Failed=$($script:AccountStats.CompsFailed)), " +
           "Groups (Created=$($script:AccountStats.GroupsCreated), Skipped=$($script:AccountStats.GroupsSkipped), Failed=$($script:AccountStats.GroupsFailed)), " +
           "Memberships (Added=$($script:AccountStats.MembersAdded), Skipped=$($script:AccountStats.MembersSkipped), Failed=$($script:AccountStats.MembersFailed))"

if ($warningCount -gt 0) {
    Write-Host "`n[!] WARNING: Account Import encountered $warningCount failure(s). See logs for details." -ForegroundColor Yellow
    Write-Log -Message "Import Accounts completed with warnings. $summary" -Level WARN
} else {
    Write-Host "`n=== Account Import Complete ===" -ForegroundColor Green
    Write-Log -Message "Import Accounts completed. $summary" -Level INFO
}