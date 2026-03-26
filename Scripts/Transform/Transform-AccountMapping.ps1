<#
.SYNOPSIS
    Map user accounts and detect conflicts in the target domain.

.DESCRIPTION
    Compares exported source users against the target domain to identify naming conflicts.
    Generates individual Account Mapping CSVs and auto-resolves OU destinations using the OU Map.
    Also generates 'Identity_Map_Final.csv' for the GPO migration table process.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain
)

# Add GUI support for conflict prompts
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

if ([string]::IsNullOrWhiteSpace($TargetDomain)) {
    $TargetDomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the FQDN of the TARGET domain (e.g., target.local):", "Account Mapping", "")
    if ([string]::IsNullOrWhiteSpace($TargetDomain)) {
        Write-Log -Message "Target domain not provided. Aborting." -Level ERROR
        throw "Target domain is required for account mapping."
    }
}

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO

$config = Get-ADMigrationConfig
$SourceSecurityPath = Join-Path $config.ExportRoot 'Security'
$MapPath = Join-Path $config.TransformRoot 'Mapping'
$ExplicitMembershipKeepGroups = @{}

# Ensure transform directory exists
if (-not (Test-Path $MapPath)) { New-Item -ItemType Directory -Path $MapPath -Force | Out-Null }

Write-Log -Message "Starting Account Mapping analysis against $TargetDomain..." -Level INFO

$script:IdentityMap = @()
$script:IdentityMapStats = [ordered]@{
    TotalRead                = 0
    MissingSid               = 0
    SkippedDomainControllers = 0
    Added                    = 0
}

# 1. Load OU Map to pre-fill TargetOU_DN
$OUMap = @{}
$ouMapFile = Join-Path $MapPath "OU_Map_Draft.csv"
if (Test-Path $ouMapFile) {
    Import-Csv $ouMapFile | ForEach-Object {
        if ($_.Action -ne 'Skip' -and -not [string]::IsNullOrWhiteSpace($_.SourceDN)) {
            $OUMap[$_.SourceDN] = $_.TargetDN
        }
    }
    Write-Log -Message "Loaded $($OUMap.Count) mapped OUs for account placement." -Level INFO
}

# Load explicitly preserved privileged memberships from Export-AccountData selection.
$memberFile = Get-ChildItem -Path $SourceSecurityPath -Filter "GroupMembers_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($memberFile) {
    try {
        Import-Csv $memberFile.FullName | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_.GroupSam)) {
                $ExplicitMembershipKeepGroups[[string]$_.GroupSam] = $true
            }
        }
        Write-Log -Message "Loaded explicit privileged membership keeps for $($ExplicitMembershipKeepGroups.Count) group(s) from $($memberFile.Name)." -Level INFO
    } catch {
        Write-Log -Message "Failed loading explicit privileged membership keeps from $($memberFile.FullName): $_" -Level WARN
    }
}

Invoke-Safely -ScriptBlock {
    function Invoke-IdentityMapping {
        param(
            [string]$FileType,
            [string]$ObjectType,
            [string]$CheckCmdlet
        )

        $files = Get-ChildItem -Path $SourceSecurityPath -Filter "${FileType}_*.csv" | Sort-Object LastWriteTime -Descending
        if (-not $files) {
            Write-Log -Message "No $FileType export found. Skipping." -Level WARN
            return
        }

        $items = Import-Csv $files[0].FullName
        Write-Log -Message "Processing $($items.Count) $ObjectType(s) from $($files[0].Name)" -Level INFO
        
        $MappingData = @()

        # Define regex patterns for well-known SIDs of built-in accounts to be labeled
        $builtinSidPatterns = @(
            'S-1-5-.*-500$', # Administrator
            'S-1-5-.*-501$', # Guest
            'S-1-5-.*-502$', # krbtgt
            'S-1-5-.*-503$', # DefaultAccount
            '^S-1-5-32-\d+$' # Builtin local/domain alias groups (e.g., System Managed Accounts Group)
        )

        # Protected default domain groups should be treated as system groups and skipped by default.
        $protectedDefaultGroupSidPatterns = @(
            'S-1-5-.*-512$', # Domain Admins
            'S-1-5-.*-516$', # Domain Controllers
            'S-1-5-.*-518$', # Schema Admins
            'S-1-5-.*-519$', # Enterprise Admins
            'S-1-5-.*-520$', # Group Policy Creator Owners
            'S-1-5-.*-521$', # Read-only Domain Controllers
            'S-1-5-.*-522$', # Cloneable Domain Controllers
            'S-1-5-.*-525$', # Protected Users
            'S-1-5-.*-526$', # Key Admins
            'S-1-5-.*-527$'  # Enterprise Key Admins
        )
        $protectedDefaultGroupNames = @(
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

        foreach ($item in $items) {
            $script:IdentityMapStats.TotalRead++
            # Defensive: Ensure $MappingData is always an array
            if ($null -eq $MappingData -or $MappingData.GetType().Name -ne 'Object[]') {
                $MappingData = @()
            }
            $sam = if ($item.SamAccountName) { $item.SamAccountName } elseif ($item.Name) { $item.Name } else { "UNKNOWN" }
            $sid = $item.SID
            $targetName = $sam
            $sourceDN = if ($item.DistinguishedName) { $item.DistinguishedName } else { "" }
            
            # --- BUILT-IN ACCOUNT DETECTION ---
            $isBuiltin = $false
            foreach ($pattern in $builtinSidPatterns) {
                if ($sid -match $pattern) {
                    $isBuiltin = $true
                    break
                }
            }
            if (-not $isBuiltin -and $sourceDN -match "(?:^|,)CN=Builtin,") {
                $isBuiltin = $true
            }
            if (-not $isBuiltin -and $sam -eq 'DefaultAccount') {
                $isBuiltin = $true
            }

            $isProtectedDefaultGroup = $false
            if ($ObjectType -eq 'Group') {
                foreach ($pattern in $protectedDefaultGroupSidPatterns) {
                    if ($sid -match $pattern) {
                        $isProtectedDefaultGroup = $true
                        break
                    }
                }
                if (-not $isProtectedDefaultGroup -and $protectedDefaultGroupNames -contains $sam) {
                    $isProtectedDefaultGroup = $true
                }
            }
            $hasExplicitMembershipKeep = ($ObjectType -eq 'Group' -and $ExplicitMembershipKeepGroups.ContainsKey($sam))

            # Computers usually have a trailing $ in SamAccountName, strip it for Name matching
            if ($ObjectType -eq 'Computer' -and $targetName -match '\$$') {
                $targetName = $targetName -replace '\$$',''
            }

            # Determine Target OU
            $parentDN = $sourceDN -replace '^[^,]+,', ''

            # Skip objects located in the Domain Controllers OU (including child OUs)
            if ($parentDN -match "(?:^|,)OU=Domain Controllers,") {
                $script:IdentityMapStats.SkippedDomainControllers++
                continue
            }
            # Only set targetOU from map if it wasn't already set by built-in logic
            $targetOU = if ($isBuiltin) { 'BUILTIN (Do not move)' }
                        elseif ($isProtectedDefaultGroup) { 'SYSTEM DEFAULT GROUP (Do not move)' }
                        elseif ($OUMap.ContainsKey($parentDN)) { $OUMap[$parentDN] } 
                        else { "" }

            # Check Target Domain for conflict
            $exists = $false
            try {
                if ($CheckCmdlet -eq 'Get-ADUser') { $null = Get-ADUser -Identity $targetName -Server $TargetDomain -ErrorAction Stop }
                elseif ($CheckCmdlet -eq 'Get-ADGroup') { $null = Get-ADGroup -Identity $targetName -Server $TargetDomain -ErrorAction Stop }
                elseif ($CheckCmdlet -eq 'Get-ADComputer') { $null = Get-ADComputer -Identity $targetName -Server $TargetDomain -ErrorAction Stop }
                $exists = $true
            } catch {
                $exists = $false
            }

            if ($isBuiltin) {
                if ($hasExplicitMembershipKeep -and $exists) {
                    $action = "Merge"
                    $notes = "Built-in principal with explicitly preserved membership from export. Action set to Merge."
                    Write-Log -Message "Override applied for built-in principal '$sam': explicit privileged membership keep detected; Action forced to Merge." -Level WARN
                } elseif ($hasExplicitMembershipKeep -and -not $exists) {
                    $action = "Skip"
                    $notes = "Built-in principal has explicitly preserved membership, but target group was not found. Manual review required."
                    Write-Log -Message "Explicit keep could not be honored for built-in principal '$sam': target group not found in '$TargetDomain'; Action remains Skip." -Level WARN
                } else {
                    $action = "Skip" # Always skip built-in accounts by default
                    $notes = "Built-in principal. Action set to Skip."
                }
            } elseif ($isProtectedDefaultGroup) {
                if ($hasExplicitMembershipKeep -and $exists) {
                    $action = "Merge"
                    $notes = "Protected default domain group with explicitly preserved membership from export. Action set to Merge."
                    Write-Log -Message "Override applied for protected default group '$sam': explicit privileged membership keep detected; Action forced to Merge." -Level WARN
                } elseif ($hasExplicitMembershipKeep -and -not $exists) {
                    $action = "Skip"
                    $notes = "Protected default domain group has explicitly preserved membership, but target group was not found. Manual review required."
                    Write-Log -Message "Explicit keep could not be honored for protected default group '$sam': target group not found in '$TargetDomain'; Action remains Skip." -Level WARN
                } else {
                    $action = "Skip"
                    $notes = "Protected default domain group. Action set to Skip."
                }
            } elseif ($exists) {
                $action = "Merge"
                $notes = "Exists in target domain. Defaulted to Merge."
            } else {
                $action = "Create"
                $notes = "Net-new object."
            }

            # Build specific output object
            if ($ObjectType -eq 'Computer') {
                 $MappingData += [PSCustomObject]@{
                    Action      = $action
                    SourceName  = $targetName
                    TargetName  = $targetName
                    TargetOU_DN = $targetOU
                    SourceDN    = $sourceDN
                    Description = $item.Description
                    Notes       = $notes
                }
            } else {
                 $MappingData += [PSCustomObject]@{
                    Action      = $action
                    SourceSam   = $sam
                    TargetSam   = $targetName
                    TargetOU_DN = $targetOU
                    SourceDN    = $sourceDN
                    Description = $item.Description
                    Notes       = $notes
                }
            }

            # We map SourceSID -> TargetSam because GPO Migration Tables use SIDs
            if ($sid) {
                 # Defensive: Ensure $IdentityMap is always an array
                 if ($null -eq $script:IdentityMap -or $script:IdentityMap.GetType().Name -ne 'Object[]') {
                     $script:IdentityMap = @()
                 }
                $script:IdentityMap += [PSCustomObject]@{
                    SourceSam = $sam
                    SourceSID = $sid
                    TargetSam = $targetName
                    Type      = $ObjectType
                }
                $script:IdentityMapStats.Added++
            } else {
                $script:IdentityMapStats.MissingSid++
            }
        }
        
        $outputFile = Join-Path $MapPath "${ObjectType}_Account_Map.csv"
        $MappingData | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Generated $outputFile" -Level INFO
        
        $mergeCount = ($MappingData | Where-Object { $_.Action -eq 'Merge' }).Count
        if ($mergeCount -gt 0) {
            Write-Host "  -> $($ObjectType): Found $mergeCount existing accounts (Set to Merge)." -ForegroundColor Yellow
        } else {
            Write-Host "  -> $($ObjectType): All accounts are net-new." -ForegroundColor Green
        }
    }

    Write-Host "`n--- Scanning Target Domain for Account Collisions ---" -ForegroundColor Cyan
    # Process Users, Groups, and Computers
    Invoke-IdentityMapping -FileType "Users" -ObjectType "User" -CheckCmdlet "Get-ADUser"
    Invoke-IdentityMapping -FileType "Groups" -ObjectType "Group" -CheckCmdlet "Get-ADGroup"
    Invoke-IdentityMapping -FileType "Computers" -ObjectType "Computer" -CheckCmdlet "Get-ADComputer"

    # Output Map (Simple format for other scripts)
    $mapFile = Join-Path $MapPath "Identity_Map_Final.csv"
    $script:IdentityMap | Export-Csv -Path $mapFile -NoTypeInformation -Encoding UTF8

    Write-Log -Message "Generated identity map: $mapFile" -Level INFO
    Write-Log -Message "Identity map diagnostics: Read=$($script:IdentityMapStats.TotalRead), Added=$($script:IdentityMapStats.Added), MissingSID=$($script:IdentityMapStats.MissingSid), SkippedDomainControllersOU=$($script:IdentityMapStats.SkippedDomainControllers)" -Level INFO
    if ($script:IdentityMapStats.Added -eq 0) {
        Write-Log -Message "Identity map contains 0 entries. Check export CSV SID values and account placement filters." -Level WARN
    }

    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Account Mapping Drafts Generated in:\n  $MapPath"
    Write-Host "\nACTION REQUIRED:" -ForegroundColor Yellow
    Write-Host "Review the User, Computer, and Group Account Map CSVs." -ForegroundColor Yellow
    Write-Host "Ensure Target OUs are correct, and verify any 'Merge' actions." -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan

} -Operation "Map Accounts"