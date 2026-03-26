<#
.SYNOPSIS
    Rebuild GPO link structure.

.DESCRIPTION
    Uses XML report data to recreate link locations and link order in target domain.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDefaultDomainPolicies,

    [Parameter(Mandatory = $false)]
    [string]$DefaultPolicyNameSuffix = ' (Migrated Copy)',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Prompt', 'Skip', 'Rename')]
    [string]$DefaultDomainPolicyMode = 'Prompt',

    [Parameter(Mandatory = $false)]
    [switch]$EnableLinks
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

$config = Get-ADMigrationConfig
$ReportPath = Join-Path $config.ExportRoot 'GPO_Reports'
$MapPath = Join-Path $config.TransformRoot 'Mapping'

function ConvertTo-NormalizedDN {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $normalized = $Value.Trim()
    # Normalize escaped separators and whitespace so map keys can match report variants.
    $normalized = $normalized -replace '\\,', ','
    $normalized = $normalized -replace '\\=', '='
    $normalized = $normalized -replace '\\#', '#'
    $normalized = $normalized -replace '\\;', ';'
    $normalized = $normalized -replace '\\"', '"'
    $normalized = $normalized -replace '\\<', '<'
    $normalized = $normalized -replace '\\>', '>'
    $normalized = $normalized -replace '\\+', '+'
    $normalized = $normalized -replace '\\ ', ' '
    $normalized = $normalized -replace '\s*,\s*', ','
    return $normalized.ToLowerInvariant()
}

function Convert-DNToCanonical {
    param([string]$DistinguishedName)

    $dn = ConvertTo-NormalizedDN -Value $DistinguishedName
    if ([string]::IsNullOrWhiteSpace($dn)) { return $null }

    $parts = $dn -split ','
    $ouParts = @()
    $dcParts = @()

    foreach ($part in $parts) {
        if ($part -like 'ou=*') {
            $ouParts += ($part.Substring(3))
            continue
        }

        if ($part -like 'dc=*') {
            $dcParts += ($part.Substring(3))
        }
    }

    if (-not $dcParts -or $dcParts.Count -eq 0) { return $null }

    $domain = ($dcParts -join '.').ToLowerInvariant()
    if ($ouParts.Count -gt 0) {
        [array]::Reverse($ouParts)
        return ($domain + '/' + ($ouParts -join '/')).ToLowerInvariant()
    }

    return $domain
}

# 1. Load OU Map (Transform Data)
$mapFiles = Get-ChildItem -Path $MapPath -Filter "OU_Map_*.csv" | Sort-Object LastWriteTime -Descending
if (-not $mapFiles) { throw "Missing OU Map (Run Transform-OUMap.ps1)" }

$OUMapByDN = @{}
$OUMapByCanonical = @{}

Import-Csv $mapFiles[0].FullName | ForEach-Object {
    $sourceDN = [string]$_.SourceDN
    $targetDN = [string]$_.TargetDN
    $action = [string]$_.Action

    if ($action -eq 'Skip' -or [string]::IsNullOrWhiteSpace($sourceDN) -or [string]::IsNullOrWhiteSpace($targetDN)) {
        return
    }

    $dnKey = ConvertTo-NormalizedDN -Value $sourceDN
    if (-not [string]::IsNullOrWhiteSpace($dnKey)) {
        $OUMapByDN[$dnKey] = $targetDN
    }

    $canonicalKey = Convert-DNToCanonical -DistinguishedName $sourceDN
    if (-not [string]::IsNullOrWhiteSpace($canonicalKey)) {
        $OUMapByCanonical[$canonicalKey] = $targetDN
    }
}

Write-Log -Message "Loaded OU Map with $($OUMapByDN.Count) DN keys and $($OUMapByCanonical.Count) canonical keys" -Level INFO

# 1.5 Load Identity Map (for Security Filtering)
$IdentityMap = @{}
$idMapFile = Join-Path $MapPath "Identity_Map_Final.csv"
if (Test-Path $idMapFile) {
    Import-Csv $idMapFile | ForEach-Object {
        if (-not [string]::IsNullOrWhiteSpace($_.SourceSam)) {
            $IdentityMap[$_.SourceSam] = $_.TargetSam
        }
    }
    Write-Log -Message "Loaded Identity Map with $($IdentityMap.Count) entries for Security Filtering." -Level INFO
}

# 2. Load GPO XML Reports (Export Data)
$xmlFiles = Get-ChildItem -Path $ReportPath -Filter "*.xml"
if (-not $xmlFiles) { throw "Missing GPO Reports (Run Export-GPOReports.ps1)" }

if (-not $TargetDomain) { $TargetDomain = (Get-ADDomain).DNSRoot }

$defaultPolicyNames = @(
    'Default Domain Policy',
    'Default Domain Controllers Policy'
)

$effectiveDefaultPolicyMode = $DefaultDomainPolicyMode
if ($effectiveDefaultPolicyMode -eq 'Prompt' -and $IncludeDefaultDomainPolicies) {
    # Backward compatibility: old switch implied importing defaults.
    $effectiveDefaultPolicyMode = 'Rename'
    Write-Log -Message "Legacy switch IncludeDefaultDomainPolicies was set. Effective default policy mode changed from Prompt to Rename." -Level WARN
}

if ($effectiveDefaultPolicyMode -eq 'Prompt') {
    # Prefer safe non-blocking behavior in direct script usage.
    $effectiveDefaultPolicyMode = 'Skip'
    Write-Log -Message "DefaultDomainPolicyMode='Prompt' was provided to link rebuild. Using Skip unless orchestrator passes explicit mode." -Level WARN
}

if ($effectiveDefaultPolicyMode -eq 'Rename' -and [string]::IsNullOrWhiteSpace($DefaultPolicyNameSuffix)) {
    $DefaultPolicyNameSuffix = ' (Migrated Copy)'
}

Write-Log -Message "Default domain policy handling mode for links: $effectiveDefaultPolicyMode" -Level INFO
Write-Log -Message "Default domain policy options for links: Mode='$effectiveDefaultPolicyMode', Suffix='$DefaultPolicyNameSuffix'" -Level INFO

if (-not $EnableLinks) {
    Write-Log -Message "Safety Feature Active: All new GPO links will be created in a DISABLED state by default." -Level INFO
} else {
    Write-Log -Message "GPO links will be restored to their original enabled/disabled state." -Level INFO
}

Write-Log -Message "Starting GPO Link Rebuild..." -Level INFO

$script:LinkStats = [ordered]@{
    GposProcessed               = 0
    LinksEvaluated              = 0
    Linked                      = 0
    AlreadyLinked               = 0
    SkippedDefaultPolicies      = 0
    SkippedMissingOUMap         = 0
    LinkFailures                = 0
    WmiFilterFailures           = 0
    SecurityFilterFailures      = 0
    AuthenticatedUsersRemovalsFailed = 0
}

Invoke-Safely -ScriptBlock {
    foreach ($xmlFile in $xmlFiles) {
        $script:LinkStats.GposProcessed++
        [xml]$xml = Get-Content $xmlFile.FullName
        $sourceGpoName = [string]$xml.GPO.Name
        $gpoName = $sourceGpoName

        if ($defaultPolicyNames -contains $sourceGpoName) {
            if ($effectiveDefaultPolicyMode -eq 'Skip') {
                $script:LinkStats.SkippedDefaultPolicies++
                Write-Log -Message "Skipping link restore for '$sourceGpoName' because default domain policies are excluded by default." -Level INFO
                continue
            }

            if ($effectiveDefaultPolicyMode -eq 'Rename') {
                $gpoName = "$sourceGpoName$DefaultPolicyNameSuffix"
            }
        }
        
        # Find LinksTo (Where this GPO is linked)
        $links = $xml.GPO.LinksTo
        
        if ($links) {
            foreach ($link in $links) {
                $script:LinkStats.LinksEvaluated++
                $sourceSOM = $link.SOMPath # e.g. "ou=sales,dc=source,dc=local"
                $targetSOM = $null
                $sourceSomText = [string]$sourceSOM
                $sourceSomKey = ConvertTo-NormalizedDN -Value $sourceSomText
                $sourceSomCanonical = if ($sourceSomText) { $sourceSomText.Trim().ToLowerInvariant() } else { '' }
                
                # Determine Target Location
                if (-not [string]::IsNullOrWhiteSpace($sourceSomKey) -and $OUMapByDN.ContainsKey($sourceSomKey)) {
                    $targetSOM = $OUMapByDN[$sourceSomKey]
                } elseif (-not [string]::IsNullOrWhiteSpace($sourceSomCanonical) -and $OUMapByCanonical.ContainsKey($sourceSomCanonical)) {
                    $targetSOM = $OUMapByCanonical[$sourceSomCanonical]
                } elseif ($sourceSomText -match '^\s*(DC=[^,]+\s*(,\s*DC=[^,]+\s*)*)$' -or ($sourceSomText -notmatch '=' -and $sourceSomText -notmatch '/' -and $sourceSomText -match '\.')) {
                    # If linked to Domain Root, map to Target Domain Root
                    $targetSOM = $TargetDomain # Simplified, assumes domain root link
                }
                
                if ($targetSOM) {
                    try {
                        # Check if link exists
                        # Note: Get-GPO doesn't easily show links, usually we just try to add it
                        # New-GPLink throws if GPO doesn't exist, but we assume Import-GPOs ran
                        
                        if ($PSCmdlet.ShouldProcess($targetSOM, "Link GPO '$gpoName'") -and -not $WhatIfPreference) {
                            $linkEnabledVal = if ($EnableLinks -and $link.Enabled -eq 'true') { 'Yes' } else { 'No' }
                            $linkEnforcedVal = if ($link.NoOverride -eq 'true') { 'Yes' } else { 'No' }
                            New-GPLink -Name $gpoName -Target $targetSOM -LinkEnabled $linkEnabledVal -Enforced $linkEnforcedVal -Server $TargetDomain -ErrorAction Stop | Out-Null
                            $script:LinkStats.Linked++
                            $stateMsg = if ($linkEnabledVal -eq 'Yes') { "Enabled" } else { "Disabled" }
                            Write-Log -Message "Linked '$gpoName' to '$targetSOM' (State: $stateMsg)" -Level INFO
                        }
                    } catch {
                        if ($_.Exception.Message -match "already linked") {
                            $script:LinkStats.AlreadyLinked++
                            Write-Log -Message "Link for '$gpoName' at '$targetSOM' already exists." -Level INFO
                        } else {
                            $script:LinkStats.LinkFailures++
                            Write-Log -Message "Failed to link '$gpoName' to '$targetSOM': $_" -Level WARN
                        }
                    }
                } else {
                    $script:LinkStats.SkippedMissingOUMap++
                    Write-Log -Message "Skipping link for '$gpoName': Source path '$sourceSOM' not found in OU Map (DN/canonical lookup)." -Level WARN
                }
            }
        }
        
        # Restore WMI Filter Link if present
        $wmiFilter = $xml.GPO.FilterData.Name
        if ($wmiFilter) {
            try {
                if ($PSCmdlet.ShouldProcess($gpoName, "Link WMI Filter '$wmiFilter'") -and -not $WhatIfPreference) {
                    $gpo = Get-GPO -Name $gpoName -Server $TargetDomain
                    $gpo.WmiFilter = $wmiFilter # This property sets the link
                    Write-Log -Message "Linked WMI Filter '$wmiFilter' to '$gpoName'" -Level INFO
                }
            } catch {
                $script:LinkStats.WmiFilterFailures++
                Write-Log -Message "Failed to link WMI Filter '$wmiFilter' to '$gpoName'" -Level WARN
            }
        }

        # Restore Security Filtering (Permissions)
        $securityDescriptor = $xml.GPO.SecurityDescriptor
        if ($securityDescriptor) {
            $customFiltersApplied = $false
            foreach ($permission in $securityDescriptor.Permissions.Permission) {
                # We only care about re-applying GpoApply filters
                if ($permission.Type -eq 'GpoApply') {
                    $trusteeName = $permission.Trustee.Name
                    $trusteeType = $permission.Trustee.Type # User, Group, Computer
                    
                    # Skip well-known SIDs that are handled by default or are not filters
                    if ($trusteeName -in @('NT AUTHORITY\ENTERPRISE DOMAIN CONTROLLERS', 'NT AUTHORITY\Authenticated Users', 'NT AUTHORITY\SYSTEM')) {
                        continue
                    }

                    # Resolve Target Identity
                    $sam = $trusteeName
                    if ($trusteeName -match "\\") {
                        $sam = ($trusteeName -split "\\")[1]
                    }

                    $targetTrustee = $sam
                    if ($IdentityMap.ContainsKey($sam)) {
                        $targetTrustee = $IdentityMap[$sam]
                    }

                    try {
                        Write-Log -Message "Applying GpoApply security filter for '$trusteeName' on GPO '$gpoName'" -Level INFO
                        if ($PSCmdlet.ShouldProcess($gpoName, "Set GPPermission '$trusteeName'") -and -not $WhatIfPreference) {
                            Set-GPPermission -Name $gpoName -PermissionLevel GpoApply -TargetName $trusteeName -TargetType $trusteeType -Server $TargetDomain -ErrorAction Stop
                        }
                        Write-Log -Message "Applying GpoApply security filter for '$targetTrustee' (Source: $trusteeName) on GPO '$gpoName'" -Level INFO
                        if ($PSCmdlet.ShouldProcess($gpoName, "Set GPPermission '$targetTrustee'") -and -not $WhatIfPreference) {
                            Set-GPPermission -Name $gpoName -PermissionLevel GpoApply -TargetName $targetTrustee -TargetType $trusteeType -Server $TargetDomain -ErrorAction Stop
                            $customFiltersApplied = $true
                        }
                    }
                    catch {
                        $script:LinkStats.SecurityFilterFailures++
                        Write-Log -Message "Failed to set GpoApply for '$trusteeName' on GPO '$gpoName'. Principal may not exist in target domain. Details: $_" -Level WARN
                        Write-Log -Message "Failed to set GpoApply for '$targetTrustee' on GPO '$gpoName'. Principal may not exist in target domain. Details: $_" -Level WARN
                    }
                }
            }

            # If custom filters were added, remove the default 'Authenticated Users' permission
            # to ensure the policy is correctly restricted.
            if ($customFiltersApplied) {
                try {
                    Write-Log -Message "Custom filters applied. Removing default 'Authenticated Users' GpoApply permission from '$gpoName'." -Level INFO
                    if ($PSCmdlet.ShouldProcess($gpoName, "Remove 'Authenticated Users'") -and -not $WhatIfPreference) {
                        Set-GPPermission -Name $gpoName -PermissionLevel None -TargetName 'Authenticated Users' -TargetType Group -Server $TargetDomain -ErrorAction Stop
                    }
                } catch {
                    $script:LinkStats.AuthenticatedUsersRemovalsFailed++
                    # This might fail if it was already removed or never existed, which is fine.
                    Write-Log -Message "Could not remove GpoApply from 'Authenticated Users' on GPO '$gpoName'. It may have already been removed. Details: $_" -Level WARN
                }
            }
        }
    }
} -Operation "Rebuild GPO Links"

$warningCount =
    $script:LinkStats.SkippedMissingOUMap +
    $script:LinkStats.LinkFailures +
    $script:LinkStats.WmiFilterFailures +
    $script:LinkStats.SecurityFilterFailures +
    $script:LinkStats.AuthenticatedUsersRemovalsFailed

$summary = "Rebuild GPO Links summary: GPOs=$($script:LinkStats.GposProcessed), LinksEvaluated=$($script:LinkStats.LinksEvaluated), Linked=$($script:LinkStats.Linked), AlreadyLinked=$($script:LinkStats.AlreadyLinked), SkippedDefaultPolicies=$($script:LinkStats.SkippedDefaultPolicies), SkippedMissingOUMap=$($script:LinkStats.SkippedMissingOUMap), LinkFailures=$($script:LinkStats.LinkFailures), WMIFilterFailures=$($script:LinkStats.WmiFilterFailures), SecurityFilterFailures=$($script:LinkStats.SecurityFilterFailures), AuthUsersRemovalFailures=$($script:LinkStats.AuthenticatedUsersRemovalsFailed)"

if ($warningCount -gt 0) {
    Write-Log -Message "Rebuild GPO Links completed with warnings. $summary" -Level WARN
} else {
    Write-Log -Message "Rebuild GPO Links completed. $summary" -Level INFO
}