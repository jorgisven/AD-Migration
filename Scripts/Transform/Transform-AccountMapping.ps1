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

# Ensure transform directory exists
if (-not (Test-Path $MapPath)) { New-Item -ItemType Directory -Path $MapPath -Force | Out-Null }

Write-Log -Message "Starting Account Mapping analysis against $TargetDomain..." -Level INFO

$IdentityMap = @()

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

        foreach ($item in $items) {
            $sam = if ($item.SamAccountName) { $item.SamAccountName } elseif ($item.Name) { $item.Name } else { "UNKNOWN" }
            $sid = $item.SID
            $targetName = $sam
            
            # Computers usually have a trailing $ in SamAccountName, strip it for Name matching
            if ($ObjectType -eq 'Computer' -and $targetName -match '\$$') {
                $targetName = $targetName -replace '\$$',''
            }

            # Determine Target OU
            $sourceDN = if ($item.DistinguishedName) { $item.DistinguishedName } else { "" }
            $parentDN = $sourceDN -replace '^[^,]+,', ''
            $targetOU = if ($OUMap.ContainsKey($parentDN)) { $OUMap[$parentDN] } else { "" }

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

            if ($exists) {
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
                $IdentityMap += [PSCustomObject]@{
                    SourceSam = $sam
                    SourceSID = $sid
                    TargetSam = $targetName
                    Type      = $ObjectType
                }
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
    $IdentityMap | Export-Csv -Path $mapFile -NoTypeInformation -Encoding UTF8

    Write-Log -Message "Generated identity map: $mapFile" -Level INFO

    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Account Mapping Drafts Generated in:\n  $MapPath"
    Write-Host "\nACTION REQUIRED:" -ForegroundColor Yellow
    Write-Host "Review the User, Computer, and Group Account Map CSVs." -ForegroundColor Yellow
    Write-Host "Ensure Target OUs are correct, and verify any 'Merge' actions." -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan

} -Operation "Map Accounts"