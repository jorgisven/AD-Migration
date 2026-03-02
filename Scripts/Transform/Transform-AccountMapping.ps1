<#
.SYNOPSIS
    Map user accounts and detect conflicts in the target domain.

.DESCRIPTION
    Compares exported source users against the target domain to identify naming conflicts.
    Generates 'User_Migration_Plan.csv' (for review) and 'UserMap.csv' (for GPO migration tables).
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO

$config = Get-ADMigrationConfig
$SourceSecurityPath = Join-Path $config.ExportRoot 'Security'
$TransformPath = $config.TransformRoot

# Ensure transform directory exists
if (-not (Test-Path $TransformPath)) { New-Item -ItemType Directory -Path $TransformPath -Force | Out-Null }

Write-Log -Message "Starting Account Mapping analysis against $TargetDomain..." -Level INFO

# 1. Find latest User export
$userFiles = Get-ChildItem -Path $SourceSecurityPath -Filter "Users_*.csv" | Sort-Object LastWriteTime -Descending
if (-not $userFiles) {
    Write-Log -Message "No User export found. Run Export-AccountData.ps1 first." -Level ERROR
    throw "Missing User Data"
}

$SourceUsers = Import-Csv $userFiles[0].FullName
Write-Log -Message "Loaded $($SourceUsers.Count) users from $($userFiles[0].Name)" -Level INFO

$MigrationPlan = @()
$UserMap = @()

Invoke-Safely -ScriptBlock {
    foreach ($user in $SourceUsers) {
        $sam = $user.SamAccountName
        $status = "New"
        $targetSam = $sam
        $action = "Create"
        $notes = ""

        # Check Target Domain
        try {
            $targetUser = Get-ADUser -Identity $sam -Server $TargetDomain -ErrorAction Stop
            $status = "Conflict"
            $action = "Skip" # Default safety: Do not overwrite existing accounts
            $notes = "Account exists in target."
        } catch {
            # Not found, which is good for migration
            $status = "New"
        }

        $MigrationPlan += [PSCustomObject]@{
            SourceSam   = $sam
            TargetSam   = $targetSam
            Status      = $status
            Action      = $action
            SourceUPN   = $user.UserPrincipalName
            Description = $notes
        }

        # Add to Map (used by GPO migration table)
        $UserMap += [PSCustomObject]@{
            SourceSam = $sam
            TargetSam = $targetSam
        }
    }

    # Output Plan
    $planFile = Join-Path $TransformPath "User_Migration_Plan.csv"
    $MigrationPlan | Export-Csv -Path $planFile -NoTypeInformation -Encoding UTF8
    
    # Output Map (Simple format for other scripts)
    $mapFile = Join-Path $TransformPath "UserMap.csv"
    $UserMap | Export-Csv -Path $mapFile -NoTypeInformation -Encoding UTF8

    Write-Log -Message "Generated migration plan: $planFile" -Level INFO
    Write-Log -Message "Generated user map: $mapFile" -Level INFO

    $conflicts = $MigrationPlan | Where-Object { $_.Status -eq 'Conflict' }
    
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Account Mapping Complete"
    Write-Host "Plan: $planFile"
    if ($conflicts) {
        Write-Host "WARNING: Found $($conflicts.Count) naming conflicts in target domain." -ForegroundColor Yellow
        Write-Host "Check the CSV and update 'TargetSam' or 'Action' columns." -ForegroundColor Yellow
    } else {
        Write-Host "No naming conflicts detected." -ForegroundColor Green
    }
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan

} -Operation "Map Accounts"