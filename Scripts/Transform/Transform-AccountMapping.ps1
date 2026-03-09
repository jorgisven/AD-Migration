<#
.SYNOPSIS
    Map user accounts and detect conflicts in the target domain.

.DESCRIPTION
    Compares exported source users against the target domain to identify naming conflicts.
    Generates 'Identity_Migration_Plan.csv' (for review) and 'Identity_Map_Final.csv' (for GPO migration tables).
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$TargetDomain
)

# Add GUI support for conflict prompts
Add-Type -AssemblyName System.Windows.Forms

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

$MigrationPlan = @()
$IdentityMap = @()
$script:ConflictCount = 0
$script:IgnoreConflictWarning = $false

Invoke-Safely -ScriptBlock {
    # Helper function to process different object types
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

        foreach ($item in $items) {
            $sam = $item.SamAccountName
            $sid = $item.SID
            $targetSam = $sam
            $status = "New"
            $action = "Create"
            $notes = ""

            # Check Target Domain for conflict
            $exists = $false
            try {
                if ($CheckCmdlet -eq 'Get-ADUser') {
                    $null = Get-ADUser -Identity $sam -Server $TargetDomain -ErrorAction Stop
                } elseif ($CheckCmdlet -eq 'Get-ADGroup') {
                    $null = Get-ADGroup -Identity $sam -Server $TargetDomain -ErrorAction Stop
                } elseif ($CheckCmdlet -eq 'Get-ADComputer') {
                    $null = Get-ADComputer -Identity $sam -Server $TargetDomain -ErrorAction Stop
                }
                $exists = $true
            } catch {
                $exists = $false
            }

            if ($exists) {
                $script:ConflictCount++
                
                # Handle High Conflict Count
                if ($script:ConflictCount -eq 10 -and -not $script:IgnoreConflictWarning) {
                    $msg = "More than 10 naming conflicts have been detected so far.`n`nDo you want to continue processing and auto-renaming/skipping?`n`nYes = Continue (Don't ask again)`nNo = Abort Script"
                    $result = [System.Windows.Forms.MessageBox]::Show($msg, "High Conflict Count", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    if ($result -eq 'No') {
                        throw "User aborted due to high conflict count."
                    } else {
                        $script:IgnoreConflictWarning = $true
                    }
                }

                # Attempt Auto-Rename (Suggestion)
                $renamed = $false
                for ($i = 2; $i -le 99; $i++) {
                    $candidate = "$sam$i"
                    try {
                        if ($CheckCmdlet -eq 'Get-ADUser') { $null = Get-ADUser -Identity $candidate -Server $TargetDomain -ErrorAction Stop }
                        elseif ($CheckCmdlet -eq 'Get-ADGroup') { $null = Get-ADGroup -Identity $candidate -Server $TargetDomain -ErrorAction Stop }
                        elseif ($CheckCmdlet -eq 'Get-ADComputer') { $null = Get-ADComputer -Identity $candidate -Server $TargetDomain -ErrorAction Stop }
                    } catch {
                        $targetSam = $candidate
                        $status = "AutoRenamed"
                        $action = "Create"
                        $notes = "Conflict detected. Suggested rename: $candidate"
                        $renamed = $true
                        break
                    }
                }

                if (-not $renamed) {
                    $status = "Conflict"
                    $action = "Skip" # Default safety
                    $notes = "Object exists in target. Auto-rename failed."
                }
            } else {
                # Not found, safe to create
                $status = "New"
            }

            # Add to Plan (Human readable)
            $MigrationPlan += [PSCustomObject]@{
                Type        = $ObjectType
                SourceSam   = $sam
                TargetSam   = $targetSam
                Status      = $status
                Action      = $action
                Description = $notes
            }

            # Add to Map (Machine readable for GPO migration)
            # We map SourceSID -> TargetSam because GPO Migration Tables use SIDs
            if ($sid) {
                $IdentityMap += [PSCustomObject]@{
                    SourceSam = $sam
                    SourceSID = $sid
                    TargetSam = $targetSam
                    Type      = $ObjectType
                }
            }
        }
    }

    # Process Users, Groups, and Computers
    Invoke-IdentityMapping -FileType "Users" -ObjectType "User" -CheckCmdlet "Get-ADUser"
    Invoke-IdentityMapping -FileType "Groups" -ObjectType "Group" -CheckCmdlet "Get-ADGroup"
    Invoke-IdentityMapping -FileType "Computers" -ObjectType "Computer" -CheckCmdlet "Get-ADComputer"

    # Output Plan
    $planFile = Join-Path $TransformPath "Identity_Migration_Plan.csv"
    $MigrationPlan | Export-Csv -Path $planFile -NoTypeInformation -Encoding UTF8
    
    # Output Map (Simple format for other scripts)
    $mapFile = Join-Path $config.TransformRoot "Mapping\Identity_Map_Final.csv"
    $IdentityMap | Export-Csv -Path $mapFile -NoTypeInformation -Encoding UTF8

    Write-Log -Message "Generated migration plan: $planFile" -Level INFO
    Write-Log -Message "Generated identity map: $mapFile" -Level INFO

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