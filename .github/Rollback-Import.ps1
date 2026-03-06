<#
.SYNOPSIS
    Rollback imported objects from the Target Domain.

.DESCRIPTION
    Reverses the actions of the Import phase by deleting objects created during migration.
    Uses the mapping files and import logs to identify specific objects to remove.
    
    WARNING: This is destructive. Use -WhatIf to preview changes.

.PARAMETER Scope
    Specifies what to rollback: 'Accounts', 'GPOs', 'WMIFilters', 'OUs', or 'All'.
#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Accounts', 'GPOs', 'WMIFilters', 'OUs', 'All')]
    [string]$Scope,

    [Parameter(Mandatory = $false)]
    [string]$TargetDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO

$config = Get-ADMigrationConfig
$TransformPath = $config.TransformRoot
$MapPath = Join-Path $TransformPath 'Mapping'

if (-not $TargetDomain) { $TargetDomain = (Get-ADDomain).DNSRoot }

Write-Log -Message "Starting Rollback ($Scope) in $TargetDomain..." -Level WARN

# -----------------------------------------------------------------------------
# 1. Rollback Accounts (Users, Groups, Computers)
# -----------------------------------------------------------------------------
if ($Scope -eq 'Accounts' -or $Scope -eq 'All') {
    $idMapFile = Join-Path $MapPath "Identity_Map_Final.csv"
    
    if (Test-Path $idMapFile) {
        $Identities = Import-Csv $idMapFile
        Write-Log -Message "Found $($Identities.Count) accounts to rollback." -Level INFO

        foreach ($id in $Identities) {
            try {
                $targetSam = $id.TargetSam
                $type = $id.Type
                
                if ($PSCmdlet.ShouldProcess("$targetSam ($type)", "Delete Object")) {
                    switch ($type) {
                        'User'     { Remove-ADUser -Identity $targetSam -Server $TargetDomain -Confirm:$false -ErrorAction Stop }
                        'Group'    { Remove-ADGroup -Identity $targetSam -Server $TargetDomain -Confirm:$false -ErrorAction Stop }
                        'Computer' { Remove-ADComputer -Identity $targetSam -Server $TargetDomain -Confirm:$false -ErrorAction Stop }
                    }
                    Write-Log -Message "Deleted $type : $targetSam" -Level INFO
                }
            } catch {
                Write-Log -Message "Failed to delete $($id.TargetSam): $_" -Level WARN
            }
        }
    } else {
        Write-Log -Message "Identity_Map_Final.csv not found. Skipping Account rollback." -Level WARN
    }
}

# -----------------------------------------------------------------------------
# 2. Rollback GPOs
# -----------------------------------------------------------------------------
if ($Scope -eq 'GPOs' -or $Scope -eq 'All') {
    $BackupPath = Join-Path $config.ExportRoot 'GPO_Backups'
    $manifestPath = Join-Path $BackupPath "manifest.xml"

    if (Test-Path $manifestPath) {
        [xml]$manifest = Get-Content $manifestPath
        $backups = $manifest.Backups.BackupInst
        Write-Log -Message "Found $($backups.Count) GPOs to check for rollback." -Level INFO

        foreach ($b in $backups) {
            $gpoName = $b.GPODisplayName
            
            # Check if exists
            if (Get-GPO -Name $gpoName -Server $TargetDomain -ErrorAction SilentlyContinue) {
                try {
                    if ($PSCmdlet.ShouldProcess($gpoName, "Delete GPO")) {
                        Remove-GPO -Name $gpoName -Server $TargetDomain -Confirm:$false -ErrorAction Stop
                        Write-Log -Message "Deleted GPO: $gpoName" -Level INFO
                    }
                } catch {
                    Write-Log -Message "Failed to delete GPO $gpoName : $_" -Level WARN
                }
            }
        }
    } else {
        Write-Log -Message "GPO Backup manifest not found. Skipping GPO rollback." -Level WARN
    }
}

# -----------------------------------------------------------------------------
# 3. Rollback WMI Filters
# -----------------------------------------------------------------------------
if ($Scope -eq 'WMIFilters' -or $Scope -eq 'All') {
    $RebuildPath = Join-Path $config.TransformRoot 'WMI_Rebuild'
    $wmiFiles = Get-ChildItem -Path $RebuildPath -Filter "WMI_Filters_Ready.csv" | Sort-Object LastWriteTime -Descending

    if ($wmiFiles) {
        $Filters = Import-Csv $wmiFiles[0].FullName
        $rootDSE = Get-ADRootDSE -Server $TargetDomain
        $wmiContainer = "CN=WMIPolicy,CN=System,$($rootDSE.defaultNamingContext)"

        foreach ($row in $Filters) {
            $name = $row.Name
            try {
                $filterObj = Get-ADObject -Filter "objectClass -eq 'msWMISom' -and msWMI-Name -eq '$name'" -SearchBase $wmiContainer -Server $TargetDomain -ErrorAction SilentlyContinue
                if ($filterObj) {
                    if ($PSCmdlet.ShouldProcess($name, "Delete WMI Filter")) {
                        Remove-ADObject -Identity $filterObj.DistinguishedName -Server $TargetDomain -Confirm:$false -ErrorAction Stop
                        Write-Log -Message "Deleted WMI Filter: $name" -Level INFO
                    }
                }
            } catch {
                Write-Log -Message "Failed to delete WMI Filter $name : $_" -Level WARN
            }
        }
    }
}

# -----------------------------------------------------------------------------
# 4. Rollback OUs
# -----------------------------------------------------------------------------
if ($Scope -eq 'OUs' -or $Scope -eq 'All') {
    $ouMapFile = Get-ChildItem -Path $MapPath -Filter "OU_Map_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    
    if ($ouMapFile) {
        $OUs = Import-Csv $ouMapFile.FullName
        # Sort by length of DN descending to delete children before parents
        $SortedOUs = $OUs | Sort-Object { $_.TargetDN.Length } -Descending

        foreach ($row in $SortedOUs) {
            if ($row.Action -ne 'Skip') {
                $targetDN = $row.TargetDN
                try {
                    if (Get-ADOrganizationalUnit -Identity $targetDN -Server $TargetDomain -ErrorAction SilentlyContinue) {
                        if ($PSCmdlet.ShouldProcess($targetDN, "Delete OU (Recursive)")) {
                            # Disable accidental deletion protection first
                            Set-ADOrganizationalUnit -Identity $targetDN -ProtectedFromAccidentalDeletion $false -Server $TargetDomain
                            Remove-ADOrganizationalUnit -Identity $targetDN -Recursive -Server $TargetDomain -Confirm:$false -ErrorAction Stop
                            Write-Log -Message "Deleted OU: $targetDN" -Level INFO
                        }
                    }
                } catch {
                    Write-Log -Message "Failed to delete OU $targetDN : $_" -Level WARN
                }
            }
        }
    } else {
        Write-Log -Message "OU Map not found. Skipping OU rollback." -Level WARN
    }
}

Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "Rollback Complete"