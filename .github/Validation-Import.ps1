<#
.SYNOPSIS
    Validate that OUs and Accounts have been correctly imported into the Target Domain.

.DESCRIPTION
    Checks the Target Domain against the Transform plans (OU Map, User Plan) and the 
    Import execution log (Identity Map) to verify object existence and status.
#>

param(
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

Write-Log -Message "Starting Import Validation against $TargetDomain..." -Level INFO

$ValidationErrors = [System.Collections.Generic.List[PSObject]]::new()

# 1. Validate OUs
$ouMapFile = Get-ChildItem -Path $MapPath -Filter "OU_Map_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($ouMapFile) {
    Write-Log -Message "Validating OUs from $($ouMapFile.Name)..." -Level INFO
    $OUs = Import-Csv $ouMapFile.FullName
    foreach ($row in $OUs) {
        if ($row.Action -ne 'Skip') {
            if (-not (Get-ADOrganizationalUnit -Identity $row.TargetDN -Server $TargetDomain -ErrorAction SilentlyContinue)) {
                $ValidationErrors.Add([PSCustomObject]@{
                    Type = "OU"
                    Name = $row.TargetDN
                    Status = "Missing"
                    Details = "Expected from OU Map but not found in Target."
                })
                Write-Log -Message "Missing OU: $($row.TargetDN)" -Level ERROR
            }
        }
    }
} else {
    Write-Log -Message "No OU Map found to validate." -Level WARN
}

# 2. Validate Accounts (Users/Groups)
# Prefer Identity_Map_Final.csv (Actual Import Results) over Plan (Theoretical)
$idMapFile = Join-Path $MapPath "Identity_Map_Final.csv"

if (Test-Path $idMapFile) {
    Write-Log -Message "Validating Accounts from Identity Map..." -Level INFO
    $Identities = Import-Csv $idMapFile
    
    foreach ($id in $Identities) {
        $exists = $false
        try {
            if ($id.Type -eq 'User') {
                $exists = bool
            } elseif ($id.Type -eq 'Group') {
                $exists = bool
            } elseif ($id.Type -eq 'Computer') {
                $exists = bool
            }
        } catch {
            $exists = $false
        }

        if (-not $exists) {
            $ValidationErrors.Add([PSCustomObject]@{
                Type = $id.Type
                Name = $id.TargetSam
                Status = "Missing"
                Details = "Listed in Identity Map but not found in Target."
            })
            Write-Log -Message "Missing Account: $($id.TargetSam)" -Level ERROR
        }
    }
} else {
    Write-Log -Message "Identity_Map_Final.csv not found. Skipping account validation." -Level WARN
}

Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "Validation Complete"
if ($ValidationErrors.Count -gt 0) {
    Write-Host "FAILED: Found $($ValidationErrors.Count) missing objects." -ForegroundColor Red
    $ValidationErrors | Format-Table -AutoSize
    
    $errFile = Join-Path $config.TransformRoot "Validation_Errors_Import.csv"
    $ValidationErrors | Export-Csv -Path $errFile -NoTypeInformation
    Write-Host "Errors saved to $errFile"
} else {
    Write-Host "SUCCESS: All checked objects exist in Target Domain." -ForegroundColor Green
}
Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan