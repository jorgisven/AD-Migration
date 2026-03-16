<#
.SYNOPSIS
    Generate OU mapping file from exported source OUs.

.DESCRIPTION
    Reads the latest OU export and creates a CSV template for mapping source OUs to target OUs.
    Administrators should edit the output CSV to define the target structure.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$TargetBaseDN
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) {
    throw "ADMigration module manifest missing, cannot continue."
}
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) {
    throw "Invoke-Safely unavailable"
}

$config = Get-ADMigrationConfig
$SourceOUPath = Join-Path $config.ExportRoot 'OU_Structure'
$MapPath = Join-Path $config.TransformRoot 'Mapping'

# Ensure transform directory exists
if (-not (Test-Path $MapPath)) { New-Item -ItemType Directory -Path $MapPath -Force | Out-Null }

Invoke-Safely -ScriptBlock {
    # Find latest export file
    $latestExport = Get-ChildItem -Path $SourceOUPath -Filter "OU_Structure_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if (-not $latestExport) {
        Write-Log -Message "No OU export files found in $SourceOUPath" -Level ERROR
        throw "Missing OU export data. Run Export-OUs.ps1 first."
    }

    Write-Log -Message "Generating OU map from $($latestExport.Name)" -Level INFO

    $OUs = @(Import-Csv -Path $latestExport.FullName)
    $MappingData = @()

    # Helper to escape special characters in DN (e.g. "Sales, West" -> "Sales\, West")
    # Characters to escape: , \ # + < > ; " =
    $specialChars = '([,#"<>;=\\])'

    foreach ($ou in $OUs) {
        # Create a draft mapping entry
        # By default, we suggest mapping to the same name, but mark it as 'Migrate'
        
        $escapedName = $ou.OU -replace $specialChars, '\$1'
        
        $targetDN = "" # Leave blank to force mapping in the GUI/Excel
        if ($TargetBaseDN) {
            $targetDN = "OU=$escapedName,$TargetBaseDN"
        }

        $MappingData += [PSCustomObject]@{
            Action        = "Migrate" # Options: Migrate, Merge, Skip, CreateNew
            SourceOU      = $ou.OU
            TargetOU      = $ou.OU
            TargetDN      = $targetDN
            SourceDN      = $ou.DistinguishedName
            Description   = $ou.Description
        }
    }

    # Sort by length of TargetDN to ensure parents are created before children
    $MappingData = $MappingData | Sort-Object { $_.TargetDN.Length }

    $mapFile = Join-Path $MapPath "OU_Map_Draft.csv"
    $MappingData | Export-Csv -Path $mapFile -NoTypeInformation -Encoding UTF8

    Write-Log -Message "Generated draft OU map at $mapFile" -Level INFO
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "OU Mapping draft created: $mapFile"
    Write-Host "ACTION REQUIRED: Open this CSV and edit the 'TargetDN' column" -ForegroundColor Yellow
    Write-Host "to match your desired structure in the new domain." -ForegroundColor Yellow
    Write-Host "TIP: Run Validation-AccountPlacement.ps1 after editing to check naming conventions." -ForegroundColor Gray
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
} -Operation "Generate OU Map"