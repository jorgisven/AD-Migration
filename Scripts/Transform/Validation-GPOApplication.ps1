<#
.SYNOPSIS
    Validate GPO linking and WMI filter dependencies prior to migration.

.DESCRIPTION
    Static analysis tool that simulates GPO link reconstruction based on the OU Map.
    Reports on:
    - GPO Links that will be dropped (Source OU not mapped).
    - GPO Links targeting OUs marked as 'Skip'.
    - WMI Filters referenced by GPOs that are missing from the transformation list.
#>

param()

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO

$config = Get-ADMigrationConfig
$ReportPath = Join-Path $config.ExportRoot 'GPO_Reports'
$MapPath = Join-Path $config.TransformRoot 'Mapping'
$WMIPath = Join-Path $config.TransformRoot 'WMI_Rebuild'
$AnalysisPath = Join-Path $config.TransformRoot 'GPO_Validation'

if (-not (Test-Path $AnalysisPath)) { New-Item -ItemType Directory -Path $AnalysisPath -Force | Out-Null }

# 1. Load OU Map
$mapFiles = Get-ChildItem -Path $MapPath -Filter "OU_Map_*.csv" | Sort-Object LastWriteTime -Descending
if (-not $mapFiles) { throw "Missing OU Map. Run Transform-OUMap.ps1 first." }
$OUMap = @{}
$OUSkip = @{}
Import-Csv $mapFiles[0].FullName | ForEach-Object { 
    $OUMap[$_.SourceDN] = $_.TargetDN 
    if ($_.Action -eq 'Skip') { $OUSkip[$_.SourceDN] = $true }
}
Write-Log -Message "Loaded OU Map ($($OUMap.Count) entries)." -Level INFO

# 2. Load WMI Filters (Optional)
$WMIFilters = @{}
$wmiFiles = Get-ChildItem -Path $WMIPath -Filter "WMI_Filters_Ready.csv" | Sort-Object LastWriteTime -Descending
if ($wmiFiles) {
    Import-Csv $wmiFiles[0].FullName | ForEach-Object { $WMIFilters[$_.Name] = $true }
    Write-Log -Message "Loaded WMI Filter list." -Level INFO
}

# 3. Analyze GPO Reports
$xmlFiles = Get-ChildItem -Path $ReportPath -Filter "*.xml"
if (-not $xmlFiles) { throw "Missing GPO Reports. Run Export-GPOReports.ps1 first." }

$Results = @()

foreach ($xmlFile in $xmlFiles) {
    [xml]$xml = Get-Content $xmlFile.FullName
    $gpoName = $xml.GPO.Name
    
    # Check Links
    $links = $xml.GPO.LinksTo
    if ($links) {
        foreach ($link in $links) {
            $sourceSOM = $link.SOMPath
            $status = "OK"
            $notes = ""

            if ($sourceSOM -match "DC=") {
                # Check Map
                if ($OUMap.ContainsKey($sourceSOM)) {
                    if ($OUSkip.ContainsKey($sourceSOM)) {
                        $status = "Warning"
                        $notes = "Target OU is marked as SKIP in mapping."
                    } else {
                        $notes = "Maps to: " + $OUMap[$sourceSOM]
                    }
                } elseif ($sourceSOM -match "^DC=") {
                    $status = "Info"
                    $notes = "Domain Root link (will map to Target Root)."
                } else {
                    $status = "Broken"
                    $notes = "Source OU not found in Map."
                }
            } else {
                $status = "Ignore"
                $notes = "Site link or non-OU link."
            }

            $Results += [PSCustomObject]@{
                GPOName = $gpoName
                Type    = "Link"
                Item    = $sourceSOM
                Status  = $status
                Notes   = $notes
            }
        }
    }

    # Check WMI Filter
    $wmi = $xml.GPO.FilterData.Name
    if ($wmi) {
        $status = "OK"
        $notes = ""
        if (-not $WMIFilters.ContainsKey($wmi)) {
            $status = "Missing"
            $notes = "WMI Filter '$wmi' not found in transformed WMI list."
        }
        $Results += [PSCustomObject]@{
            GPOName = $gpoName
            Type    = "WMI"
            Item    = $wmi
            Status  = $status
            Notes   = $notes
        }
    }
}

# Output
$csvPath = Join-Path $AnalysisPath "GPO_Validation_Report.csv"
$Results | Export-Csv -Path $csvPath -NoTypeInformation

Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "GPO Validation Complete"
Write-Host "Report: $csvPath"
Write-Host "Found $(($Results | Where-Object {$_.Status -eq 'Broken'}).Count) Broken Links" -ForegroundColor Red
Write-Host "Found $(($Results | Where-Object {$_.Status -eq 'Missing'}).Count) Missing WMI Filters" -ForegroundColor Red
Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
