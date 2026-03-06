<#
.SYNOPSIS
    Validate GPO existence and linkage in the Target Domain.

.DESCRIPTION
    Compares the Source GPO Reports (XML) against the Target Domain to ensure:
    1. GPOs exist.
    2. GPOs are linked to the correct Target OUs (based on OU Map).
    3. WMI Filters are attached.
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
$ReportPath = Join-Path $config.ExportRoot 'GPO_Reports'
$MapPath = Join-Path $config.TransformRoot 'Mapping'

if (-not $TargetDomain) { $TargetDomain = (Get-ADDomain).DNSRoot }

# Load OU Map for Link Translation
$ouMapFile = Get-ChildItem -Path $MapPath -Filter "OU_Map_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$OUMap = @{}
if ($ouMapFile) {
    Import-Csv $ouMapFile.FullName | ForEach-Object { $OUMap[$_.SourceDN] = $_.TargetDN }
}

$xmlFiles = Get-ChildItem -Path $ReportPath -Filter "*.xml"
if (-not $xmlFiles) { throw "No GPO Reports found." }

Write-Log -Message "Starting GPO Integrity Check against $TargetDomain..." -Level INFO

$Errors = [System.Collections.Generic.List[PSObject]]::new()

foreach ($xmlFile in $xmlFiles) {
    [xml]$xml = Get-Content $xmlFile.FullName
    $gpoName = $xml.GPO.Name
    
    # 1. Check GPO Existence
    $gpo = Get-GPO -Name $gpoName -Server $TargetDomain -ErrorAction SilentlyContinue
    if (-not $gpo) {
        $Errors.Add([PSCustomObject]@{ GPO = $gpoName; Type = "GPO"; Status = "Missing"; Details = "GPO object not found" })
        Write-Log -Message "Missing GPO: $gpoName" -Level ERROR
        continue
    }

    # 2. Check WMI Filter
    $sourceWmi = $xml.GPO.FilterData.Name
    if ($sourceWmi) {
        if (-not $gpo.WmiFilter) {
            $Errors.Add([PSCustomObject]@{ GPO = $gpoName; Type = "WMI"; Status = "Missing Link"; Details = "Expected filter '$sourceWmi' not linked" })
        } elseif ($gpo.WmiFilter.Name -ne $sourceWmi) {
            $Errors.Add([PSCustomObject]@{ GPO = $gpoName; Type = "WMI"; Status = "Mismatch"; Details = "Expected '$sourceWmi', found '$($gpo.WmiFilter.Name)'" })
        }
    }

    # 3. Check Links
    # Get current links for this GPO in Target
    # Note: Get-GPOReport is heavy, we can infer links by querying the SOMs (OUs)
    # Easier approach: Iterate expected links, check if they exist on the Target OU
    
    foreach ($link in $xml.GPO.LinksTo) {
        $sourceSOM = $link.SOMPath
        $targetSOM = $null

        if ($OUMap.ContainsKey($sourceSOM)) {
            $targetSOM = $OUMap[$sourceSOM]
        } elseif ($sourceSOM -match "DC=") {
            $targetSOM = $TargetDomain # Domain root assumption
        }

        if ($targetSOM) {
            # Check if GPO is linked to this Target OU
            $linksOnOU = Get-GPInheritance -Target $targetSOM -Server $TargetDomain -ErrorAction SilentlyContinue
            $isLinked = $false
            if ($linksOnOU.GpoLinks) {
                $isLinked = bool
            }

            if (-not $isLinked) {
                $Errors.Add([PSCustomObject]@{ GPO = $gpoName; Type = "Link"; Status = "Missing"; Details = "Not linked to $targetSOM" })
                Write-Log -Message "Missing Link: $gpoName -> $targetSOM" -Level WARN
            }
        }
    }
}

Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
if ($Errors.Count -gt 0) {
    Write-Host "GPO Validation Failed with $($Errors.Count) issues." -ForegroundColor Red
    $Errors | Format-Table -AutoSize
} else {
    Write-Host "GPO Validation Success." -ForegroundColor Green
}
Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan