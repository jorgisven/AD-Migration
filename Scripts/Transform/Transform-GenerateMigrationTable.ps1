<#
.SYNOPSIS
    Generate a GPO Migration Table for mapping security principals and paths.

.DESCRIPTION
    Scans exported GPO reports (XML) for references to the source domain (users, groups, UNC paths).
    Creates a CSV mapping file. This file can be reviewed and then converted into a 
    standard .migtable file for use with Import-GPO.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO

$config = Get-ADMigrationConfig
$GPOReportPath = Join-Path $config.ExportRoot 'GPO_Reports'
$TransformPath = Join-Path $config.TransformRoot 'Mapping'

# Ensure transform directory exists
if (-not (Test-Path $TransformPath)) { New-Item -ItemType Directory -Path $TransformPath -Force | Out-Null }

Write-Log -Message "Starting Migration Table generation..." -Level INFO

if (-not (Test-Path $GPOReportPath)) {
    Write-Log -Message "GPO Reports folder not found at $GPOReportPath. Run Export-GPOReports.ps1 first." -Level ERROR
    throw "Missing GPO Reports"
}

$xmlFiles = Get-ChildItem -Path $GPOReportPath -Filter "*.xml"
if ($xmlFiles.Count -eq 0) {
    Write-Log -Message "No XML reports found in $GPOReportPath." -Level WARN
    return
}

Write-Log -Message "Scanning $($xmlFiles.Count) GPO reports..." -Level INFO

$Principals = @()
$UNCPaths = @()

foreach ($file in $xmlFiles) {
    try {
        # Load XML
        [xml]$xml = Get-Content $file.FullName -ErrorAction Stop

        # 1. Extract Trustees (Security Principals)
        # Look for <Trustee><Name>...</Name></Trustee>
        $trusteeNodes = $xml.SelectNodes("//Trustee/Name")
        foreach ($node in $trusteeNodes) {
            if (-not [string]::IsNullOrWhiteSpace($node.'#text')) {
                $Principals += $node.'#text'
            }
        }

        # 2. Extract UNC Paths
        # Regex scan raw content for \\Server\Share patterns as they can appear in various nodes
        $rawContent = Get-Content $file.FullName -Raw
        $uncMatches = [regex]::Matches($rawContent, '\\\\[a-zA-Z0-9\-\._]+\\[a-zA-Z0-9\-\._\$]+')
        foreach ($match in $uncMatches) {
            $UNCPaths += $match.Value
        }

    } catch {
        Write-Log -Message "Error parsing $($file.Name): $_" -Level WARN
    }
}

# Deduplicate and Sort
$UniquePrincipals = $Principals | Select-Object -Unique | Sort-Object
$UniqueUNCs = $UNCPaths | Select-Object -Unique | Sort-Object

$TableData = @()

# Add Principals
foreach ($p in $UniquePrincipals) {
    $TableData += [PSCustomObject]@{
        Type   = "Principal"
        Source = $p
        Target = "" # Placeholder for user input
        Notes  = "Detected in GPO Report"
    }
}

# Add UNCs
foreach ($u in $UniqueUNCs) {
    $TableData += [PSCustomObject]@{
        Type   = "UNCPath"
        Source = $u
        Target = "" # Placeholder for user input
        Notes  = "Detected in GPO Report"
    }
}

$outFile = Join-Path $TransformPath "GPO_MigrationTable_Draft.csv"
$TableData | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8

Write-Log -Message "Generated migration table draft with $($TableData.Count) entries at $outFile" -Level INFO

Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "Migration Table Draft Generated: $outFile"
Write-Host "Found $(($UniquePrincipals).Count) Principals and $(($UniqueUNCs).Count) UNC Paths."
Write-Host "ACTION REQUIRED: Edit the 'Target' column in the CSV to map values to the new domain." -ForegroundColor Yellow
Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan