<#
.SYNOPSIS
    Generates a GPO migration table (.migtable).
.DESCRIPTION
    Scans GPO XML reports for security principals (SIDs) and UNC paths, then creates a
    .migtable file. This file is used by Import-GPO to map old values to new ones.
    It pre-populates mappings from the Identity_Map_Final.csv where possible.
#>

[CmdletBinding()]
param(
    [string]$SourceDomain,
    [string]$TargetDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$ReportPath = Join-Path $config.ExportRoot 'GPO_Reports'
$MapPath = Join-Path $config.TransformRoot 'Mapping'
$TransformPath = Join-Path $config.TransformRoot 'GPO_Analysis' # Save table here

if (-not (Test-Path $TransformPath)) { New-Item -ItemType Directory -Path $TransformPath -Force | Out-Null }

Write-Log -Message "Starting GPO Migration Table generation..." -Level INFO

# 1. Load Identity Map (for pre-populating)
$IdentityMap = @{}
$idMapFile = Join-Path $MapPath "Identity_Map_Final.csv"
if (Test-Path $idMapFile) {
    Import-Csv $idMapFile | ForEach-Object {
        if (-not [string]::IsNullOrWhiteSpace($_.SourceSID)) {
            $IdentityMap[$_.SourceSID] = $_.TargetSam
        }
    }
    Write-Log -Message "Loaded $($IdentityMap.Count) entries from Identity Map." -Level INFO
}

# 2. Scan GPO XMLs for principals and UNC paths
$xmlFiles = Get-ChildItem -Path $ReportPath -Filter "*.xml"
if (-not $xmlFiles) { throw "Missing GPO Reports in '$ReportPath'. Run Export-GPOReports.ps1 first." }

$uniqueSIDs = [System.Collections.Generic.HashSet[string]]::new()
$uniqueUNCs = [System.Collections.Generic.HashSet[string]]::new()

foreach ($file in $xmlFiles) {
    [xml]$xml = Get-Content $file.FullName

    # Find all SIDs (from security filtering, user rights, restricted groups, etc.)
    $xml.SelectNodes("//SID|//Member/SID") | ForEach-Object { $uniqueSIDs.Add($_.InnerText) | Out-Null }
    
    # Find all UNC paths
    # This regex finds paths like \\server\share, allowing for spaces in the share name and underscores in the server name.
    # It is intentionally not greedy and stops at the first share level, which is what the migration table requires.
    $uncMatches = [regex]::Matches($xml.OuterXml, "\\\\[a-zA-Z0-9\.\-_]+\\[a-zA-Z0-9`$_\.\- ]+")
    $uncMatches | ForEach-Object { $uniqueUNCs.Add($_.Value) | Out-Null }
}

Write-Log -Message "Found $($uniqueSIDs.Count) unique SIDs and $($uniqueUNCs.Count) unique UNC paths in GPO reports." -Level INFO

# 3. Build the Migration Table XML
$doc = New-Object System.Xml.XmlDocument
$xmlDeclaration = $doc.CreateXmlDeclaration("1.0", "utf-8", $null)
$doc.AppendChild($xmlDeclaration) | Out-Null

$root = $doc.CreateElement("MigrationTable")
$root.SetAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
$root.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
$root.SetAttribute("xmlns", "http://www.microsoft.com/GroupPolicy/GPOOperations/MigrationTable")
$doc.AppendChild($root) | Out-Null

# Add SID mappings
foreach ($sid in $uniqueSIDs) {
    $entry = $doc.CreateElement("Mapping")
    
    $source = $doc.CreateElement("Source"); $source.SetAttribute("xsi:type", "GPMTrustee"); $source.SetAttribute("Name", $sid); $source.SetAttribute("SID", $sid)
    $entry.AppendChild($source) | Out-Null

    $destination = $doc.CreateElement("Destination")
    if ($IdentityMap.ContainsKey($sid)) {
        $targetName = $IdentityMap[$sid]
        # If TargetDomain is provided, try to qualify the name
        if ($TargetDomain -and $targetName -notmatch "\\") {
            $shortDomain = $TargetDomain.Split('.')[0]
            $targetName = "$shortDomain\$targetName"
        }
        $destination.SetAttribute("xsi:type", "GPMTrustee"); $destination.SetAttribute("Name", $targetName); $destination.SetAttribute("SID", "")
    } else {
        $destination.SetAttribute("xsi:type", "GPMTrustee"); $destination.SetAttribute("Name", ""); $destination.SetAttribute("SID", "")
    }
    $entry.AppendChild($destination) | Out-Null
    
    $root.AppendChild($entry) | Out-Null
}

# Add UNC Path mappings
foreach ($unc in $uniqueUNCs) {
    $entry = $doc.CreateElement("Mapping")
    
    $source = $doc.CreateElement("Source"); $source.SetAttribute("xsi:type", "GPMPath"); $source.SetAttribute("Path", $unc)
    $entry.AppendChild($source) | Out-Null

    $destination = $doc.CreateElement("Destination"); $destination.SetAttribute("xsi:type", "GPMPath"); $destination.SetAttribute("Path", "")
    $entry.AppendChild($destination) | Out-Null
    
    $root.AppendChild($entry) | Out-Null
}

# 4. Save the file
$outFile = Join-Path $MapPath "MigrationTable.migtable"
$doc.Save($outFile)

Write-Host "[+] Migration Table generated successfully." -ForegroundColor Green
Write-Host "File saved to: $outFile" -ForegroundColor Green
Write-Host "Please review and complete the empty <Destination> entries in this file before running the GPO import." -ForegroundColor Yellow
Write-Log -Message "Migration Table saved to $outFile" -Level INFO