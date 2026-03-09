<#
.SYNOPSIS
    Analyzes GPO settings for items requiring manual review.
.DESCRIPTION
    Scans GPO XML reports to find domain-specific settings like UNC paths,
    scripts, and security principals in User Rights Assignments or Restricted Groups.
    Outputs a CSV report for manual analysis.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$ReportPath = Join-Path $config.ExportRoot 'GPO_Reports'
$TransformPath = Join-Path $config.TransformRoot 'GPO_Analysis'

if (-not (Test-Path $TransformPath)) { New-Item -ItemType Directory -Path $TransformPath -Force | Out-Null }

$xmlFiles = Get-ChildItem -Path $ReportPath -Filter "*.xml"
if (-not $xmlFiles) { throw "Missing GPO Reports in '$ReportPath'. Run Export-GPOReports.ps1 first." }

Write-Log -Message "Starting GPO settings analysis..." -Level INFO
$AnalysisResults = [System.Collections.Generic.List[PSObject]]::new()

$progress = 0
foreach ($file in $xmlFiles) {
    $progress++
    Write-Progress -Activity "Analyzing GPO Reports" -Status "Processing $($file.Name)" -PercentComplete (($progress / $xmlFiles.Count) * 100)

    [xml]$xml = Get-Content $file.FullName
    $gpoName = $xml.GPO.Name
    
    # Regex to find UNC paths like \\server\share, allowing for spaces in the share name and underscores in the server name.
    $uncRegex = "\\\\[a-zA-Z0-9\.\-_]+\\[a-zA-Z0-9`$_\.\- ]+"

    # 1. Find UNC Paths in Scripts (Logon, Logoff, etc.)
    $xml.SelectNodes("//Script") | ForEach-Object {
        $uncMatches = [regex]::Matches($_.Exec, $uncRegex)
        foreach ($match in $uncMatches) {
            $AnalysisResults.Add([PSCustomObject]@{
                GPOName     = $gpoName; Category    = "Scripts"
                SettingName = $_.Name;   Value       = $match.Value
                Finding     = "UNC Path"; Notes       = "Verify path is accessible from target domain."
            })
        }
    }

    # 2. Find User Rights Assignments
    $xml.SelectNodes("//Policy[Name='User Rights Assignment']/Value/Assignment") | ForEach-Object {
        $rightName = $_.Name
        $_.Member | ForEach-Object {
            $AnalysisResults.Add([PSCustomObject]@{
                GPOName     = $gpoName
                Category    = "User Rights Assignment"
                SettingName = $rightName
                Value       = $_.Name
                Finding     = "Security Principal"
                Notes       = "Ensure this principal is mapped to the target domain."
            })
        }
    }

    # 3. Find Restricted Groups
    $xml.SelectNodes("//Memberof/Group") | ForEach-Object {
        $groupName = $_.Name
        $_.Members | ForEach-Object {
            $AnalysisResults.Add([PSCustomObject]@{
                GPOName     = $gpoName
                Category    = "Restricted Groups"
                SettingName = "Members of '$groupName'"
                Value       = $_.Name
                Finding     = "Security Principal"
                Notes       = "Ensure this principal is mapped to the target domain."
            })
        }
    }
    
    # 4. Find Drive Mappings
    $xml.SelectNodes("//DriveMapSettings/Drive") | ForEach-Object {
        $uncMatches = [regex]::Matches($_.path, $uncRegex)
        foreach ($match in $uncMatches) {
            $AnalysisResults.Add([PSCustomObject]@{
                GPOName     = $gpoName
                Category    = "Drive Mappings"
                SettingName = "Drive $($_.letter)"
                Value       = $match.Value
                Finding     = "UNC Path"
                Notes       = "Verify path is accessible from target domain."
            })
        }
    }
}

if ($AnalysisResults.Count -gt 0) {
    $outFile = Join-Path $TransformPath "GPO_Analysis_Report.csv"
    $AnalysisResults | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8
    Write-Host "[+] GPO Analysis Complete. Found $($AnalysisResults.Count) items for review." -ForegroundColor Green
    Write-Host "Report saved to: $outFile" -ForegroundColor Green
    Write-Log -Message "GPO Analysis complete. Report saved to $outFile" -Level INFO
} else {
    Write-Host "[+] GPO Analysis Complete. No specific items flagged for manual review." -ForegroundColor Green
    Write-Log -Message "GPO Analysis complete. No specific items flagged for manual review." -Level INFO
}