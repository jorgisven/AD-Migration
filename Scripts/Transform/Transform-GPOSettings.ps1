<#
.SYNOPSIS
    Analyze GPO XML reports for domain-specific references.

.DESCRIPTION
    Scans exported GPO XML files for hardcoded paths (UNC), script locations, 
    and domain-specific group references that need rewriting.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

$config = Get-ADMigrationConfig
$SourceGPOPath = Join-Path $config.ExportRoot 'GPO_Reports'
$RewritePath = Join-Path $config.TransformRoot 'GPO_Rewrites'

# Ensure transform directory exists
if (-not (Test-Path $RewritePath)) { New-Item -ItemType Directory -Path $RewritePath -Force | Out-Null }

if (-not $SourceDomain) {
    $SourceDomain = Read-Host "Enter source domain name to search for (e.g. source.local)"
}

Write-Log -Message "Starting GPO analysis for domain references: $SourceDomain" -Level INFO

try {
    $xmlFiles = Get-ChildItem -Path $SourceGPOPath -Filter "*.xml"
    
    if ($xmlFiles.Count -eq 0) {
        Write-Log -Message "No GPO XML reports found in $SourceGPOPath. Run Export-GPOReports.ps1 first." -Level WARN
        return
    }

    $findings = @()

    foreach ($file in $xmlFiles) {
        Write-Log -Message "Analyzing $($file.Name)..." -Level INFO
        
        # Read XML content as text for broad searching first
        $content = Get-Content -Path $file.FullName -Raw
        
        # 1. Search for Source Domain references
        if ($content -match [regex]::Escape($SourceDomain)) {
            $findings += [PSCustomObject]@{
                GPOName = ($file.Name -split '_')[0]
                Type    = "DomainReference"
                Detail  = "Contains reference to $SourceDomain"
                File    = $file.Name
            }
        }

        # 2. Search for UNC Paths
        $uncMatches = [regex]::Matches($content, '\\\\[a-zA-Z0-9\-\.]+\\[a-zA-Z0-9\-\.\\\$]+')
        foreach ($match in $uncMatches) {
            $findings += [PSCustomObject]@{
                GPOName = ($file.Name -split '_')[0]
                Type    = "UNCPath"
                Detail  = $match.Value
                File    = $file.Name
            }
        }
    }

    $reportFile = Join-Path $RewritePath "GPO_Analysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $findings | Select-Object GPOName, Type, Detail, File | Export-Csv -Path $reportFile -NoTypeInformation

    Write-Log -Message "Analysis complete. Found $($findings.Count) potential issues." -Level INFO
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    Write-Host "GPO Analysis generated: $reportFile"
    Write-Host "Review this CSV to see which GPOs contain hardcoded paths" -ForegroundColor Yellow
    Write-Host "or legacy domain references." -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan

} catch {
    Write-Log -Message "Failed to analyze GPOs: $_" -Level ERROR
}