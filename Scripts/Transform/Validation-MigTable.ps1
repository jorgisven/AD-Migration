<#
.SYNOPSIS
    Validates the GPO Migration Table to ensure no empty destinations exist.
#>

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$MapPath = Join-Path $config.TransformRoot 'Mapping'
$MigTablePath = Join-Path $MapPath "MigrationTable.migtable"

if (-not (Test-Path $MigTablePath)) { throw "Migration Table not found at $MigTablePath" }

Write-Host "Validating Migration Table..." -ForegroundColor Cyan
[xml]$xml = Get-Content $MigTablePath
$unmapped = @()

foreach ($m in $xml.GetElementsByTagName("Mapping")) {
    $destName = $m.Destination.GetAttribute("Name")
    $destPath = $m.Destination.GetAttribute("Path")
    
    if ([string]::IsNullOrWhiteSpace($destName) -and [string]::IsNullOrWhiteSpace($destPath)) {
        $sourceVal = if ($m.Source.HasAttribute("Name")) { $m.Source.GetAttribute("Name") } else { $m.Source.GetAttribute("Path") }
        $unmapped += $sourceVal
    }
}

if ($unmapped.Count -gt 0) {
    Write-Host "[-] ERROR: Found $($unmapped.Count) unmapped entries in the Migration Table:" -ForegroundColor Red
    foreach ($u in $unmapped) { Write-Host "  - $u" -ForegroundColor Red }
    throw "Validation Failed"
} else {
    Write-Host "[+] PASSED: All Migration Table entries have a destination." -ForegroundColor Green
}