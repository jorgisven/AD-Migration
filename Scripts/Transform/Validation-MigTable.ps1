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
        $context = ""
        $srcType = $m.Source.Name
        $srcName = $m.Source.GetAttribute("Name")
        $srcPath = $m.Source.GetAttribute("Path")
        $gpoName = $m.Source.GetAttribute("GPOName")
        $gpoGuid = $m.Source.GetAttribute("GPOGuid")
        # Try to infer context
        if ($gpoName -or $gpoGuid) {
            $context = "GPO: $gpoName ($gpoGuid)"
        } elseif ($srcType -eq "Source" -and $srcPath -match "SYSVOL|\\Domain Controllers\\") {
            $context = "Built-in/System Path: $srcPath"
        } elseif ($srcName) {
            $context = "$srcType: $srcName"
        } elseif ($srcPath) {
            $context = "$srcType Path: $srcPath"
        } else {
            $context = "Unknown Source"
        }
        $unmapped += @{ Context = $context; Name = $srcName; Path = $srcPath }
    }
}

if ($unmapped.Count -gt 0) {
    Write-Host "[-] ERROR: Found $($unmapped.Count) unmapped entries in the Migration Table:" -ForegroundColor Red
    foreach ($u in $unmapped) {
        $msg = "  - "
        if ($u.Context) { $msg += $u.Context }
        if ($u.Name) { $msg += ", Name: $($u.Name)" }
        if ($u.Path) { $msg += ", Path: $($u.Path)" }
        Write-Host $msg -ForegroundColor Red
    }
    throw "Validation Failed"
} else {
    Write-Host "[+] PASSED: All Migration Table entries have a destination." -ForegroundColor Green
}