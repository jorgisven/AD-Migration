<#
.SYNOPSIS
    Validate that all required exports exist and contain data.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$ExportRoot = $config.ExportRoot

$requiredExports = @(
    @{ Name="Users"; Path="Security"; Pattern="Users_*.csv" },
    @{ Name="Computers"; Path="Security"; Pattern="Computers_*.csv" },
    @{ Name="OUs"; Path="OU_Structure"; Pattern="OU_Structure_*.csv" },
    @{ Name="GPOs"; Path="GPO_Reports"; Pattern="_GPO_Summary_*.csv" }
)

Write-Host "`n=== Validating Export Data ===" -ForegroundColor Cyan
$allValid = $true

foreach ($req in $requiredExports) {
    $dir = Join-Path $ExportRoot $req.Path
    $files = Get-ChildItem -Path $dir -Filter $req.Pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    
    if ($files) {
        $latest = $files[0]
        $data = Import-Csv $latest.FullName
        $count = @($data).Count
        
        if ($count -gt 0) {
            Write-Host " [OK] $($req.Name): Found $($count) records in $($latest.Name)" -ForegroundColor Green
        } else {
            Write-Host " [!!] $($req.Name): File exists but is EMPTY ($($latest.Name))" -ForegroundColor Red
            $allValid = $false
        }
    } else {
        Write-Host " [!!] $($req.Name): MISSING in $dir" -ForegroundColor Red
        $allValid = $false
    }
}

# Check WMI separately (0 is acceptable)
$wmiDir = Join-Path $ExportRoot "WMI_Filters"
$wmiFiles = Get-ChildItem -Path $wmiDir -Filter "WMI_Filters_*.csv" -ErrorAction SilentlyContinue
if ($wmiFiles) {
    $count = (Import-Csv $wmiFiles[0].FullName).Count
    Write-Host " [OK] WMI Filters: Found $count filters" -ForegroundColor Green
} else {
    Write-Host " [OK] WMI Filters: No export file found (Acceptable if source has 0 filters)" -ForegroundColor Yellow
}

Write-Host "`nValidation Complete."