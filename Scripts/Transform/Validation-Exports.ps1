<#
.SYNOPSIS
    Validates that all required export files exist and are not empty.
.DESCRIPTION
    Checks the Export directory for key files generated during the export phase.
    This script is the first step in the Transform phase to ensure data is ready for processing.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$ExportRoot = $config.ExportRoot

Write-Log -Message "Starting validation of export files..." -Level INFO
$hasErrors = $false

function Test-ExportFile {
    param(
        [string]$Path,
        [string]$Pattern,
        [string]$Description,
        [switch]$AllowEmpty
    )
    
    $file = Get-ChildItem -Path $Path -Filter $Pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    
    if (-not $file) {
        Write-Host "[-] MISSING: $Description (Expected pattern: $Pattern in $Path)" -ForegroundColor Red
        Write-Log -Message "Validation FAILED: Missing $Description export file." -Level ERROR
        $script:hasErrors = $true
        return
    }

    if (-not $AllowEmpty -and $file.Length -lt 100) { # Arbitrary small size to detect empty/header-only files
        Write-Host "[!] WARNING: $Description file is very small or empty ($($file.Name))" -ForegroundColor Yellow
        Write-Log -Message "Validation WARNING: $Description file is very small or empty." -Level WARN
    } else {
        Write-Host "[+] FOUND: $Description ($($file.Name))" -ForegroundColor Green
        Write-Log -Message "Validation PASSED: Found $Description file." -Level INFO
    }
}

# Define expected files
$expectedFiles = @(
    @{ Path = (Join-Path $ExportRoot 'OU_Structure'); Pattern = 'OU_Structure_*.csv'; Description = 'OU Structure' },
    @{ Path = (Join-Path $ExportRoot 'GPO_Reports'); Pattern = '_GPO_Summary_*.csv'; Description = 'GPO Summary' },
    @{ Path = (Join-Path $ExportRoot 'WMI_Filters'); Pattern = 'WMI_Filters_*.csv'; Description = 'WMI Filters' },
    @{ Path = (Join-Path $ExportRoot 'Security'); Pattern = 'Users_*.csv'; Description = 'User Accounts' },
    @{ Path = (Join-Path $ExportRoot 'Security'); Pattern = 'Groups_*.csv'; Description = 'Groups' },
    @{ Path = (Join-Path $ExportRoot 'Security'); Pattern = 'Computers_*.csv'; Description = 'Computer Accounts' },
    @{ Path = (Join-Path $ExportRoot 'Security'); Pattern = 'GroupMembers_*.csv'; Description = 'Group Memberships' },
    @{ Path = (Join-Path $ExportRoot 'DNS'); Pattern = 'DNS_Zones_*.csv'; Description = 'DNS Zones'; AllowEmpty = $true },
    @{ Path = (Join-Path $ExportRoot 'DNS'); Pattern = 'DNS_Records_*.csv'; Description = 'DNS Records'; AllowEmpty = $true },
    @{ Path = (Join-Path $ExportRoot 'Security'); Pattern = 'ACLs_OUs_*.csv'; Description = 'OU ACLs'; AllowEmpty = $true }
)

foreach ($fileCheck in $expectedFiles) {
    Test-ExportFile @fileCheck
}

Write-Host ""
if ($hasErrors) {
    Write-Host "=== Validation FAILED ===" -ForegroundColor Red
    Write-Host "One or more critical export files are missing. Please re-run the Export phase." -ForegroundColor Red
} else {
    Write-Host "=== Validation Complete ===" -ForegroundColor Green
    Write-Host "All major export files are present."
}