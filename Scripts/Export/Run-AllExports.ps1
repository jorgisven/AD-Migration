<#
.SYNOPSIS
    Orchestrate the full export process for a specific domain.

.DESCRIPTION
    Runs all individual export scripts (OUs, GPOs, WMI, Accounts) in sequence against a target domain.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDomain
)

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# List of export scripts to run
$scripts = @(
    "Export-OUs.ps1",
    "Export-GPOReports.ps1",
    "Export-WMIFilters.ps1",
    "Export-AccountData.ps1",
    "Export-ACLs.ps1"
)

Write-Host "=== Starting Full Export for Domain: $SourceDomain ===" -ForegroundColor Cyan

foreach ($script in $scripts) {
    $scriptPath = Join-Path $ScriptRoot $script
    
    if (Test-Path $scriptPath) {
        Write-Host "`n[+] Running $script..." -ForegroundColor Green
        try {
            # Execute script with SourceDomain parameter
            & $scriptPath -SourceDomain $SourceDomain
        } catch {
            Write-Host "[-] Error running $($script): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "[-] Script not found: $scriptPath" -ForegroundColor Red
    }
}

Write-Host "`n=== Export Sequence Complete ===" -ForegroundColor Cyan