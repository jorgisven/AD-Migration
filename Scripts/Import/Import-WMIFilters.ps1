<#
.SYNOPSIS
    Recreate WMI filters in target domain.

.DESCRIPTION
    Imports normalized WMI filter definitions and links them to GPOs.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
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
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

$config = Get-ADMigrationConfig
$RebuildPath = Join-Path $config.TransformRoot 'WMI_Rebuild'

# Find transformed file
$wmiFiles = Get-ChildItem -Path $RebuildPath -Filter "WMI_Filters_Ready.csv" | Sort-Object LastWriteTime -Descending

if (-not $wmiFiles) {
    # Check if we have the "No Filters" marker file
    if (Test-Path (Join-Path $RebuildPath "No_WMI_Filters_To_Transform.txt")) {
        Write-Log -Message "No WMI filters to import (as per Transform phase)." -Level INFO
        return
    }
    Write-Log -Message "No transformed WMI filters found in $RebuildPath. Run Transform-WMIFilters.ps1 first." -Level ERROR
    throw "Missing WMI Transform Data"
}

$Filters = Import-Csv $wmiFiles[0].FullName
if (-not $TargetDomain) { $TargetDomain = (Get-ADDomain).DNSRoot }

Write-Log -Message "Starting WMI Filter Import to $TargetDomain..." -Level INFO

$script:WmiStats = [ordered]@{ Evaluated = 0; Created = 0; Skipped = 0; Failed = 0 }

Invoke-Safely -ScriptBlock {
    # WMI Filters are stored in the System container
    $rootDSE = Get-ADRootDSE -Server $TargetDomain
    $wmiContainer = "CN=WMIPolicy,CN=System,$($rootDSE.defaultNamingContext)"
    
    foreach ($row in $Filters) {
        $script:WmiStats.Evaluated++
        $name = $row.Name
        $desc = $row.Description
        $query = $row.Query
        
        # Check if exists
        $existing = Get-ADObject -Filter "objectClass -eq 'msWMISom' -and msWMI-Name -eq '$name'" -SearchBase $wmiContainer -Server $TargetDomain -ErrorAction SilentlyContinue
        
        if ($existing) {
            Write-Log -Message "WMI Filter '$name' already exists. Skipping." -Level WARN
            $script:WmiStats.Skipped++
        } else {
            # Create new WMI Filter Object
            $params = @{
                Name = [guid]::NewGuid().ToString("B") # WMI Filters use GUID as CN
                Path = $wmiContainer
                Type = "msWMISom"
                OtherAttributes = @{
                    'msWMI-Name' = $name
                    'msWMI-Description' = $desc
                    'msWMI-Parm1' = $desc
                    'msWMI-Parm2' = $query
                    'msWMI-IntFormatDecimal' = 2 # Standard version
                }
            }
            try {
                if ($PSCmdlet.ShouldProcess($name, "Create WMI Filter") -and -not $WhatIfPreference) {
                    New-ADObject @params -Server $TargetDomain -ErrorAction Stop
                    Write-Log -Message "Created WMI Filter: $name" -Level INFO
                    $script:WmiStats.Created++
                }
            } catch {
                Write-Log -Message "Failed to create WMI Filter '$name': $_" -Level ERROR
                $script:WmiStats.Failed++
            }
        }
    }
} -Operation "Import WMI Filters"

$warningCount = $script:WmiStats.Failed
$summary = "Import WMI Filters summary: Evaluated=$($script:WmiStats.Evaluated), Created=$($script:WmiStats.Created), Skipped=$($script:WmiStats.Skipped), Failed=$($script:WmiStats.Failed)"

if ($warningCount -gt 0) {
    Write-Host "[!] WARNING: WMI Filter Import encountered $warningCount failure(s). See logs for details." -ForegroundColor Yellow
    Write-Log -Message "Import WMI Filters completed with warnings. $summary" -Level WARN
} else {
    Write-Log -Message "Import WMI Filters completed. $summary" -Level INFO
}