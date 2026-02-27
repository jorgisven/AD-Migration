<#
.SYNOPSIS
    Prepare WMI filters for import into target domain.

.DESCRIPTION
    Normalizes WMI filter queries and prepares recreation scripts.
    Checks for domain-specific strings in WMI queries that might need updating.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDomain,

    [Parameter(Mandatory = $false)]
    [string]$TargetDomain,

    [Parameter(Mandatory = $false)]
    [hashtable]$CustomReplacements
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

$config = Get-ADMigrationConfig
$SourceWMIPath = Join-Path $config.ExportRoot 'WMI_Filters'
$RebuildPath = Join-Path $config.TransformRoot 'WMI_Rebuild'

# Ensure transform directory exists
if (-not (Test-Path $RebuildPath)) { New-Item -ItemType Directory -Path $RebuildPath -Force | Out-Null }

Write-Log -Message "Starting WMI Filter Transformation..." -Level INFO

try {
    # Find latest export
    $wmiFiles = Get-ChildItem -Path $SourceWMIPath -Filter "WMI_Filters_*.csv" | Sort-Object LastWriteTime -Descending
    
    if (-not $wmiFiles) {
        $msg = "No WMI filter export found. If source had 0 filters, this is expected."
        Write-Log -Message $msg -Level WARN
        
        $noteFile = Join-Path $RebuildPath "No_WMI_Filters_To_Transform.txt"
        Set-Content -Path $noteFile -Value $msg
        Write-Host "No WMI filters found to transform. Created info file: $noteFile" -ForegroundColor Yellow
        return
    }

    $Filters = Import-Csv $wmiFiles[0].FullName
    if ($Filters.Count -eq 0) {
        $msg = "WMI export file is empty. No filters to transform."
        Write-Log -Message $msg -Level INFO
        
        $noteFile = Join-Path $RebuildPath "No_WMI_Filters_To_Transform.txt"
        Set-Content -Path $noteFile -Value $msg
        Write-Host "WMI export empty. Created info file: $noteFile" -ForegroundColor Yellow
        return
    }

    Write-Log -Message "Processing $($Filters.Count) WMI filters from $($wmiFiles[0].Name)" -Level INFO

    $TransformedFilters = @()

    foreach ($filter in $Filters) {
        $query = $filter.Query
        $name = $filter.FilterName
        $desc = $filter.FilterDescription
        $status = "Unchanged"

        # Simple replacement logic if domains are provided
        if ($SourceDomain -and $TargetDomain -and $query -match $SourceDomain) {
            $query = $query -replace [regex]::Escape($SourceDomain), $TargetDomain
            $status = "Modified"
            Write-Log -Message "Updated domain reference in filter '$name'" -Level INFO
        }

        # Apply custom replacements (e.g. server names, share paths)
        if ($CustomReplacements) {
            foreach ($key in $CustomReplacements.Keys) {
                if ($query -match [regex]::Escape($key)) {
                    $query = $query -replace [regex]::Escape($key), $CustomReplacements[$key]
                    $status = "Modified"
                    Write-Log -Message "Applied custom replacement '$key' -> '$($CustomReplacements[$key])' in filter '$name'" -Level INFO
                }
            }
        }

        $TransformedFilters += [PSCustomObject]@{
            Name        = $name
            Description = $desc
            Query       = $query
            OriginalDN  = $filter.DistinguishedName
            Status      = $status
        }
    }

    $outputFile = Join-Path $RebuildPath "WMI_Filters_Ready.csv"
    $TransformedFilters | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

    Write-Log -Message "Transformation complete. Saved to $outputFile" -Level INFO
    Write-Host "WMI Transform Complete: $outputFile"

} catch {
    Write-Log -Message "Failed to transform WMI filters: $_" -Level ERROR
}