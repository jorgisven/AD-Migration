<#
.SYNOPSIS
    Export GPO XML reports from source domain.

.DESCRIPTION
    Generates HTML and XML reports for all GPOs, including links, WMI filters, and settings.
    Reports are useful for analyzing policies before migration.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDomain
)

# Import module and config
Import-Module .\Scripts\ADMigration\ADMigration.psd1 -Force
$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'GPO_Reports'

# Ensure export directory exists
if (-not (Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
    Write-Log -Message "Created export directory: $ExportPath" -Level INFO
}

# Prompt for source domain if not provided
if (-not $SourceDomain) {
    $SourceDomain = Read-Host "Enter source domain name (e.g., source.local)"
}

Write-Log -Message "Starting GPO report export from domain: $SourceDomain" -Level INFO

try {
    Invoke-Safely -ScriptBlock {
        # Get all GPOs from source domain
        $GPOs = Get-GPO -All -Domain $SourceDomain
        
        Write-Log -Message "Found $($GPOs.Count) GPOs to export" -Level INFO
        
        # Generate reports for each GPO
        $reportCount = 0
        foreach ($GPO in $GPOs) {
            try {
                # Generate HTML report
                $htmlPath = Join-Path $ExportPath "$($GPO.DisplayName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
                Get-GPOReport -Guid $GPO.Id -ReportType Html -Path $htmlPath -Server $SourceDomain
                
                # Also create a simple CSV entry for quick reference
                [PSCustomObject]@{
                    GPOName            = $GPO.DisplayName
                    GPOID              = $GPO.Id
                    CreationTime       = $GPO.CreationTime
                    ModificationTime   = $GPO.ModificationTime
                    Owner              = $GPO.Owner
                    HTMLReport         = (Split-Path $htmlPath -Leaf)
                } | Export-Csv -Path (Join-Path $ExportPath '_GPO_Summary.csv') -NoTypeInformation -Append -Encoding UTF8
                
                $reportCount++
            } catch {
                Write-Log -Message "Failed to export report for GPO: $($GPO.DisplayName)" -Level WARN
            }
        }
        
        Write-Log -Message "Exported $reportCount GPO reports to $ExportPath" -Level INFO
        Write-Host "✓ GPO report export complete: $reportCount reports generated"
        
    } -Operation "Export GPO reports from $SourceDomain"
    
} catch {
    Write-Log -Message "Failed to export GPO reports: $_" -Level ERROR
    Write-Host "✗ GPO report export failed. Check logs for details."
}