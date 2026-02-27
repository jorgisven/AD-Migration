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
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) {
    throw "ADMigration module manifest missing, cannot continue."
}
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) {
    Write-Log -Message "Required function Invoke-Safely not defined after module import" -Level ERROR
    throw "Invoke-Safely unavailable"
}
$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'GPO_Reports'
$BackupPath = Join-Path $config.ExportRoot 'GPO_Backups'

# Ensure export directory exists
if (-not (Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
}
if (-not (Test-Path $BackupPath)) {
    New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
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
        
        $summaryFile = Join-Path $ExportPath "_GPO_Summary_${SourceDomain}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        # Generate reports for each GPO
        $reportCount = 0
        foreach ($GPO in $GPOs) {
            try {
                $safeName = $GPO.DisplayName -replace '[\\/:"*?<>|]', '_'

                # Generate HTML report
                $htmlPath = Join-Path $ExportPath "${safeName}.html"
                Get-GPOReport -Guid $GPO.Id -ReportType Html -Path $htmlPath -Server $SourceDomain
                
                # Generate XML report (Required for Transform phase)
                $xmlPath = Join-Path $ExportPath "${safeName}.xml"
                Get-GPOReport -Guid $GPO.Id -ReportType Xml -Path $xmlPath -Server $SourceDomain

                # Perform actual GPO Backup (Required for Import)
                Backup-GPO -Guid $GPO.Id -Path $BackupPath -Server $SourceDomain | Out-Null

                # Also create a simple CSV entry for quick reference
                [PSCustomObject]@{
                    GPOName            = $GPO.DisplayName
                    GPOID              = $GPO.Id
                    CreationTime       = $GPO.CreationTime
                    ModificationTime   = $GPO.ModificationTime
                    Owner              = $GPO.Owner
                    HTMLReport         = (Split-Path $htmlPath -Leaf)
                    XMLReport          = (Split-Path $xmlPath -Leaf)
                } | Export-Csv -Path $summaryFile -NoTypeInformation -Append -Encoding UTF8
                
                $reportCount++
            } catch {
                Write-Log -Message "Failed to export report for GPO: $($GPO.DisplayName)" -Level WARN
            }
        }
        
        Write-Log -Message "Exported $reportCount GPO reports to $ExportPath" -Level INFO
        Write-Host "GPO report export complete: $reportCount reports generated"
        
    } -Operation "Export GPO reports from $SourceDomain"
    
} catch {
    Write-Log -Message "Failed to export GPO reports: $_" -Level ERROR
    Write-Host "GPO report export failed. Check logs for details."
}