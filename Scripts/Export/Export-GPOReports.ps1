<#
.SYNOPSIS
    Export GPO XML reports from source domain.

.DESCRIPTION
    Generates XML reports for all GPOs, including links, WMI filters, and settings.
#>

$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'GPO_Reports'

# ensure export folder exists
if (-not (Test-Path $ExportPath)) {
    New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
}

Invoke-Safely -Operation 'Export GPO reports' -ScriptBlock {
    # retrieve all GPOs and write XML report for each
    $gpos = Get-GPO -All
    foreach ($gpo in $gpos) {
        $fileName = ($gpo.DisplayName -replace '[\\/:\*?"<>\|]','_') + '.xml'
        $reportPath = Join-Path $ExportPath $fileName
        Get-GPOReport -Guid $gpo.Id -ReportType Xml -Path $reportPath
    }
}

Write-Log -Message 'GPO reports exported' -Level INFO
