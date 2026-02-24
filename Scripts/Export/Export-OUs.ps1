<#
.SYNOPSIS
    Export OU structure from source domain.

.DESCRIPTION
    Retrieves the full OU hierarchy and exports it to CSV for mapping and reconstruction.
#>

$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'OU_Structure'

# ensure export folder exists
if (-not (Test-Path $ExportPath)) {
    New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
}

Invoke-Safely -Operation 'Export OU structure' -ScriptBlock {
    Get-ADOrganizationalUnit -Filter * -Properties * |
        Select-Object DistinguishedName,Name,Description,WhenCreated,WhenChanged,ProtectedFromAccidentalDeletion |
        Export-Csv -Path (Join-Path $ExportPath 'OUs.csv') -NoTypeInformation -Force
}

Write-Log -Message 'OU structure exported' -Level INFO
