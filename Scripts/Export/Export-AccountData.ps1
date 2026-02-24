<#
.SYNOPSIS
    Export user and service account attributes.

.DESCRIPTION
    Used for account reconciliation between source and target domains.
#>

$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'Security'

# ensure export folder exists
if (-not (Test-Path $ExportPath)) {
    New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
}

Invoke-Safely -Operation 'Export account data' -ScriptBlock {
    # export user accounts
    Get-ADUser -Filter * -Properties * |
        Select-Object SamAccountName,DistinguishedName,Enabled,PasswordLastSet,WhenCreated,WhenChanged,MemberOf |
        Export-Csv -Path (Join-Path $ExportPath 'Users.csv') -NoTypeInformation -Force

    # export managed and standalone service accounts
    Get-ADServiceAccount -Filter * |
        Select-Object Name,SamAccountName,DistinguishedName,Enabled,WhenCreated,WhenChanged |
        Export-Csv -Path (Join-Path $ExportPath 'ServiceAccounts.csv') -NoTypeInformation -Force
}

Write-Log -Message 'Account data exported' -Level INFO
