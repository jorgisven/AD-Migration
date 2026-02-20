<#
.SYNOPSIS
    Export user and service account attributes.

.DESCRIPTION
    Used for account reconciliation between source and target domains.
    Exports user and computer attributes needed for migration planning.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDomain
)

# Import module and config
Import-Module .\Scripts\ADMigration\ADMigration.psd1 -Force
$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'Security'

# Ensure export directory exists
if (-not (Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
    Write-Log -Message "Created export directory: $ExportPath" -Level INFO
}

# Prompt for source domain if not provided
if (-not $SourceDomain) {
    $SourceDomain = Read-Host "Enter source domain name (e.g., source.local)"
}

Write-Log -Message "Starting account data export from domain: $SourceDomain" -Level INFO

try {
    Invoke-Safely -ScriptBlock {
        # Export user accounts
        Write-Log -Message "Exporting user accounts from $SourceDomain" -Level INFO
        $Users = Get-ADUser -Filter * -Server $SourceDomain -Properties * | `
            Select-Object @{Name = 'SamAccountName'; Expression = { $_.SamAccountName }},
                          @{Name = 'UserPrincipalName'; Expression = { $_.UserPrincipalName }},
                          @{Name = 'DisplayName'; Expression = { $_.DisplayName }},
                          @{Name = 'GivenName'; Expression = { $_.GivenName }},
                          @{Name = 'Surname'; Expression = { $_.Surname }},
                          @{Name = 'DistinguishedName'; Expression = { $_.DistinguishedName }},
                          @{Name = 'Enabled'; Expression = { $_.Enabled }},
                          @{Name = 'LastLogonDate'; Expression = { $_.LastLogonDate }},
                          @{Name = 'PasswordLastSet'; Expression = { $_.PasswordLastSet }},
                          @{Name = 'WhenCreated'; Expression = { $_.WhenCreated }},
                          @{Name = 'WhenChanged'; Expression = { $_.WhenChanged }},
                          @{Name = 'AccountExpirationDate'; Expression = { $_.AccountExpirationDate }} | `
            Sort-Object SamAccountName
        
        $userFile = Join-Path $ExportPath "Users_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Users | Export-Csv -Path $userFile -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Exported $($Users.Count) user accounts to $userFile" -Level INFO
        
        # Export service accounts (users with specific naming patterns)
        Write-Log -Message "Exporting service accounts from $SourceDomain" -Level INFO
        $ServiceAccounts = Get-ADUser -Filter { (SamAccountName -like 'svc_*') -or (SamAccountName -like '*service*') } `
            -Server $SourceDomain -Properties * | `
            Select-Object @{Name = 'SamAccountName'; Expression = { $_.SamAccountName }},
                          @{Name = 'UserPrincipalName'; Expression = { $_.UserPrincipalName }},
                          @{Name = 'DisplayName'; Expression = { $_.DisplayName }},
                          @{Name = 'DistinguishedName'; Expression = { $_.DistinguishedName }},
                          @{Name = 'Enabled'; Expression = { $_.Enabled }},
                          @{Name = 'PasswordNotRequired'; Expression = { $_.PasswordNotRequired }},
                          @{Name = 'AccountNeverExpires'; Expression = { $_.AccountNeverExpires }} | `
            Sort-Object SamAccountName
        
        if ($ServiceAccounts.Count -gt 0) {
            $svcFile = Join-Path $ExportPath "ServiceAccounts_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $ServiceAccounts | Export-Csv -Path $svcFile -NoTypeInformation -Encoding UTF8
            Write-Log -Message "Exported $($ServiceAccounts.Count) service accounts to $svcFile" -Level INFO
        }
        
        # Export computer accounts
        Write-Log -Message "Exporting computer accounts from $SourceDomain" -Level INFO
        $Computers = Get-ADComputer -Filter * -Server $SourceDomain -Properties * | `
            Select-Object @{Name = 'ComputerName'; Expression = { $_.Name }},
                          @{Name = 'SamAccountName'; Expression = { $_.SamAccountName }},
                          @{Name = 'DistinguishedName'; Expression = { $_.DistinguishedName }},
                          @{Name = 'OperatingSystem'; Expression = { $_.OperatingSystem }},
                          @{Name = 'OperatingSystemVersion'; Expression = { $_.OperatingSystemVersion }},
                          @{Name = 'Enabled'; Expression = { $_.Enabled }},
                          @{Name = 'LastLogonDate'; Expression = { $_.LastLogonDate }},
                          @{Name = 'WhenCreated'; Expression = { $_.WhenCreated }},
                          @{Name = 'WhenChanged'; Expression = { $_.WhenChanged }} | `
            Sort-Object ComputerName
        
        $computerFile = Join-Path $ExportPath "Computers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Computers | Export-Csv -Path $computerFile -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Exported $($Computers.Count) computer accounts to $computerFile" -Level INFO
        
        # Summary
        $summary = @"
=== Account Export Summary ===
Domain: $SourceDomain
Exported: $(Get-Date)

Users: $($Users.Count)
Service Accounts: $($ServiceAccounts.Count)
Computers: $($Computers.Count)

Total Objects: $($Users.Count + $ServiceAccounts.Count + $Computers.Count)
"@
        
        Add-Content -Path (Join-Path $ExportPath 'EXPORT_SUMMARY.txt') -Value $summary
        
        Write-Host "✓ Account data export complete:"
        Write-Host "  - Users: $($Users.Count)"
        Write-Host "  - Service Accounts: $($ServiceAccounts.Count)"
        Write-Host "  - Computers: $($Computers.Count)"
        
    } -Operation "Export account data from $SourceDomain"
    
} catch {
    Write-Log -Message "Failed to export account data: $_" -Level ERROR
    Write-Host "✗ Account data export failed. Check logs for details."
}