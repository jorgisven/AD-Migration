<#
.SYNOPSIS
    Export OU structure from source domain.

.DESCRIPTION
    Retrieves the full OU hierarchy and exports it to CSV for mapping and reconstruction.
    Includes OU path, attributes, and description for analysis.
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
$ExportPath = Join-Path $config.ExportRoot 'OU_Structure'

# Ensure export directory exists
if (-not (Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
    Write-Log -Message "Created export directory: $ExportPath" -Level INFO
}

# Prompt for source domain if not provided
if (-not $SourceDomain) {
    $SourceDomain = Read-Host "Enter source domain name (e.g., source.local)"
}

Write-Log -Message "Starting OU export from domain: $SourceDomain" -Level INFO

try {
    # Get all OUs from source domain
    Invoke-Safely -ScriptBlock {
        $OUs = Get-ADOrganizationalUnit -Filter * -Server $SourceDomain -Properties *, ProtectedFromAccidentalDeletion | `
            Select-Object @{Name = 'OU'; Expression = { $_.Name }},
                          @{Name = 'DistinguishedName'; Expression = { $_.DistinguishedName }},
                          @{Name = 'CanonicalName'; Expression = { $_.CanonicalName }},
                          @{Name = 'Description'; Expression = { $_.Description }},
                          @{Name = 'Protected'; Expression = { $_.ProtectedFromAccidentalDeletion }},
                          @{Name = 'Created'; Expression = { $_.Created }},
                          @{Name = 'Modified'; Expression = { $_.Modified }} | `
            Sort-Object DistinguishedName
        
        # Export to CSV
        $outputFile = Join-Path $ExportPath "OU_Structure_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $OUs | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
        
        Write-Log -Message "Exported $($OUs.Count) OUs to $outputFile" -Level INFO
        Write-Host "OU export complete: $outputFile"
        
    } -Operation "Export OUs from $SourceDomain"
    
} catch {
    Write-Log -Message "Failed to export OUs: $_" -Level ERROR
    Write-Host "OU export failed. Check logs for details."
}