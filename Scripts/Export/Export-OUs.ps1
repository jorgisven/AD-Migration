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

# Add GUI support
Add-Type -AssemblyName System.Windows.Forms

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
        
        # Check for Empty OUs
        Write-Log -Message "Checking for empty OUs..." -Level INFO
        $EmptyLinkedOUs = @()
        $EmptyUnlinkedOUs = @()
        $FinalOUs = @()

        foreach ($ou in $OUs) {
            # Check if OU has any children (OneLevel search, stop at 1 result for speed)
            $hasChildren = Get-ADObject -Filter * -SearchBase $ou.DistinguishedName -SearchScope OneLevel -ResultSetSize 1 -Server $SourceDomain
            if (-not $hasChildren) {
                if (-not [string]::IsNullOrEmpty($ou.gPLink)) {
                    $EmptyLinkedOUs += $ou.DistinguishedName
                } else {
                    $EmptyUnlinkedOUs += $ou.DistinguishedName
                }
            }
            $FinalOUs += $ou
        }

        $OUsToSkip = @()

        # Prompt for Empty Linked OUs (Warning)
        if ($EmptyLinkedOUs.Count -gt 0) {
            $msg = "Found $($EmptyLinkedOUs.Count) empty OUs that have GPOs linked to them.`n`nSkipping these may break GPO links in the target.`n`nDo you want to EXPORT them?"
            $result = [System.Windows.Forms.MessageBox]::Show($msg, "Empty Linked OUs Detected", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            
            if ($result -eq 'No') {
                Write-Log -Message "User selected SKIP empty linked OUs." -Level WARN
                $OUsToSkip += $EmptyLinkedOUs
            } else {
                Write-Log -Message "User selected EXPORT empty linked OUs." -Level INFO
            }
        }

        # Prompt for Empty Unlinked OUs (Standard)
        if ($EmptyUnlinkedOUs.Count -gt 0) {
            $msg = "Found $($EmptyUnlinkedOUs.Count) empty OUs with NO GPO links.`n`nDo you want to EXPORT them?"
            $result = [System.Windows.Forms.MessageBox]::Show($msg, "Empty Unlinked OUs Detected", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            
            if ($result -eq 'No') {
                Write-Log -Message "User selected SKIP empty unlinked OUs." -Level INFO
                $OUsToSkip += $EmptyUnlinkedOUs
            } else {
                Write-Log -Message "User selected EXPORT empty unlinked OUs." -Level INFO
            }
        }

        # Filter FinalOUs
        if ($OUsToSkip.Count -gt 0) {
             $FinalOUs = $FinalOUs | Where-Object { $_.DistinguishedName -notin $OUsToSkip }
        }

        # Save Empty OU list for GPO Export script to use later
        $AllEmptyOUs = $EmptyLinkedOUs + $EmptyUnlinkedOUs
        $emptyListFile = Join-Path $ExportPath "EmptyOUs.txt"
        $AllEmptyOUs | Out-File -FilePath $emptyListFile -Encoding UTF8

        # Export to CSV
        $outputFile = Join-Path $ExportPath "OU_Structure_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $FinalOUs | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
        
        Write-Log -Message "Exported $($FinalOUs.Count) OUs to $outputFile" -Level INFO
        Write-Host "OU export complete: $outputFile"
        
    } -Operation "Export OUs from $SourceDomain"
    
} catch {
    Write-Log -Message "Failed to export OUs: $_" -Level ERROR
    throw "OU export failed. Check logs for details."
}
