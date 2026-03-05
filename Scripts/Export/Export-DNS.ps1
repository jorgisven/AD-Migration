<#
.SYNOPSIS
    Export DNS zones and records from source domain.

.DESCRIPTION
    Retrieves DNS zones and their resource records for migration planning.
    Requires RSAT-DNS-Server tools or access to the DnsServer module.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'DNS'

# GUI Prompt for Stale Records
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$FilterStale = $false
$StaleDays = 365

$msgResult = [System.Windows.Forms.MessageBox]::Show("Do you want to remove STALE DNS records from the export?`n`nYes = Remove Stale Records`nNo = Keep All Records", "DNS Export Options", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

if ($msgResult -eq 'Yes') {
    $FilterStale = $true
    $inputDays = [Microsoft.VisualBasic.Interaction]::InputBox("Enter age in days to consider stale (e.g. 365):", "Stale Record Definition", "365")
    if (-not [int]::TryParse($inputDays, [ref]$StaleDays)) {
        $StaleDays = 365
    }
}

# Ensure export directory exists
if (-not (Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
}

# Prompt for source domain if not provided
if (-not $SourceDomain) {
    $SourceDomain = Read-Host "Enter source domain name (e.g., source.local)"
}

Write-Log -Message "Starting DNS export from: $SourceDomain" -Level INFO

try {
    Invoke-Safely -ScriptBlock {
        # Find a DC to query DNS (Preferably the PDC or a reliable DC)
        $DC = Get-ADDomainController -DomainName $SourceDomain -Discover
        $TargetDC = $DC.HostName
        Write-Log -Message "Querying DNS Server: $TargetDC" -Level INFO

        # Get Zones
        $Zones = Get-DnsServerZone -ComputerName $TargetDC
        
        $ZoneList = @()
        $AllRecords = @()

        foreach ($Zone in $Zones) {
            $ZoneName = $Zone.ZoneName
            $ZoneList += [PSCustomObject]@{
                ZoneName       = $ZoneName
                ZoneType       = $Zone.ZoneType
                IsDsIntegrated = $Zone.IsDsIntegrated
                FileName       = $Zone.ZoneFile
            }

            # Export Records for this zone
            try {
                $Records = Get-DnsServerResourceRecord -ZoneName $ZoneName -ComputerName $TargetDC -ErrorAction Stop
                foreach ($Rec in $Records) {
                    # Stale Record Filter (Pre-Export)
                    if ($FilterStale -and $Rec.Timestamp) {
                        $recAge = (Get-Date) - $Rec.Timestamp
                        if ($recAge.TotalDays -gt $StaleDays) {
                            continue
                        }
                    }

                    $AllRecords += [PSCustomObject]@{
                        ZoneName   = $ZoneName
                        HostName   = $Rec.HostName
                        RecordType = $Rec.RecordType
                        Data       = if ($Rec.RecordData) { $Rec.RecordData.ToString() } else { $null }
                        Timestamp  = $Rec.Timestamp
                        TTL        = $Rec.TimeToLive
                    }
                }
            } catch {
                Write-Log -Message "Could not retrieve records for zone $ZoneName : $_" -Level WARN
            }
        }

        # Export Zones
        $zoneFile = Join-Path $ExportPath "DNS_Zones_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $ZoneList | Export-Csv -Path $zoneFile -NoTypeInformation -Encoding UTF8
        
        # Export Records
        $recFile = Join-Path $ExportPath "DNS_Records_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $AllRecords | Export-Csv -Path $recFile -NoTypeInformation -Encoding UTF8

        Write-Log -Message "Exported $($ZoneList.Count) zones and $($AllRecords.Count) records to $ExportPath" -Level INFO
        Write-Host "DNS export complete: $($ZoneList.Count) zones, $($AllRecords.Count) records."
        
    } -Operation "Export DNS from $SourceDomain"
    
} catch {
    Write-Log -Message "DNS Export Failed: $_" -Level ERROR
    Write-Host "DNS export failed. Check logs."
}