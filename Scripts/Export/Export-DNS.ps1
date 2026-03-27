<#
.SYNOPSIS
    Export DNS zones and records from source domain.

.DESCRIPTION
    Retrieves DNS zones and their resource records for migration planning.
    Requires RSAT-DNS-Server tools or access to the DnsServer module.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDomain,

    [Parameter(Mandatory = $false)]
    [string]$DnsServer
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

# Check environment context
$isElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$isDC = (Get-CimInstance Win32_OperatingSystem).ProductType -eq 2

if ($isDC -and -not $isElevated) {
    $elevMsg = "You are running on a Domain Controller without Administrator privileges.`n`nDNS export requires elevation to access local zones.`n`nPlease restart PowerShell as Administrator."
    $elevResult = [System.Windows.Forms.MessageBox]::Show($elevMsg, "Elevation Required", [System.Windows.Forms.MessageBoxButtons]::AbortRetryIgnore, [System.Windows.Forms.MessageBoxIcon]::Error)
    if ($elevResult -eq 'Abort') {
        Write-Host "Exiting per user request. Please run as Administrator." -ForegroundColor Red
        exit
    }
    Write-Log -Message "User chose to ignore elevation warning on DC." -Level WARN
} elseif (-not $isElevated) {
    Write-Log -Message "Script is not running as Administrator. Local DNS operations may fail." -Level WARN
}

try {
    Invoke-Safely -ScriptBlock {
        # Find a DC to query DNS (Preferably the PDC or a reliable DC)
        $DC = Get-ADDomainController -DomainName $SourceDomain -Discover
        $TargetDC = $DC.HostName
        # Determine Target DNS Server
        if ($DnsServer) {
            $TargetDC = $DnsServer
        } else {
            # Prefer PDC as it is most likely to be writable and authoritative
            try {
                $DC = Get-ADDomainController -DomainName $SourceDomain -Discover -Service PrimaryDC -ErrorAction Stop
            } catch {
                Write-Log -Message "Could not find PDC, falling back to any available DC." -Level WARN
                $DC = Get-ADDomainController -DomainName $SourceDomain -Discover
            }
            $TargetDC = $DC.HostName
        }

        # Check if running locally to avoid RPC loopback issues
        if ($TargetDC -and ($env:COMPUTERNAME -eq ($TargetDC -split '\.')[0])) {
             Write-Log -Message "Running on the target DC ($TargetDC). Switching to local mode." -Level INFO
             $TargetDC = $null
        }

        Write-Log -Message "Querying DNS Server: $(if ($TargetDC) { $TargetDC } else { 'Localhost' })" -Level INFO

        # Get Zones
        try {
            $zoneParams = @{ ErrorAction = 'Stop' }
            if ($TargetDC) { $zoneParams['ComputerName'] = $TargetDC }
            $Zones = Get-DnsServerZone @zoneParams
        } catch {
            Write-Log -Message "RPC query to $TargetDC failed. Attempting local DNS query..." -Level WARN
            # Fallback to local query (omitting ComputerName)
            $Zones = Get-DnsServerZone -ErrorAction Stop
            $TargetDC = $null # Switch to local mode for records
        }
        
        $ZoneList = @()
        $AllRecords = @()
        $staleRecordsSkipped = 0

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
                $dnsParams = @{ ZoneName = $ZoneName; ErrorAction = 'Stop' }
                if ($TargetDC) { $dnsParams['ComputerName'] = $TargetDC }
                $Records = Get-DnsServerResourceRecord @dnsParams

                foreach ($Rec in $Records) {
                    # Stale Record Filter (Pre-Export)
                    if ($FilterStale -and $Rec.Timestamp) {
                        $recAge = (Get-Date) - $Rec.Timestamp
                        if ($recAge.TotalDays -gt $StaleDays) {
                            $staleRecordsSkipped++
                            continue
                        }
                    }

                    $recordDataValue = $null
                    if ($Rec.RecordData) {
                        switch ($Rec.RecordType) {
                            "A"     { $recordDataValue = $Rec.RecordData.IPv4Address.ToString() }
                            "AAAA"  { $recordDataValue = $Rec.RecordData.IPv6Address.ToString() }
                            "CNAME" { $recordDataValue = $Rec.RecordData.HostNameAlias }
                            "TXT"   { $recordDataValue = $Rec.RecordData.DescriptiveText -join " " }
                            "PTR"   { $recordDataValue = $Rec.RecordData.PtrDomainName }
                            "MX"    { $recordDataValue = "$($Rec.RecordData.Preference) $($Rec.RecordData.MailExchange)" }
                            "NS"    { $recordDataValue = $Rec.RecordData.NameServer }
                            "SOA"   { $recordDataValue = "$($Rec.RecordData.PrimaryServer)" }
                            default { $recordDataValue = $Rec.RecordData.ToString() }
                        }
                    }

                    $AllRecords += [PSCustomObject]@{
                        ZoneName   = $ZoneName
                        HostName   = $Rec.HostName
                        RecordType = $Rec.RecordType
                        Data       = $recordDataValue
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

        if ($FilterStale) {
            Write-Log -Message "Exported $($ZoneList.Count) zones and $($AllRecords.Count) records to $ExportPath (stale records skipped: $staleRecordsSkipped; threshold: $StaleDays days)." -Level INFO
            Write-Host "DNS export complete: $($ZoneList.Count) zones, $($AllRecords.Count) records, $staleRecordsSkipped stale records skipped."
        } else {
            Write-Log -Message "Exported $($ZoneList.Count) zones and $($AllRecords.Count) records to $ExportPath" -Level INFO
            Write-Host "DNS export complete: $($ZoneList.Count) zones, $($AllRecords.Count) records."
        }
        
    } -Operation "Export DNS from $SourceDomain"
    
} catch {
    Write-Log -Message "DNS Export Failed: $_" -Level ERROR
    throw "DNS export failed. Check logs."
}