<#
.SYNOPSIS
    Import DNS zones and records into the target domain.

.DESCRIPTION
    Reads the transformed DNS CSV files and recreates Zones and Resource Records.
    Supports A, AAAA, CNAME, TXT, MX, and PTR records.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$TargetServer,

    [Parameter(Mandatory = $false)]
    [string]$TargetDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

# GUI Setup for Conflict Resolution
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-ConflictDialog {
    param(
        [string]$Message,
        [string]$Title
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(450, 200)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(10, 10)
    $lbl.Size = New-Object System.Drawing.Size(410, 60)
    $lbl.Text = $Message
    $form.Controls.Add($lbl)

    $chkAll = New-Object System.Windows.Forms.CheckBox
    $chkAll.Text = "Apply choice to all remaining conflicts of this type"
    $chkAll.Location = New-Object System.Drawing.Point(10, 80)
    $chkAll.Size = New-Object System.Drawing.Size(410, 20)
    $form.Controls.Add($chkAll)

    $btnSkip = New-Object System.Windows.Forms.Button
    $btnSkip.Text = "Skip"
    $btnSkip.Location = New-Object System.Drawing.Point(120, 110)
    $btnSkip.DialogResult = [System.Windows.Forms.DialogResult]::No
    $form.Controls.Add($btnSkip)

    $btnOverwrite = New-Object System.Windows.Forms.Button
    $btnOverwrite.Text = "Overwrite"
    $btnOverwrite.Location = New-Object System.Drawing.Point(210, 110)
    $btnOverwrite.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $form.Controls.Add($btnOverwrite)
    
    $form.AcceptButton = $btnOverwrite
    $form.CancelButton = $btnSkip
    $result = $form.ShowDialog()

    return @{ Action = if ($result -eq 'Yes') { 'Overwrite' } else { 'Skip' }; ApplyAll = $chkAll.Checked }
}

$config = Get-ADMigrationConfig
$TransformPath = Join-Path $config.TransformRoot 'DNS'
$ExportPath = Join-Path $config.ExportRoot 'DNS'

# Determine Source Path (Prefer Transform, fallback to Export)
if (Test-Path $TransformPath) {
    $SourcePath = $TransformPath
    Write-Log -Message "Using Transformed DNS data from: $SourcePath" -Level INFO
} elseif (Test-Path $ExportPath) {
    $SourcePath = $ExportPath
    Write-Log -Message "Using Raw Export DNS data from: $SourcePath" -Level INFO
} else {
    throw "No DNS data found in Export or Transform folders."
}

# Determine Target Server (supports Run-AllImports passing -TargetDomain)
if ([string]::IsNullOrWhiteSpace($TargetServer) -and -not [string]::IsNullOrWhiteSpace($TargetDomain)) {
    $TargetServer = $TargetDomain
    Write-Log -Message "Import-DNS received TargetDomain '$TargetDomain'; using it as TargetServer." -Level INFO
}
if ([string]::IsNullOrWhiteSpace($TargetServer)) {
    $TargetServer = $env:COMPUTERNAME
}

Write-Log -Message "Starting DNS Import to Server: $TargetServer" -Level INFO

# Global conflict preferences (Prompt, Overwrite, Skip)
$GlobalZoneAction = "Prompt"
$GlobalRecordAction = "Prompt"

Invoke-Safely -ScriptBlock {
    # Check for DNS Module
    if (-not (Get-Command Add-DnsServerPrimaryZone -ErrorAction SilentlyContinue)) {
        throw "DNS Server module not found. Please install RSAT-DNS-Server."
    }

    # 1. Find latest Zone and Record files
    $zoneFile = Get-ChildItem -Path $SourcePath -Filter "DNS_Zones_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $recFile  = Get-ChildItem -Path $SourcePath -Filter "DNS_Records_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if (-not $zoneFile -or -not $recFile) {
        throw "Missing DNS CSV files in $SourcePath"
    }

    $Zones = Import-Csv $zoneFile.FullName
    $Records = Import-Csv $recFile.FullName

    Write-Log -Message "Found $($Zones.Count) zones and $($Records.Count) records to process." -Level INFO

    # 2. Process Zones
    foreach ($z in $Zones) {
        $zName = $z.ZoneName
        
        # Skip system zones if they weren't filtered earlier
        if ($zName -match "^_msdcs" -or $zName -eq "." -or $zName -eq "TrustAnchors") {
            continue
        }

        try {
            $exists = Get-DnsServerZone -Name $zName -ComputerName $TargetServer -ErrorAction SilentlyContinue
            if ($exists) {
                # Handle Zone Conflict
                $action = "Skip"
                if ($GlobalZoneAction -eq "Prompt") {
                    $res = Show-ConflictDialog -Message "Zone '$zName' already exists on $TargetServer.`nOverwrite will DELETE the existing zone and recreate it." -Title "Zone Conflict"
                    $action = $res.Action
                    if ($res.ApplyAll) { $GlobalZoneAction = $action }
                } else {
                    $action = $GlobalZoneAction
                }

                if ($action -eq "Overwrite") {
                    Write-Log -Message "Overwriting Zone: $zName" -Level WARN
                    Remove-DnsServerZone -Name $zName -ComputerName $TargetServer -Force -Confirm:$false -ErrorAction Stop
                    # Re-create logic falls through to below? No, need to re-run creation code.
                    # Simpler to just delete here and let the 'else' block or a flag handle creation, 
                    # but since the 'else' is structured for creation, we can just proceed to create if we deleted it.
                    $exists = $false 
                } else {
                    Write-Log -Message "Skipping existing Zone: $zName" -Level INFO
                }
            } else {
                # Create Zone
                if ($z.IsDsIntegrated -eq 'True') {
                    Add-DnsServerPrimaryZone -Name $zName -ReplicationScope "Domain" -ComputerName $TargetServer -ErrorAction Stop
                } else {
                    Add-DnsServerPrimaryZone -Name $zName -ZoneFile "$zName.dns" -ComputerName $TargetServer -ErrorAction Stop
                }
                Write-Log -Message "Created Zone: $zName" -Level INFO
            }
            
            # If we overwrote (deleted), we need to create it now
            if ($action -eq "Overwrite" -and -not $exists) {
                 if ($z.IsDsIntegrated -eq 'True') {
                    Add-DnsServerPrimaryZone -Name $zName -ReplicationScope "Domain" -ComputerName $TargetServer -ErrorAction Stop
                } else {
                    Add-DnsServerPrimaryZone -Name $zName -ZoneFile "$zName.dns" -ComputerName $TargetServer -ErrorAction Stop
                }
                Write-Log -Message "Re-created Zone: $zName" -Level INFO
            }
        } catch {
            Write-Log -Message "Failed to create zone '$zName': $_" -Level ERROR
        }
    }

    # 3. Process Records
    foreach ($r in $Records) {
        $zName = $r.ZoneName
        $hostName = $r.HostName
        $type = $r.RecordType
        $data = $r.Data
        
        # Skip SOA and NS as they are auto-generated
        if ($type -eq "SOA" -or $type -eq "NS") { continue }

        $retry = $true
        while ($retry) {
            $retry = $false
            try {
                switch ($type) {
                    "A"     { Add-DnsServerResourceRecordA -ZoneName $zName -Name $hostName -IPv4Address $data -ComputerName $TargetServer -ErrorAction Stop }
                    "AAAA"  { Add-DnsServerResourceRecordAAAA -ZoneName $zName -Name $hostName -IPv6Address $data -ComputerName $TargetServer -ErrorAction Stop }
                    "CNAME" { Add-DnsServerResourceRecordCName -ZoneName $zName -Name $hostName -HostNameAlias $data -ComputerName $TargetServer -ErrorAction Stop }
                    "TXT"   { Add-DnsServerResourceRecord -ZoneName $zName -Name $hostName -Txt -DescriptiveText $data -ComputerName $TargetServer -ErrorAction Stop }
                    "PTR"   { Add-DnsServerResourceRecordPtr -ZoneName $zName -Name $hostName -PtrDomainName $data -ComputerName $TargetServer -ErrorAction Stop }
                    "MX"    { 
                        # Parse "10 mail.example.com."
                        if ($data -match "^(\d+)\s+(.+)$") {
                            Add-DnsServerResourceRecordMX -ZoneName $zName -Name $hostName -Preference $matches[1] -MailExchange $matches[2] -ComputerName $TargetServer -ErrorAction Stop
                        }
                    }
                }
            } catch {
                $errMsg = $_.Exception.Message
                # Check for conflict errors (Already exists, or CNAME conflict)
                if ($errMsg -like "*already exists*" -or $errMsg -like "*CNAME record cannot coexist*") {
                    
                    $action = "Skip"
                    if ($GlobalRecordAction -eq "Prompt") {
                        $res = Show-ConflictDialog -Message "Record Conflict in zone '$zName'.`nRecord: $hostName ($type)`nError: $errMsg" -Title "Record Conflict"
                        $action = $res.Action
                        if ($res.ApplyAll) { $GlobalRecordAction = $action }
                    } else {
                        $action = $GlobalRecordAction
                    }

                    if ($action -eq "Overwrite") {
                        Write-Log -Message "Overwriting Record: $hostName ($type) in $zName" -Level INFO
                        try {
                            # Remove existing record(s) with this name to clear the path
                            # Note: This removes ALL records with this name (e.g. all A records for round robin)
                            Get-DnsServerResourceRecord -ZoneName $zName -Name $hostName -ComputerName $TargetServer -ErrorAction SilentlyContinue | 
                                Remove-DnsServerResourceRecord -ZoneName $zName -ComputerName $TargetServer -Force -Confirm:$false -ErrorAction Stop
                            $retry = $true # Retry the Add
                        } catch {
                            Write-Log -Message "Failed to remove conflicting record '$hostName': $_" -Level ERROR
                        }
                    } else {
                        Write-Log -Message "Skipping conflicting record: $hostName ($type)" -Level INFO
                    }
                } else {
                    Write-Log -Message "Failed to add $type record '$hostName' in '$zName': $errMsg" -Level WARN
                }
            }
        }
    }
    
    Write-Host "DNS Import Complete."

} -Operation "Import DNS Records"