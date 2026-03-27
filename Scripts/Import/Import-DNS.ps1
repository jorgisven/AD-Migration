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
        [string]$Title,
        [switch]$AllowKeepBoth
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
    
    if ($AllowKeepBoth) {
        $btnKeepBoth = New-Object System.Windows.Forms.Button
        $btnKeepBoth.Text = "Keep Both"
        $btnKeepBoth.Location = New-Object System.Drawing.Point(300, 110)
        $btnKeepBoth.DialogResult = [System.Windows.Forms.DialogResult]::Ignore
        $form.Controls.Add($btnKeepBoth)
    }

    $form.AcceptButton = $btnOverwrite
    $form.CancelButton = $btnSkip
    $result = $form.ShowDialog()

    $action = if ($result -eq 'Yes') { 'Overwrite' } elseif ($result -eq 'Ignore') { 'KeepBoth' } else { 'Skip' }
    return @{ Action = $action; ApplyAll = $chkAll.Checked }
}

function Convert-IPv4ToUInt32 {
    param([string]$IPv4Address)

    $ipObj = $null
    if (-not [System.Net.IPAddress]::TryParse($IPv4Address, [ref]$ipObj)) { return $null }
    if ($ipObj.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { return $null }

    $bytes = $ipObj.GetAddressBytes()
    [array]::Reverse($bytes)
    return [BitConverter]::ToUInt32($bytes, 0)
}

function Convert-SubnetMaskToPrefixLength {
    param([string]$SubnetMask)

    $maskValue = Convert-IPv4ToUInt32 -IPv4Address $SubnetMask
    if ($null -eq $maskValue) { return $null }

    $bits = [Convert]::ToString([uint32]$maskValue, 2)
    return ($bits.ToCharArray() | Where-Object { $_ -eq '1' }).Count
}

function Test-IPv4InSubnet {
    param(
        [string]$IPv4Address,
        [string]$NetworkAddress,
        [int]$PrefixLength
    )

    if ($PrefixLength -lt 0 -or $PrefixLength -gt 32) { return $false }

    $ipValue = Convert-IPv4ToUInt32 -IPv4Address $IPv4Address
    $networkValue = Convert-IPv4ToUInt32 -IPv4Address $NetworkAddress
    if ($null -eq $ipValue -or $null -eq $networkValue) { return $false }

    $mask = if ($PrefixLength -eq 0) { [uint32]0 } else { ([uint32]::MaxValue -shl (32 - $PrefixLength)) }
    return (($ipValue -band $mask) -eq ($networkValue -band $mask))
}

function Get-LocalIPv4Subnets {
    $results = @()

    try {
        $netRows = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop | Where-Object {
            $_.IPAddress -notlike '127.*' -and
            $_.IPAddress -notlike '169.254.*' -and
            $_.PrefixLength -ge 1 -and
            $_.PrefixLength -le 32
        }

        foreach ($row in $netRows) {
            $results += [PSCustomObject]@{
                IPAddress    = [string]$row.IPAddress
                PrefixLength = [int]$row.PrefixLength
            }
        }
    } catch {
        $cfgRows = Get-CimInstance Win32_NetworkAdapterConfiguration -ErrorAction SilentlyContinue | Where-Object { $_.IPEnabled -eq $true }
        foreach ($cfg in $cfgRows) {
            $ips = @($cfg.IPAddress)
            $masks = @($cfg.IPSubnet)
            for ($i = 0; $i -lt $ips.Count; $i++) {
                $ip = [string]$ips[$i]
                if ($ip -notmatch '^\d+\.\d+\.\d+\.\d+$') { continue }
                if ($ip -like '127.*' -or $ip -like '169.254.*') { continue }

                $mask = if ($i -lt $masks.Count) { [string]$masks[$i] } else { '' }
                $prefix = Convert-SubnetMaskToPrefixLength -SubnetMask $mask
                if ($null -eq $prefix) { continue }

                $results += [PSCustomObject]@{
                    IPAddress    = $ip
                    PrefixLength = [int]$prefix
                }
            }
        }
    }

    return @($results | Group-Object { "$($_.IPAddress)/$($_.PrefixLength)" } | ForEach-Object { $_.Group[0] })
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

$dnsWarningMsg = "WARNING: You are about to import DNS zones and records into the target domain.`n`nIf the source and target domains operate on the same physical network or routable subnets, importing identical A/CNAME records could cause severe IP conflicts or routing issues.`n`nHave you verified your DNS transformation strategy and are you sure you want to proceed?"
$dnsWarningResult = [System.Windows.Forms.MessageBox]::Show($dnsWarningMsg, "DNS IP Conflict Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)

if ($dnsWarningResult -eq 'No') {
    Write-Log -Message "DNS Import cancelled by user at IP conflict warning." -Level WARN
    throw "DNS Import cancelled by user due to potential IP conflicts."
}

# Global conflict preferences (Prompt, Overwrite, Skip)
$GlobalRecordAction = "Prompt"
$GlobalIpConflictAction = "Prompt"

$script:DnsStats = [ordered]@{
    ZonesEvaluated   = 0
    ZonesCreated     = 0
    ZonesSkipped     = 0
    ZonesFailed      = 0
    RecordsEvaluated = 0
    RecordsCreated   = 0
    RecordsSkipped   = 0
    RecordsFailed    = 0
}

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

    # Safety preflight: compare incoming A-record network ranges with local host IPv4 subnet ranges.
    $incomingA = @($Records | Where-Object { $_.RecordType -eq 'A' -and ([string]$_.Data) -match '^\d+\.\d+\.\d+\.\d+$' })
    $incomingUniqueIps = @($incomingA.Data | Sort-Object -Unique)
    $localSubnets = @(Get-LocalIPv4Subnets)

    if ($incomingUniqueIps.Count -gt 0 -and $localSubnets.Count -gt 0) {
        $outsideSubnetCount = 0
        foreach ($ip in $incomingUniqueIps) {
            $inAnyLocalSubnet = $false
            foreach ($local in $localSubnets) {
                if (Test-IPv4InSubnet -IPv4Address $ip -NetworkAddress $local.IPAddress -PrefixLength $local.PrefixLength) {
                    $inAnyLocalSubnet = $true
                    break
                }
            }
            if (-not $inAnyLocalSubnet) { $outsideSubnetCount++ }
        }

        $outsidePct = [math]::Round((100.0 * $outsideSubnetCount / [math]::Max(1, $incomingUniqueIps.Count)), 1)
        $localSubnetText = ($localSubnets | ForEach-Object { "$($_.IPAddress)/$($_.PrefixLength)" } | Select-Object -First 5) -join ', '
        if ([string]::IsNullOrWhiteSpace($localSubnetText)) { $localSubnetText = 'No local IPv4 subnet data available' }

        $keepExactLikely = $false
        if ($SourcePath -eq $TransformPath) {
            $rawRecFile = Get-ChildItem -Path $ExportPath -Filter "DNS_Records_*.csv" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($rawRecFile) {
                try {
                    $rawRecords = Import-Csv $rawRecFile.FullName
                    $rawASet = @($rawRecords | Where-Object { $_.RecordType -eq 'A' -and ([string]$_.Data) -match '^\d+\.\d+\.\d+\.\d+$' } | Select-Object -ExpandProperty Data -Unique)
                    if ($rawASet.Count -gt 0) {
                        $overlap = @($incomingUniqueIps | Where-Object { $_ -in $rawASet }).Count
                        $overlapPct = 100.0 * $overlap / $rawASet.Count
                        if ($overlapPct -ge 90) { $keepExactLikely = $true }
                        Write-Log -Message "DNS import preflight: A-record overlap with raw export is $([math]::Round($overlapPct,1))% (heuristic for Keep Exact IPs)." -Level INFO
                    }
                } catch {
                    Write-Log -Message "DNS import preflight could not compare transformed and raw A-records: $($_.Exception.Message)" -Level WARN
                }
            }
        }

        if ($outsidePct -ge 70) {
            $riskMsg = "Potential subnet mismatch detected before DNS import.`n`nIncoming unique A-record IPs: $($incomingUniqueIps.Count)`nIPs outside local host subnet ranges: $outsideSubnetCount ($outsidePct%)`nLocal host subnet(s): $localSubnetText"
            if ($keepExactLikely) {
                $riskMsg += "`n`nThe transformed DNS file appears to keep most original source A-record IPs (high overlap with raw export)."
            }
            $riskMsg += "`n`nExample risk: importing 10.x records while this host operates in 192.168.x ranges can create routing/address conflicts depending on your topology.`n`nContinue with DNS import?"

            Write-Log -Message "DNS import preflight WARNING: subnet mismatch risk detected. OutsideLocalSubnet=$outsideSubnetCount/$($incomingUniqueIps.Count) ($outsidePct%). LocalSubnets='$localSubnetText'. KeepExactLikely=$keepExactLikely" -Level WARN
            $riskResult = [System.Windows.Forms.MessageBox]::Show($riskMsg, "DNS Subnet Mismatch Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($riskResult -eq 'No') {
                Write-Log -Message "DNS Import cancelled by user due to subnet mismatch preflight warning." -Level WARN
                throw "DNS Import cancelled by user after subnet mismatch warning."
            }
        } else {
            Write-Log -Message "DNS import preflight passed: $outsideSubnetCount/$($incomingUniqueIps.Count) unique A-record IPs are outside local host subnet ranges ($outsidePct%)." -Level INFO
        }
    } else {
        Write-Log -Message "DNS import preflight skipped subnet mismatch comparison (incoming A records or local IPv4 subnet data unavailable)." -Level INFO
    }

    # 2. Process Zones
    foreach ($z in $Zones) {
        $script:DnsStats.ZonesEvaluated++
        $zName = $z.ZoneName
        
        # Skip system zones if they weren't filtered earlier
        if ($zName -match "^_msdcs" -or $zName -eq "." -or $zName -eq "TrustAnchors" -or $zName -match "^(0|127|255)\.in-addr\.arpa$") {
            $script:DnsStats.ZonesSkipped++
            continue
        }

        try {
            $exists = Get-DnsServerZone -Name $zName -ComputerName $TargetServer -ErrorAction SilentlyContinue
            if ($exists) {
                $script:DnsStats.ZonesSkipped++
                Write-Log -Message "Zone '$zName' already exists on $TargetServer. Operating additively (skipping creation)." -Level INFO
            } else {
                # Create Zone
                if ($z.IsDsIntegrated -eq 'True') {
                    Add-DnsServerPrimaryZone -Name $zName -ReplicationScope "Domain" -ComputerName $TargetServer -ErrorAction Stop
                } else {
                    Add-DnsServerPrimaryZone -Name $zName -ZoneFile "$zName.dns" -ComputerName $TargetServer -ErrorAction Stop
                }
                $script:DnsStats.ZonesCreated++
                Write-Log -Message "Created Zone: $zName" -Level INFO
            }
        } catch {
            $script:DnsStats.ZonesFailed++
            Write-Log -Message "Failed to create zone '$zName': $_" -Level ERROR
        }
    }

    $zoneCache = @{}

    # 3. Process Records
    foreach ($r in $Records) {
        $script:DnsStats.RecordsEvaluated++
        $zName = $r.ZoneName
        $hostName = $r.HostName
        $type = $r.RecordType
        $data = $r.Data
        
        # Skip records for system/critical zones
        if ($zName -match "^_msdcs" -or $zName -eq "." -or $zName -eq "TrustAnchors" -or $zName -match "^(0|127|255)\.in-addr\.arpa$") {
            $script:DnsStats.RecordsSkipped++
            continue
        }

        # Skip SOA and NS as they are auto-generated
        if ($type -eq "SOA" -or $type -eq "NS") { 
            $script:DnsStats.RecordsSkipped++
            continue 
        }

        # Flag Zone Apex A/AAAA records in logs (Usually Domain Controllers or Root Web Servers)
        if ($hostName -eq "@" -and $type -in @('A', 'AAAA')) {
            Write-Log -Message "Notice: Importing Zone Apex (@) $type record pointing to IP '$data' in zone '$zName'. Ensure this IP is correct for the target environment." -Level WARN
        }

        # IP Conflict Check
        if (-not $zoneCache.ContainsKey($zName)) {
            $zoneCache[$zName] = Get-DnsServerResourceRecord -ZoneName $zName -ComputerName $TargetServer -ErrorAction SilentlyContinue
        }
        $existingTargetRecords = $zoneCache[$zName]

        $skipDueToIpConflict = $false
        if ($type -in @('A', 'AAAA')) {
            $ipConflicts = @($existingTargetRecords | Where-Object {
                $_.HostName -ne $hostName -and (
                    ($_.RecordType -eq 'A' -and $_.RecordData.IPv4Address.ToString() -eq $data) -or
                    ($_.RecordType -eq 'AAAA' -and $_.RecordData.IPv6Address.ToString() -eq $data)
                )
            })

            if ($ipConflicts.Count -gt 0) {
                $conflictNames = ($ipConflicts.HostName | Select-Object -Unique) -join ', '
                $action = "Skip"
                
                if ($GlobalIpConflictAction -eq "Prompt") {
                    $res = Show-ConflictDialog -Message "IP Conflict in zone '$zName'.`nSource record '$hostName' uses IP '$data', but this IP is already used by target record(s): $conflictNames.`n`nKeep Both will allow multiple names to share this IP." -Title "IP Conflict Detected" -AllowKeepBoth
                    $action = $res.Action
                    if ($res.ApplyAll) { $GlobalIpConflictAction = $action }
                } else {
                    $action = $GlobalIpConflictAction
                }

                if ($action -eq "Overwrite") {
                    Write-Log -Message "Overwriting IP Conflict: Deleting existing records using IP '$data' ($conflictNames) in $zName" -Level INFO
                    try {
                        foreach ($conflict in $ipConflicts) {
                            Get-DnsServerResourceRecord -ZoneName $zName -Name $conflict.HostName -RecordType $conflict.RecordType -ComputerName $TargetServer -ErrorAction SilentlyContinue |
                                Remove-DnsServerResourceRecord -ZoneName $zName -ComputerName $TargetServer -Force -Confirm:$false -ErrorAction Stop
                        }
                        # Update cache to remove them
                        $zoneCache[$zName] = $zoneCache[$zName] | Where-Object { $_.HostName -notin $ipConflicts.HostName }
                    } catch {
                        $script:DnsStats.RecordsFailed++
                        Write-Log -Message "Failed to remove conflicting IP record: $_" -Level ERROR
                        $skipDueToIpConflict = $true
                    }
                } elseif ($action -eq "KeepBoth") {
                    Write-Log -Message "Keeping both records for IP '$data' (Existing: $conflictNames, New: $hostName)." -Level INFO
                } else {
                    $script:DnsStats.RecordsSkipped++
                    Write-Log -Message "Skipping record '$hostName' due to IP conflict with '$conflictNames' on IP '$data'." -Level INFO
                    $skipDueToIpConflict = $true
                }
            }
        }

        if ($skipDueToIpConflict) {
            continue
        }

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
                $script:DnsStats.RecordsCreated++
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
                            $script:DnsStats.RecordsFailed++
                            Write-Log -Message "Failed to remove conflicting record '$hostName': $_" -Level ERROR
                        }
                    } else {
                        $script:DnsStats.RecordsSkipped++
                        Write-Log -Message "Skipping conflicting record: $hostName ($type)" -Level INFO
                    }
                } else {
                    $script:DnsStats.RecordsFailed++
                    Write-Log -Message "Failed to add $type record '$hostName' in '$zName': $errMsg" -Level WARN
                }
            }
        }
    }
    
    Write-Host "DNS Import Complete."

} -Operation "Import DNS Records"

$warningCount = $script:DnsStats.ZonesFailed + $script:DnsStats.RecordsFailed
$summary = "DNS Import summary: Zones (Eval=$($script:DnsStats.ZonesEvaluated), Created=$($script:DnsStats.ZonesCreated), Skipped=$($script:DnsStats.ZonesSkipped), Failed=$($script:DnsStats.ZonesFailed)), Records (Eval=$($script:DnsStats.RecordsEvaluated), Created=$($script:DnsStats.RecordsCreated), Skipped=$($script:DnsStats.RecordsSkipped), Failed=$($script:DnsStats.RecordsFailed))"

if ($warningCount -gt 0) {
    Write-Host "[!] WARNING: DNS Import encountered $warningCount failure(s). See logs for details." -ForegroundColor Yellow
    Write-Log -Message "Import DNS Records completed with warnings. $summary" -Level WARN
} else {
    Write-Log -Message "Import DNS Records completed. $summary" -Level INFO
}