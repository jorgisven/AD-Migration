<#
.SYNOPSIS
    Clean and transform DNS records for migration.

.DESCRIPTION
    Filters stale/infrastructure DNS records (SOA, NS, _msdcs) and rewrites FQDNs in record data
    (e.g., updating CNAME targets) and maps IP addresses to new subnets.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$OldSuffix,

    [Parameter(Mandatory = $false)]
    [string]$NewSuffix,

    [Parameter(Mandatory = $false)]
    [string]$OldIPPrefix,

    [Parameter(Mandatory = $false)]
    [string]$NewIPPrefix,

    [Parameter(Mandatory = $false)]
    [int]$MaxRecordAgeDays = 0 # 0 = Keep all
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

$config = Get-ADMigrationConfig
$SourcePath = Join-Path $config.ExportRoot 'DNS'
$DestPath   = Join-Path $config.TransformRoot 'DNS'

# GUI Prompt for IP Mapping Strategy
if (-not $OldIPPrefix -and -not $NewIPPrefix) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName Microsoft.VisualBasic

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "DNS IP Transformation Strategy"
    $form.Size = New-Object System.Drawing.Size(400,350)
    $form.StartPosition = "CenterScreen"

    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Location = New-Object System.Drawing.Point(10, 10)
    $gb.Size = New-Object System.Drawing.Size(360, 220)
    $gb.Text = "Select IP Mapping Strategy"
    $form.Controls.Add($gb)

    $r1 = New-Object System.Windows.Forms.RadioButton; $r1.Text = "Keep Exact IPs (/32)"; $r1.Location = New-Object System.Drawing.Point(10, 20); $r1.AutoSize = $true; $r1.Checked = $true; $gb.Controls.Add($r1)
    $r2 = New-Object System.Windows.Forms.RadioButton; $r2.Text = "Only last octet varies (/24)"; $r2.Location = New-Object System.Drawing.Point(10, 50); $r2.AutoSize = $true; $gb.Controls.Add($r2)
    $r3 = New-Object System.Windows.Forms.RadioButton; $r3.Text = "Last two octets vary (/16)"; $r3.Location = New-Object System.Drawing.Point(10, 80); $r3.AutoSize = $true; $gb.Controls.Add($r3)
    $r4 = New-Object System.Windows.Forms.RadioButton; $r4.Text = "Only last 3 octets vary (/8)"; $r4.Location = New-Object System.Drawing.Point(10, 110); $r4.AutoSize = $true; $gb.Controls.Add($r4)
    $r5 = New-Object System.Windows.Forms.RadioButton; $r5.Text = "First two octets vary (*.*.255.255)"; $r5.Location = New-Object System.Drawing.Point(10, 140); $r5.AutoSize = $true; $gb.Controls.Add($r5)
    $r6 = New-Object System.Windows.Forms.RadioButton; $r6.Text = "Full Custom"; $r6.Location = New-Object System.Drawing.Point(10, 170); $r6.AutoSize = $true; $gb.Controls.Add($r6)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(280, 250)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)
    $form.AcceptButton = $btnOk

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($r2.Checked) { # /24
            $OldIPPrefix = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Old Network Prefix (e.g. 192.168.1.):", "IP Map /24", "")
            $NewIPPrefix = [Microsoft.VisualBasic.Interaction]::InputBox("Enter New Network Prefix (e.g. 10.0.0.):", "IP Map /24", "")
        }
        elseif ($r3.Checked -or $r5.Checked) { # /16
            $OldIPPrefix = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Old Network Prefix (e.g. 192.168.):", "IP Map /16", "")
            $NewIPPrefix = [Microsoft.VisualBasic.Interaction]::InputBox("Enter New Network Prefix (e.g. 10.0.):", "IP Map /16", "")
        }
        elseif ($r4.Checked) { # /8
            $OldIPPrefix = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Old Network Prefix (e.g. 10.):", "IP Map /8", "")
            $NewIPPrefix = [Microsoft.VisualBasic.Interaction]::InputBox("Enter New Network Prefix (e.g. 192.):", "IP Map /8", "")
        }
        elseif ($r6.Checked) { # Custom
            $msgResult = [System.Windows.Forms.MessageBox]::Show("Do you want to generate a CSV template to fill out manually?`n`nYes = Generate Template`nNo = Select Existing CSV File", "Custom DNS Import", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
            
            if ($msgResult -eq 'Yes') {
                $templateFile = Join-Path $DestPath "DNS_Import_Template.csv"
                $templateData = [PSCustomObject]@{
                    ZoneName   = "example.com"
                    HostName   = "www"
                    RecordType = "A"
                    Data       = "192.168.1.10"
                    Timestamp  = ""
                    TTL        = "01:00:00"
                }
                $templateData | Export-Csv -Path $templateFile -NoTypeInformation
                [System.Windows.Forms.MessageBox]::Show("Template generated at:`n$templateFile`n`nPlease fill it out with your hosts/IPs and run this script again (Select 'Full Custom' -> 'No' to load it).", "Template Generated", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                exit
            }
            elseif ($msgResult -eq 'No') {
                $ofd = New-Object System.Windows.Forms.OpenFileDialog
                $ofd.Title = "Select Custom DNS Records CSV"
                $ofd.Filter = "CSV Files (*.csv)|*.csv"
                if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $CustomFile = $ofd.FileName
                } else {
                    exit
                }
            } else {
                exit
            }
        }
    }
}

if (-not (Test-Path $DestPath)) { New-Item -ItemType Directory -Path $DestPath -Force | Out-Null }

Invoke-Safely -ScriptBlock {
    # 1. Find latest Export files
    $zoneFile = Get-ChildItem -Path $SourcePath -Filter "DNS_Zones_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $recFile  = Get-ChildItem -Path $SourcePath -Filter "DNS_Records_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if (-not $zoneFile -or -not $recFile) { throw "Missing DNS export files in $SourcePath" }

    $Records = Import-Csv $recFile.FullName
    $Zones   = Import-Csv $zoneFile.FullName

    # Filter out system zones (e.g. _msdcs, TrustAnchors) to prevent importing infrastructure garbage
    $Zones = $Zones | Where-Object { 
        $_.ZoneName -notmatch "^_msdcs" -and 
        $_.ZoneName -ne "TrustAnchors" -and 
        $_.ZoneName -ne "." 
    }
    $ValidZoneNames = @($Zones.ZoneName)

    Write-Log -Message "Processing $($Records.Count) records..." -Level INFO

    $TransformedRecords = @()
    $SkippedCount = 0
    $ModifiedCount = 0

    if ($CustomFile) {
        Write-Log -Message "Using custom record file: $CustomFile" -Level INFO
        $TransformedRecords = Import-Csv $CustomFile
    } else {
        $CutoffDate = (Get-Date).AddDays(-$MaxRecordAgeDays)

        foreach ($rec in $Records) {
            # 0. Infrastructure Filter
            # Skip records for zones we filtered out
            if ($rec.ZoneName -notin $ValidZoneNames) {
                $SkippedCount++
                continue
            }
            # Skip SOA and NS records (these are regenerated by the target DC)
            if ($rec.RecordType -in @('SOA', 'NS')) {
                $SkippedCount++
                continue
            }

            # 1. Stale Record Filter
            if ($MaxRecordAgeDays -gt 0 -and $rec.Timestamp) {
                try {
                    $recDate = [DateTime]::Parse($rec.Timestamp)
                    if ($recDate -lt $CutoffDate) {
                        $SkippedCount++
                        continue
                    }
                } catch {
                    # If timestamp parsing fails, keep the record to be safe
                }
            }

            # 2. FQDN Rewrite (CNAME, MX, TXT data)
            if ($OldSuffix -and $NewSuffix -and $rec.Data -like "*$OldSuffix*") {
                $rec.Data = $rec.Data -replace [regex]::Escape($OldSuffix), $NewSuffix
                $ModifiedCount++
            }

            # 3. IP Address Rewrite (A Records)
            if ($OldIPPrefix -and $NewIPPrefix -and $rec.RecordType -eq 'A' -and $rec.Data -like "$OldIPPrefix*") {
                $rec.Data = $rec.Data -replace [regex]::Escape($OldIPPrefix), $NewIPPrefix
                $ModifiedCount++
            }

            $TransformedRecords += $rec
        }
    }

    # Export Transformed Data
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    
    # Copy Zones (pass-through, but saved to Transform for consistency)
    $newZoneFile = Join-Path $DestPath "DNS_Zones_Transformed_$timestamp.csv"
    $Zones | Export-Csv -Path $newZoneFile -NoTypeInformation -Encoding UTF8

    # Save Records
    $newRecFile = Join-Path $DestPath "DNS_Records_Transformed_$timestamp.csv"
    $TransformedRecords | Export-Csv -Path $newRecFile -NoTypeInformation -Encoding UTF8

    Write-Log -Message "Transformation Complete." -Level INFO
    Write-Log -Message "  - Input: $($Records.Count)" -Level INFO
    Write-Log -Message "  - Removed (Stale): $SkippedCount" -Level INFO
    Write-Log -Message "  - Modified (Rewrite): $ModifiedCount" -Level INFO
    Write-Log -Message "  - Output: $($TransformedRecords.Count)" -Level INFO
    
    Write-Host "DNS Transform Complete."
    Write-Host "  Removed: $SkippedCount stale records"
    Write-Host "  Updated: $ModifiedCount records (Suffix/IP rewrite)"
    Write-Host "  Saved to: $DestPath"

} -Operation "Transform DNS Records"
