<#
.SYNOPSIS
    Export GPO XML reports from source domain.

.DESCRIPTION
    Generates HTML and XML reports for all GPOs, including links, WMI filters, and settings.
    Reports are useful for analyzing policies before migration.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$SourceDomain
)

# Add GUI support
Add-Type -AssemblyName System.Windows.Forms

# Add P/Invoke to bring window to foreground to ensure prompts are visible
Add-Type -MemberDefinition @"
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
"@ -Name Win32SetForegroundWindow -Namespace Win32Functions -PassThru | Out-Null


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
$ExportPath = Join-Path $config.ExportRoot 'GPO_Reports'
$BackupPath = Join-Path $config.ExportRoot 'GPO_Backups'
$OUExportPath = Join-Path $config.ExportRoot 'OU_Structure'

# Ensure export directory exists
if (-not (Test-Path $ExportPath)) {
    New-Item -ItemType Directory -Path $ExportPath -Force | Out-Null
}
if (-not (Test-Path $BackupPath)) {
    New-Item -ItemType Directory -Path $BackupPath -Force | Out-Null
}

# Prompt for source domain if not provided
while ([string]::IsNullOrWhiteSpace($SourceDomain)) {
    $SourceDomain = Read-Host "Enter source domain name (e.g., source.local)"
}

# Load Empty OUs list if available
$EmptyOUs = @()
$emptyListFile = Join-Path $OUExportPath "EmptyOUs.txt"
if (Test-Path $emptyListFile) {
    $EmptyOUs = Get-Content $emptyListFile
    Write-Log -Message "Loaded $($EmptyOUs.Count) empty OUs for link analysis" -Level INFO
}

Write-Log -Message "Starting GPO report export from domain: $SourceDomain" -Level INFO

function Get-FirstXmlInnerText {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$Document,

        [Parameter(Mandatory = $true)]
        [string[]]$LocalNames
    )

    foreach ($localName in $LocalNames) {
        $nodes = $Document.SelectNodes("//*[local-name()='$localName']")
        foreach ($node in @($nodes)) {
            $value = ([string]$node.InnerText).Trim()
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
    }

    return ''
}

function Get-GpoManifestEntriesFromBackupFolders {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $entries = @()
    $folders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match '^\{?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\}?$'
    }

    foreach ($folder in $folders) {
        $descriptor = @(
            (Join-Path $folder.FullName 'Backup.xml'),
            (Join-Path $folder.FullName 'backup.xml'),
            (Join-Path $folder.FullName 'bkupInfo.xml')
        ) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

        if (-not $descriptor) { continue }

        try {
            [xml]$xml = Get-Content -LiteralPath $descriptor -ErrorAction Stop
            $displayName = Get-FirstXmlInnerText -Document $xml -LocalNames @('GPODisplayName', 'Name')
            if ([string]::IsNullOrWhiteSpace($displayName)) { continue }

            $idValue = Get-FirstXmlInnerText -Document $xml -LocalNames @('BackupId', 'ID')
            if ([string]::IsNullOrWhiteSpace($idValue)) { $idValue = $folder.Name }
            $normalizedId = $idValue.Trim().TrimStart('{').TrimEnd('}')

            if ($normalizedId -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') { continue }

            $entries += [PSCustomObject]@{
                ID             = "{$normalizedId}"
                GPODisplayName = $displayName
                BackupTime     = $folder.LastWriteTime.ToString('o')
            }
        } catch {
            Write-Log -Message "Failed reading backup descriptor '$descriptor' while building manifest fallback: $($_.Exception.Message)" -Level WARN
        }
    }

    return @($entries | Group-Object ID | ForEach-Object { $_.Group | Select-Object -First 1 })
}

function Write-GpoBackupManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [object[]]$Entries
    )

    $manifestPath = Join-Path $Path 'manifest.xml'
    $xmlSettings = New-Object System.Xml.XmlWriterSettings
    $xmlSettings.Indent = $true
    $xmlSettings.Encoding = [System.Text.UTF8Encoding]::new($false)

    $writer = [System.Xml.XmlWriter]::Create($manifestPath, $xmlSettings)
    try {
        $writer.WriteStartDocument()
        $writer.WriteStartElement('GPOBackupManifest')
        $writer.WriteAttributeString('schemaVersion', '1.0')

        foreach ($entry in $Entries) {
            $writer.WriteStartElement('BackupInst')
            $writer.WriteElementString('ID', [string]$entry.ID)
            $writer.WriteElementString('GPODisplayName', [string]$entry.GPODisplayName)
            $writer.WriteElementString('BackupTime', [string]$entry.BackupTime)
            $writer.WriteEndElement()
        }

        $writer.WriteEndElement()
        $writer.WriteEndDocument()
    } finally {
        $writer.Dispose()
    }

    return $manifestPath
}

try {
    Invoke-Safely -ScriptBlock {
        # Get all GPOs from source domain
        $GPOs = Get-GPO -All -Domain $SourceDomain
        
        Write-Log -Message "Found $($GPOs.Count) GPOs to export" -Level INFO
        
        $summaryFile = Join-Path $ExportPath "_GPO_Summary_${SourceDomain}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $backupStatusFile = Join-Path $BackupPath "_GPO_Backup_Status_${SourceDomain}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $EmptyLinkedGPOs = @()
        $BackupStatus = @()
        
        # Generate reports for each GPO
        $ReportSummary = foreach ($GPO in $GPOs) {
            try {
                $safeName = $GPO.DisplayName -replace '[\\/:"*?<>|]', '_'

                # Generate HTML report
                $htmlPath = Join-Path $ExportPath "${safeName}.html"
                Get-GPOReport -Guid $GPO.Id -ReportType Html -Path $htmlPath -Server $SourceDomain
                
                # Generate XML report (Required for Transform phase)
                $xmlPath = Join-Path $ExportPath "${safeName}.xml"
                Get-GPOReport -Guid $GPO.Id -ReportType Xml -Path $xmlPath -Server $SourceDomain

                # Perform actual GPO Backup (Required for Import)
                $backupSucceeded = $true
                $backupError = ''
                $backupErrorId = ''
                $backupErrorCategory = ''
                try {
                    Backup-GPO -Guid $GPO.Id -Path $BackupPath -Server $SourceDomain -ErrorAction Stop | Out-Null
                } catch {
                    $backupSucceeded = $false
                    $backupError = $_.Exception.Message
                    $backupErrorId = $_.FullyQualifiedErrorId
                    $backupErrorCategory = $_.CategoryInfo.Category
                    Write-Log -Message "Backup failed for GPO '$($GPO.DisplayName)' (GUID: $($GPO.Id)) on domain '$SourceDomain'. Reason: $backupError | Category: $backupErrorCategory | ErrorId: $backupErrorId" -Level WARN
                }

                $BackupStatus += [PSCustomObject]@{
                    GPOName      = $GPO.DisplayName
                    GPOID        = $GPO.Id
                    BackupStatus = if ($backupSucceeded) { 'Succeeded' } else { 'Failed' }
                    Error        = $backupError
                    ErrorCategory = $backupErrorCategory
                    ErrorId      = $backupErrorId
                }

                # Check if GPO is linked ONLY to empty OUs
                [xml]$xmlData = Get-Content $xmlPath
                $links = $xmlData.GPO.LinksTo
                $isLinkedOnlyToEmpty = $false
                
                if ($links -and $EmptyOUs.Count -gt 0) {
                    $allLinksEmpty = $true
                    foreach ($link in $links) {
                        if ($link.SOMPath -notin $EmptyOUs) { $allLinksEmpty = $false; break }
                    }
                    if ($allLinksEmpty) { 
                        $isLinkedOnlyToEmpty = $true 
                        $EmptyLinkedGPOs += $GPO.DisplayName
                    }
                }

                # Output the object to be collected by the $ReportSummary variable
                [PSCustomObject]@{
                    GPOName            = $GPO.DisplayName
                    GPOID              = $GPO.Id
                    CreationTime       = $GPO.CreationTime
                    ModificationTime   = $GPO.ModificationTime
                    Owner              = $GPO.Owner
                    HTMLReport         = (Split-Path $htmlPath -Leaf)
                    XMLReport          = (Split-Path $xmlPath -Leaf)
                    IsLinkedOnlyToEmpty = $isLinkedOnlyToEmpty
                }
            } catch {
                Write-Log -Message "Failed to process GPO '$($GPO.DisplayName)'. Error: $_" -Level WARN
                $BackupStatus += [PSCustomObject]@{
                    GPOName      = $GPO.DisplayName
                    GPOID        = $GPO.Id
                    BackupStatus = 'Failed'
                    Error        = $_.Exception.Message
                }
            }
        }
        
        Write-Log -Message "Generated summary data for $($ReportSummary.Count) GPOs before filtering." -Level INFO

        # Handle Empty-Linked GPOs
        if ($EmptyLinkedGPOs.Count -gt 0) {
            # Bring the console window to the front to make sure the user sees this prompt
            try {
                $myWindowHandle = (Get-Process -Id $PID).MainWindowHandle
                [Win32Functions.Win32SetForegroundWindow]::SetForegroundWindow($myWindowHandle) | Out-Null
            } catch {}

            $msg = "Found $($EmptyLinkedGPOs.Count) GPOs that are linked ONLY to empty OUs.`n`nDo you want to EXPORT these empty-linked GPOs?`n`nClick YES to Export them.`nClick NO to Skip (delete) them."
            $result = [System.Windows.Forms.MessageBox]::Show($msg, "Empty-Linked GPOs Detected", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            
            if ($result -eq 'No') {
                Write-Log -Message "User selected SKIP for empty-linked GPOs. Cleaning up..." -Level INFO
                # Filter summary
                $ReportSummary = $ReportSummary | Where-Object { -not $_.IsLinkedOnlyToEmpty }
                Write-Log -Message "$($ReportSummary.Count) GPOs remain in summary after filtering." -Level INFO
                
                # Delete files for skipped GPOs
                foreach ($gpoName in $EmptyLinkedGPOs) {
                    $safeName = $gpoName -replace '[\\/:"*?<>|]', '_'
                    Remove-Item (Join-Path $ExportPath "${safeName}.html") -ErrorAction SilentlyContinue
                    Remove-Item (Join-Path $ExportPath "${safeName}.xml") -ErrorAction SilentlyContinue
                    # Note: We don't delete the Backup folder content easily as it uses GUIDs, but the XML/HTML reports are gone.
                }
            }
        }

        # Final validation check
        if ($GPOs.Count -gt 0 -and $ReportSummary.Count -eq 0) {
            $finalWarnMsg = "WARNING: Found $($GPOs.Count) GPOs but 0 were exported to the summary file. This may be because all GPOs were linked to empty OUs and you chose to skip them, or an error occurred. Please check the logs."
            Write-Log -Message $finalWarnMsg -Level WARN
            Write-Host $finalWarnMsg -ForegroundColor Yellow
        }

        if ($ReportSummary) {
            $ReportSummary | Export-Csv -Path $summaryFile -NoTypeInformation -Encoding UTF8
        }

        if ($BackupStatus.Count -gt 0) {
            $BackupStatus | Export-Csv -Path $backupStatusFile -NoTypeInformation -Encoding UTF8
            Write-Log -Message "Exported GPO backup status report to $backupStatusFile" -Level INFO
        }

        $manifestPath = Join-Path $BackupPath 'manifest.xml'
        $manifestAltPath = Join-Path $BackupPath 'Manifest.xml'
        $manifestEntryCount = 0
        $manifestUsable = $false

        $existingManifestPath = ''
        if (Test-Path -LiteralPath $manifestPath) {
            $existingManifestPath = $manifestPath
        } elseif (Test-Path -LiteralPath $manifestAltPath) {
            $existingManifestPath = $manifestAltPath
        }

        if (-not [string]::IsNullOrWhiteSpace($existingManifestPath)) {
            try {
                [xml]$manifestXml = Get-Content -LiteralPath $existingManifestPath -ErrorAction Stop
                $manifestNodes = @($manifestXml.SelectNodes("//*[local-name()='BackupInst']"))
                $manifestEntryCount = $manifestNodes.Count
                $manifestUsable = ($manifestEntryCount -gt 0)
                if ($manifestUsable) {
                    Write-Log -Message "GPO backup manifest detected at '$existingManifestPath' with $manifestEntryCount entries." -Level INFO
                    Write-Host "GPO backup manifest ready: $manifestEntryCount entries." -ForegroundColor Green
                } else {
                    Write-Log -Message "GPO backup manifest detected at '$existingManifestPath' but contains no BackupInst entries. Rebuilding fallback manifest." -Level WARN
                }
            } catch {
                Write-Log -Message "Failed parsing existing GPO backup manifest at '$existingManifestPath': $($_.Exception.Message). Rebuilding fallback manifest." -Level WARN
                $manifestUsable = $false
            }
        }

        if (-not $manifestUsable) {
            $manifestEntries = @(Get-GpoManifestEntriesFromBackupFolders -Path $BackupPath)
            if ($manifestEntries.Count -gt 0) {
                $writtenManifest = Write-GpoBackupManifest -Path $BackupPath -Entries $manifestEntries
                Write-Log -Message "Generated fallback GPO backup manifest at '$writtenManifest' with $($manifestEntries.Count) entries." -Level WARN
                Write-Host "Generated fallback manifest.xml with $($manifestEntries.Count) GPO backup entries." -ForegroundColor Yellow
            } else {
                Write-Log -Message "Could not generate fallback GPO backup manifest because no valid backup descriptors were discovered under '$BackupPath'." -Level WARN
                Write-Host "[!] WARNING: Unable to build manifest.xml from backup folders. Check GPO_Backups contents." -ForegroundColor Yellow
            }
        }

        $failedBackups = @($BackupStatus | Where-Object { $_.BackupStatus -eq 'Failed' })
        if ($failedBackups.Count -gt 0) {
            Write-Host "[!] WARNING: $($failedBackups.Count) GPO backup(s) failed. See $backupStatusFile for details." -ForegroundColor Yellow
            Write-Log -Message "Detected $($failedBackups.Count) failed GPO backup(s). See $backupStatusFile for details." -Level WARN
        }
        
        Write-Log -Message "Exported $($ReportSummary.Count) GPO reports to $ExportPath" -Level INFO
        Write-Host "GPO report export complete: $($ReportSummary.Count) reports generated"
        
    } -Operation "Export GPO reports from $SourceDomain"
    
} catch {
    Write-Log -Message "Failed to export GPO reports: $_" -Level ERROR
    throw "GPO report export failed. Check logs for details."
}
