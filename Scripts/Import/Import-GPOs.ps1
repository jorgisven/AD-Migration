<#
.SYNOPSIS
    Import Group Policy Objects (GPOs) from backups into the target domain.

.DESCRIPTION
    Reads GPO backups from the Export directory and imports them into the current (Target) domain.
    Supports using a Migration Table to map security principals and UNC paths.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain,

    [Parameter(Mandatory = $false)]
    [string]$MigrationTablePath,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeDefaultDomainPolicies,

    [Parameter(Mandatory = $false)]
    [string]$DefaultPolicyNameSuffix = ' (Migrated Copy)',

    [Parameter(Mandatory = $false)]
    [ValidateSet('Prompt', 'Skip', 'Rename')]
    [string]$DefaultDomainPolicyMode = 'Prompt',

    [Parameter(Mandatory = $false)]
    [switch]$SkipExistingOnly
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

# Check for GroupPolicy module
if (-not (Get-Module -Name GroupPolicy -ListAvailable)) {
    throw "GroupPolicy module (GPMC) is required but not installed."
}

$config = Get-ADMigrationConfig
$BackupPath = Join-Path $config.ExportRoot 'GPO_Backups'

# Default migration table path if not provided
if (-not $MigrationTablePath) {
    # Look for a .migtable file in the Transform/Mapping folder
    $MapPath = Join-Path $config.TransformRoot 'Mapping'
    $foundTable = Get-ChildItem -Path $MapPath -Filter "*.migtable" | Select-Object -First 1
    if ($foundTable) {
        $MigrationTablePath = $foundTable.FullName
        Write-Log -Message "Auto-detected migration table: $MigrationTablePath" -Level INFO
    }
}

# Validate Migration Table if present
if ($MigrationTablePath) {
    if (-not (Test-Path $MigrationTablePath)) {
        throw "Migration Table file not found at '$MigrationTablePath'"
    }
    
    # Check for unmapped entries which cause Import-GPO to fail
    try {
        [xml]$migTableXml = Get-Content $MigrationTablePath
        $unmapped = $migTableXml.MigrationTable.Mapping | Where-Object { 
            ($_.Destination.Path -eq "" -or $_.Destination.Path -eq $null) 
        }
        if ($unmapped) {
            throw "The Migration Table contains $($unmapped.Count) unmapped entries (empty Destination). Please review and edit '$MigrationTablePath' before importing."
        }
        Write-Log -Message "Using validated Migration Table: $MigrationTablePath" -Level INFO
    } catch {
        throw "Migration Table Validation Failed: $_"
    }
} else {
    Write-Log -Message "No Migration Table specified or detected. GPOs will be imported with original security principals and paths." -Level WARN
}

if (-not (Test-Path $BackupPath)) {
    Write-Log -Message "GPO Backup directory not found at $BackupPath. Ensure Export-GPOs (or manual backup) was run." -Level ERROR
    throw "Missing GPO Backups"
}

if (-not $TargetDomain) { $TargetDomain = (Get-ADDomain).DNSRoot }

$defaultPolicyNames = @(
    'Default Domain Policy',
    'Default Domain Controllers Policy'
)

$effectiveDefaultPolicyMode = $DefaultDomainPolicyMode
if ($effectiveDefaultPolicyMode -eq 'Prompt' -and $IncludeDefaultDomainPolicies) {
    # Backward compatibility: old switch implied importing defaults.
    $effectiveDefaultPolicyMode = 'Rename'
    Write-Log -Message "Legacy switch IncludeDefaultDomainPolicies was set. Effective default policy mode changed from Prompt to Rename." -Level WARN
}

if ($effectiveDefaultPolicyMode -eq 'Prompt') {
    if ([Environment]::UserInteractive) {
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            $prompt = "Default domain policies were found in backups.`n`nYES = Skip them (recommended)`nNO = Import as renamed copies`nCancel = Abort"
            $choice = [System.Windows.Forms.MessageBox]::Show($prompt, "Default GPO Policy Handling", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
                $effectiveDefaultPolicyMode = 'Skip'
                Write-Log -Message "Default GPO policy prompt selection: Skip" -Level INFO
            } elseif ($choice -eq [System.Windows.Forms.DialogResult]::No) {
                $effectiveDefaultPolicyMode = 'Rename'
                Write-Log -Message "Default GPO policy prompt selection: Rename" -Level INFO
            } else {
                throw "Import cancelled by user while selecting default policy handling mode."
            }
        } catch {
            Write-Log -Message "Default policy mode prompt unavailable or cancelled. Falling back to Skip. Details: $($_.Exception.Message)" -Level WARN
            $effectiveDefaultPolicyMode = 'Skip'
        }
    } else {
        Write-Log -Message "Non-interactive session detected and DefaultDomainPolicyMode='Prompt'. Falling back to Skip." -Level WARN
        $effectiveDefaultPolicyMode = 'Skip'
    }
}

if ($effectiveDefaultPolicyMode -eq 'Rename' -and [string]::IsNullOrWhiteSpace($DefaultPolicyNameSuffix)) {
    $DefaultPolicyNameSuffix = ' (Migrated Copy)'
}

Write-Log -Message "Default domain policy handling mode: $effectiveDefaultPolicyMode" -Level INFO
Write-Log -Message "Default domain policy options: Mode='$effectiveDefaultPolicyMode', Suffix='$DefaultPolicyNameSuffix'" -Level INFO
Write-Log -Message "Import-GPOs script version: 2026-03-27.8" -Level INFO

# Diagnostics: Check GroupPolicy module availability
$gpModule = Get-Module -Name GroupPolicy -ListAvailable -ErrorAction SilentlyContinue
if ($gpModule) {
    Write-Log -Message "GroupPolicy module found: Version $($gpModule.Version)" -Level INFO
} else {
    Write-Log -Message "GroupPolicy module (GPMC/RSAT) not found in ListAvailable. Import-GPO may still work but Get-GPOBackup cmdlet will be unavailable. Proceeding with folder-based discovery." -Level WARN
}

$gpoBackupCmdlet = Get-Command Get-GPOBackup -ErrorAction SilentlyContinue
if ($gpoBackupCmdlet) {
    Write-Log -Message "Get-GPOBackup cmdlet available." -Level INFO
} else {
    Write-Log -Message "Get-GPOBackup cmdlet not available. This usually means GroupPolicy module is not properly installed or loaded. The script will use folder-based discovery instead, but Import-GPO cmdlet must still be available for imports to work." -Level WARN
}

Write-Log -Message "Starting GPO Import to $TargetDomain..." -Level INFO

function ConvertTo-GpoBackupGuid {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }

    return $Value.Trim().TrimStart('{').TrimEnd('}').ToLowerInvariant()
}

function Test-GpoBackupGuid {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value
    )

    return (ConvertTo-GpoBackupGuid -Value $Value) -match '^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$'
}

function Get-GpoBackupFolderMap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $folders = @{}
    $backupFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue

    foreach ($folder in $backupFolders) {
        $normalizedId = ConvertTo-GpoBackupGuid -Value $folder.Name
        if (Test-GpoBackupGuid -Value $normalizedId) {
            $folders[$normalizedId] = $folder
        }
    }

    return $folders
}

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

function Get-GpoDisplayNameFromXmlFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$XmlFile
    )

    $rawContent = Get-Content -LiteralPath $XmlFile.FullName -Raw -ErrorAction Stop

    try {
        [xml]$document = $rawContent
    } catch {
        $document = $null
    }

    if ($document) {
        $displayName = Get-FirstXmlInnerText -Document $document -LocalNames @('GPODisplayName', 'Name')
        if (-not [string]::IsNullOrWhiteSpace($displayName)) {
            return $displayName
        }
    }

    $patterns = @(
        '<GPODisplayName>\s*(?<Value>[^<]+?)\s*</GPODisplayName>',
        '<Name>\s*(?<Value>[^<]+?)\s*</Name>'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($rawContent, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($match.Success) {
            $value = $match.Groups['Value'].Value.Trim()
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
    }

    return ''
}

function Get-GpoGuidCandidatesFromXmlFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$XmlFile
    )

    $rawContent = Get-Content -LiteralPath $XmlFile.FullName -Raw -ErrorAction Stop
    [xml]$document = $null
    try {
        [xml]$document = $rawContent
    } catch {
        $document = $null
    }

    $results = @()

    if ($document) {
        # Restrict extraction to backup instance identity fields only.
        $backupIdNodes = @($document.SelectNodes("//*[local-name()='BackupId']"))
        foreach ($node in $backupIdNodes) {
            $normalized = ConvertTo-GpoBackupGuid -Value ([string]$node.InnerText)
            if (Test-GpoBackupGuid -Value $normalized) {
                $results += $normalized
            }
        }

        $backupInstIdNodes = @($document.SelectNodes("//*[local-name()='BackupInst']/*[local-name()='ID']"))
        foreach ($node in $backupInstIdNodes) {
            $normalized = ConvertTo-GpoBackupGuid -Value ([string]$node.InnerText)
            if (Test-GpoBackupGuid -Value $normalized) {
                $results += $normalized
            }
        }
    }

    if ($results.Count -eq 0) {
        $tagPatterns = @(
            '<BackupId>\s*(?<Value>\{?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\}?)\s*</BackupId>',
            '<BackupInst[^>]*>[\s\S]*?<ID>\s*(?<Value>\{?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\}?)\s*</ID>'
        )

        foreach ($pattern in $tagPatterns) {
            $regexMatches = [regex]::Matches($rawContent, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($match in $regexMatches) {
                $normalized = ConvertTo-GpoBackupGuid -Value $match.Groups['Value'].Value
                if (Test-GpoBackupGuid -Value $normalized) {
                    $results += $normalized
                }
            }
        }
    }

    return @($results | Select-Object -Unique)
}

function Get-GpoBackupDescriptorInfo {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$BackupXml
    )

    [xml]$descriptorXml = Get-Content -LiteralPath $BackupXml.FullName -ErrorAction Stop

    $gpoName = Get-FirstXmlInnerText -Document $descriptorXml -LocalNames @('GPODisplayName', 'Name')
    $backupId = Get-FirstXmlInnerText -Document $descriptorXml -LocalNames @('BackupId', 'ID')

    $gpoName = $gpoName.Trim()
    $backupId = $backupId.Trim()

    if ([string]::IsNullOrWhiteSpace($backupId)) {
        $backupId = $BackupXml.Directory.Name
    }

    if ((Test-GpoBackupGuid -Value $backupId) -and -not [string]::IsNullOrWhiteSpace($gpoName)) {
        return [PSCustomObject]@{
            ID             = $backupId
            GPODisplayName = $gpoName
            BackupPath     = $BackupXml.Directory.FullName
            CandidateBackupIds = @((ConvertTo-GpoBackupGuid -Value $backupId))
            CandidateBackupNames = @($gpoName)
        }
    }

    return $null
}

function Get-GpoBackupsFromManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [hashtable]$BackupFolders
    )

    $manifestCandidates = @(
        (Join-Path $Path 'manifest.xml'),
        (Join-Path $Path 'Manifest.xml')
    )

    foreach ($manifestPath in $manifestCandidates) {
        if (-not (Test-Path -LiteralPath $manifestPath)) {
            continue
        }

        try {
            [xml]$manifestXml = Get-Content -LiteralPath $manifestPath -ErrorAction Stop
            $backupNodes = $manifestXml.SelectNodes("//*[local-name()='BackupInst']")
            $manifestBackups = @()

            foreach ($node in $backupNodes) {
                $idNode = $node.SelectSingleNode("./*[local-name()='ID' or local-name()='BackupId']")
                $nameNode = $node.SelectSingleNode("./*[local-name()='GPODisplayName' or local-name()='Name']")

                $backupId = if ($idNode) { [string]$idNode.InnerText } else { '' }
                $gpoName = if ($nameNode) { [string]$nameNode.InnerText } else { '' }

                $backupId = $backupId.Trim()
                $gpoName = $gpoName.Trim()
                $normalizedId = ConvertTo-GpoBackupGuid -Value $backupId

                if ((Test-GpoBackupGuid -Value $normalizedId) -and -not [string]::IsNullOrWhiteSpace($gpoName) -and $BackupFolders.ContainsKey($normalizedId)) {
                    $manifestBackups += [PSCustomObject]@{
                        ID             = $BackupFolders[$normalizedId].Name
                        GPODisplayName = $gpoName
                        BackupPath     = $BackupFolders[$normalizedId].FullName
                        CandidateBackupIds = @($normalizedId)
                        CandidateBackupNames = @($gpoName)
                    }
                } elseif ((Test-GpoBackupGuid -Value $normalizedId) -and -not [string]::IsNullOrWhiteSpace($gpoName)) {
                    Write-Log -Message "Ignoring manifest entry '$gpoName' with backup ID '$backupId' because no matching backup folder exists under '$Path'." -Level WARN
                }
            }

            if ($manifestBackups.Count -gt 0) {
                Write-Log -Message "Discovered $($manifestBackups.Count) GPO backup(s) from manifest '$manifestPath'." -Level INFO
                return @($manifestBackups)
            }

            Write-Log -Message "Manifest '$manifestPath' was found but did not contain any usable backup entries." -Level WARN
        } catch {
            Write-Log -Message "Failed parsing manifest '$manifestPath': $($_.Exception.Message)" -Level WARN
        }
    }

    return @()
}

function Get-GpoBackupsFromFolders {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$BackupFolders,

        [Parameter(Mandatory = $false)]
        [object[]]$ExistingBackups = @()
    )

    $existingIds = @{}
    foreach ($backup in @($ExistingBackups)) {
        $existingIds[(ConvertTo-GpoBackupGuid -Value $backup.ID)] = $true
    }

    $discovered = @()
    foreach ($entry in $BackupFolders.GetEnumerator()) {
        if ($existingIds.ContainsKey($entry.Key)) {
            continue
        }

        $folder = $entry.Value
        $xmlFiles = @(
            'Backup.xml',
            'backup.xml',
            'bkupInfo.xml',
            'gpreport.xml'
        ) | ForEach-Object {
            $candidatePath = Join-Path $folder.FullName $_
            if (Test-Path -LiteralPath $candidatePath) {
                Get-Item -LiteralPath $candidatePath -ErrorAction SilentlyContinue
            }
        } | Where-Object { $_ }

        $gpoName = ''

        foreach ($xmlFile in $xmlFiles) {
            try {
                $gpoName = Get-GpoDisplayNameFromXmlFile -XmlFile $xmlFile
                if (-not [string]::IsNullOrWhiteSpace($gpoName)) {
                    break
                }
            } catch {
                Write-Log -Message "Failed parsing backup metadata XML '$($xmlFile.FullName)' while resolving backup folder '$($folder.Name)': $($_.Exception.Message)" -Level WARN
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($gpoName)) {
            $guidCandidates = @((ConvertTo-GpoBackupGuid -Value $folder.Name))

            $metadataIdFiles = @($xmlFiles | Where-Object { $_.Name -imatch '^(backup|bkupinfo)\.xml$' })
            foreach ($xmlFile in $metadataIdFiles) {
                try {
                    $guidCandidates += Get-GpoGuidCandidatesFromXmlFile -XmlFile $xmlFile
                } catch {
                    Write-Log -Message "Failed extracting GUID candidates from '$($xmlFile.FullName)' while resolving backup folder '$($folder.Name)': $($_.Exception.Message)" -Level WARN
                }
            }

            $guidCandidates = @($guidCandidates | Where-Object { Test-GpoBackupGuid -Value $_ } | Select-Object -Unique)

            $discovered += [PSCustomObject]@{
                ID             = $folder.Name
                GPODisplayName = $gpoName
                BackupPath     = $folder.FullName
                CandidateBackupIds = $guidCandidates
                CandidateBackupNames = @($gpoName)
            }
        } else {
            Write-Log -Message "Skipping backup folder '$($folder.FullName)': Could not determine a GPO display name from its XML content." -Level WARN
        }
    }

    return @($discovered)
}

function Get-ErrorMessageText {
    param(
        [Parameter(Mandatory = $true)]
        $ErrorRecord
    )

    if ($null -ne $ErrorRecord.Exception -and -not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.Exception.Message)) {
        return [string]$ErrorRecord.Exception.Message
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.FullyQualifiedErrorId)) {
        return [string]$ErrorRecord.FullyQualifiedErrorId
    }

    $asString = ($ErrorRecord | Out-String).Trim()
    if (-not [string]::IsNullOrWhiteSpace($asString)) {
        return $asString
    }

    return 'Unknown error (no message returned by PowerShell error record).'
}

function Test-GpoBackupFolderIntegrity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BackupFolderPath,

        [Parameter(Mandatory = $false)]
        [string]$ExpectedDisplayName
    )

    $result = [ordered]@{
        IsValid         = $false
        DescriptorPath  = ''
        Reason          = ''
    }

    if ([string]::IsNullOrWhiteSpace($BackupFolderPath) -or -not (Test-Path -LiteralPath $BackupFolderPath)) {
        $result.Reason = "Backup folder does not exist: '$BackupFolderPath'."
        return [PSCustomObject]$result
    }

    $descriptorCandidates = @(
        (Join-Path $BackupFolderPath 'Backup.xml'),
        (Join-Path $BackupFolderPath 'backup.xml'),
        (Join-Path $BackupFolderPath 'bkupInfo.xml')
    ) | Where-Object { Test-Path -LiteralPath $_ }

    if ($descriptorCandidates.Count -eq 0) {
        $result.Reason = "No backup descriptor XML found (expected Backup.xml/backup.xml/bkupInfo.xml)."
        return [PSCustomObject]$result
    }

    $descriptorPath = $descriptorCandidates[0]
    $result.DescriptorPath = $descriptorPath

    try {
        [xml]$descriptorXml = Get-Content -LiteralPath $descriptorPath -ErrorAction Stop
        $xmlName = Get-FirstXmlInnerText -Document $descriptorXml -LocalNames @('GPODisplayName', 'Name')
        $xmlBackupId = Get-FirstXmlInnerText -Document $descriptorXml -LocalNames @('BackupId', 'ID')

        if ([string]::IsNullOrWhiteSpace($xmlName) -and [string]::IsNullOrWhiteSpace($ExpectedDisplayName)) {
            $result.Reason = "Descriptor XML parsed but does not contain GPODisplayName/Name."
            return [PSCustomObject]$result
        }

        if (-not [string]::IsNullOrWhiteSpace($xmlBackupId)) {
            $normalizedId = ConvertTo-GpoBackupGuid -Value $xmlBackupId
            if (-not (Test-GpoBackupGuid -Value $normalizedId)) {
                $result.Reason = "Descriptor XML contains an invalid BackupId/ID value: '$xmlBackupId'."
                return [PSCustomObject]$result
            }
        }

        $result.IsValid = $true
        return [PSCustomObject]$result
    } catch {
        $result.Reason = "Descriptor XML '$descriptorPath' could not be parsed: $($_.Exception.Message)"
        return [PSCustomObject]$result
    }
}

$script:GpoImportStats = [ordered]@{
    BackupsDiscovered          = 0
    SkippedDefaultPolicy       = 0
    SkippedDuplicateCollision  = 0
    SkippedExisting            = 0
    ImportAttempted            = 0
    Imported                   = 0
    ImportFailed               = 0
    ImportedWithoutMigTable    = 0
    InvalidBackupDataDetected  = 0
    CorruptedBackupNames       = @()
    ExistingGpoNames           = @()
    NewlyImportedGpoNames      = @()
    FailedGpoNames             = @()
}

Invoke-Safely -ScriptBlock {
    # 1. Prefer authoritative discovery via GroupPolicy cmdlet.
    $backupFolders = Get-GpoBackupFolderMap -Path $BackupPath
    $backups = @()

    try {
        $cmdletBackups = @(Get-GPOBackup -All -Path $BackupPath -ErrorAction Stop)
        foreach ($item in $cmdletBackups) {
            if ($item -and $item.Id -and -not [string]::IsNullOrWhiteSpace([string]$item.DisplayName)) {
                $backups += [PSCustomObject]@{
                    ID             = ([string]$item.Id)
                    GPODisplayName = ([string]$item.DisplayName)
                    BackupPath     = $BackupPath
                    CandidateBackupIds = @((ConvertTo-GpoBackupGuid -Value ([string]$item.Id)) )
                    CandidateBackupNames = @(([string]$item.DisplayName))
                }
            }
        }

        if ($backups.Count -gt 0) {
            Write-Log -Message "Discovered $($backups.Count) GPO backup(s) using Get-GPOBackup -All." -Level INFO
        }
    } catch {
        Write-Log -Message "Get-GPOBackup discovery failed at '$BackupPath': $($_.Exception.Message). Falling back to manifest/folder discovery." -Level WARN
    }

    # 2. Fall back to manifest + folder parsing only if cmdlet-based discovery returned nothing.
    if ($backups.Count -eq 0) {
        $backups = @(Get-GpoBackupsFromManifest -Path $BackupPath -BackupFolders $backupFolders)
    }

    $folderBackups = @(Get-GpoBackupsFromFolders -BackupFolders $backupFolders -ExistingBackups $backups)
    if ($folderBackups.Count -gt 0) {
        Write-Log -Message "Discovered $($folderBackups.Count) additional GPO backup(s) by inspecting backup folders directly." -Level INFO
        $backups += $folderBackups
    }

    if ($backups.Count -eq 0) {
        $backupXmlFiles = Get-ChildItem -Path $BackupPath -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "^(?:bkupInfo|backup)\.xml$" }

        foreach ($backupXml in $backupXmlFiles) {
            try {
                $backupInfo = Get-GpoBackupDescriptorInfo -BackupXml $backupXml
                if ($backupInfo) {
                    $backups += $backupInfo
                } else {
                    Write-Log -Message "Skipped backup descriptor '$($backupXml.FullName)': Could not extract a valid GPO display name and backup ID." -Level WARN
                }
            } catch {
                Write-Log -Message "Failed parsing backup descriptor '$($backupXml.FullName)': $($_.Exception.Message)" -Level WARN
            }
        }
    }

    # De-duplicate by actual backup folder ID in case multiple discovery methods find the same backup.
    $backups = @($backups | Group-Object { ConvertTo-GpoBackupGuid -Value $_.ID } | ForEach-Object { $_.Group | Select-Object -First 1 })

    if (-not $backups -or $backups.Count -eq 0) {
        throw "No valid GPO backups were found in '$BackupPath'."
    }
    $script:GpoImportStats.BackupsDiscovered = @($backups).Count

    # 2. Import directly from each backup folder to avoid depending on a potentially stale/invalid root manifest.xml.

    # Pre-compute target names and detect collisions before importing.
    $plan = @()
    foreach ($b in $backups) {
        $sourceName = [string]$b.GPODisplayName
        $backupId = [string]$b.ID
        $backupFolderPath = [string]$b.BackupPath
        $candidateBackupIds = @($b.CandidateBackupIds)
        if (-not $candidateBackupIds -or $candidateBackupIds.Count -eq 0) {
            $candidateBackupIds = @((ConvertTo-GpoBackupGuid -Value $backupId))
        }

        $candidateBackupNames = @($b.CandidateBackupNames)
        if (-not $candidateBackupNames -or $candidateBackupNames.Count -eq 0) {
            $candidateBackupNames = @($sourceName)
        }

        $targetName = $sourceName
        $isDefaultPolicy = $defaultPolicyNames -contains $sourceName
        $skipByDefaultPolicyRule = $false

        if ($isDefaultPolicy) {
            if ($effectiveDefaultPolicyMode -eq 'Skip') {
                $skipByDefaultPolicyRule = $true
            } elseif ($effectiveDefaultPolicyMode -eq 'Rename') {
                $targetName = "$sourceName$DefaultPolicyNameSuffix"
            }
        }

        $plan += [PSCustomObject]@{
            SourceName               = $sourceName
            TargetName               = $targetName
            BackupId                 = $backupId
            BackupPath               = $backupFolderPath
            CandidateBackupIds       = $candidateBackupIds
            CandidateBackupNames     = $candidateBackupNames
            IsDefaultPolicy          = $isDefaultPolicy
            SkipByDefaultPolicyRule  = $skipByDefaultPolicyRule
            SkipByDuplicateCollision = $false
        }
    }

    $activePlan = @($plan | Where-Object { -not $_.SkipByDefaultPolicyRule })
    $duplicateGroups = @($activePlan | Group-Object TargetName | Where-Object { $_.Count -gt 1 })
    foreach ($dup in $duplicateGroups) {
        $entries = @($dup.Group)
        $keep = $entries | Select-Object -First 1
        $dupeIds = ($entries | ForEach-Object { $_.BackupId }) -join ', '
        Write-Log -Message "Duplicate GPO backup names detected for target '$($dup.Name)' (Backup IDs: $dupeIds). Keeping first backup '$($keep.BackupId)' and skipping the rest." -Level WARN

        $entries | Select-Object -Skip 1 | ForEach-Object {
            $_.SkipByDuplicateCollision = $true
        }
    }
    
    $abortRemainingImports = $false
    foreach ($entry in $plan) {
        if ($abortRemainingImports) {
            Write-Log -Message "Stopping remaining GPO imports due to pattern of invalid backup data detected. Re-export GPO backups before retrying." -Level ERROR
            break
        }

        $gpoName = $entry.SourceName
        $backupId = $entry.BackupId
        $backupFolderPath = $entry.BackupPath
        $candidateBackupIds = @($entry.CandidateBackupIds)
        $candidateBackupNames = @($entry.CandidateBackupNames)
        $targetGpoName = $entry.TargetName

        if ($entry.SkipByDefaultPolicyRule) {
            $script:GpoImportStats.SkippedDefaultPolicy++
            Write-Log -Message "Skipping '$gpoName' by default. Built-in default domain policies should typically be managed natively in the target domain." -Level WARN
            continue
        }

        if ($entry.IsDefaultPolicy -and $effectiveDefaultPolicyMode -eq 'Rename') {
            Write-Log -Message "Importing built-in policy '$gpoName' as renamed copy '$targetGpoName'." -Level WARN
        }

        if ($entry.SkipByDuplicateCollision) {
            $script:GpoImportStats.SkippedDuplicateCollision++
            Write-Log -Message "Skipping duplicate backup '$backupId' for GPO '$gpoName' because target name '$targetGpoName' already has a selected backup in this run." -Level WARN
            continue
        }
        
        # Check if GPO exists to ensure idempotency
        $existing = Get-GPO -Name $targetGpoName -Server $TargetDomain -ErrorAction SilentlyContinue
        if ($existing) {
            if ($script:GpoImportStats.ExistingGpoNames -notcontains $targetGpoName) {
                $script:GpoImportStats.ExistingGpoNames += $targetGpoName
            }
            
            if (-not $Force) {
                $script:GpoImportStats.SkippedExisting++
                Write-Log -Message "GPO '$targetGpoName' already exists in target domain. Skipping (Use -Force to overwrite)." -Level WARN
                continue
            } elseif ($Force) {
                Write-Log -Message "GPO '$targetGpoName' already exists but -Force flag set. Will attempt to update via re-import." -Level INFO
            }
        }
        
        # If SkipExistingOnly mode, don't attempt any imports
        if ($SkipExistingOnly -and $existing) {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($backupFolderPath) -or -not (Test-Path -LiteralPath $backupFolderPath)) {
            $script:GpoImportStats.ImportFailed++
            Write-Log -Message "Skipping '$gpoName': Backup folder was not resolved or no longer exists ('$backupFolderPath')." -Level ERROR
            continue
        }

        Write-Log -Message "Importing GPO: '$gpoName' as '$targetGpoName' (ID: $backupId, Folder: $backupFolderPath)" -Level INFO

        $backupIntegrity = Test-GpoBackupFolderIntegrity -BackupFolderPath $backupFolderPath -ExpectedDisplayName $gpoName
        if (-not $backupIntegrity.IsValid) {
            $script:GpoImportStats.ImportFailed++
            Write-Log -Message "Failed pre-validation for backup '$gpoName' at '$backupFolderPath': $($backupIntegrity.Reason)" -Level ERROR
            continue
        }
        Write-Log -Message "Backup descriptor validation passed for '$gpoName' using '$($backupIntegrity.DescriptorPath)'." -Level INFO

        $importBase = @{
            TargetName     = $targetGpoName
            Domain         = $TargetDomain
            CreateIfNeeded = $true
            ErrorAction    = 'Stop'
        }

        $strategies = @()

        foreach ($candidateId in @($candidateBackupIds | Where-Object { $_ } | Select-Object -Unique)) {
            $guid = ConvertTo-GpoBackupGuid -Value $candidateId
            if (-not (Test-GpoBackupGuid -Value $guid)) { continue }

            $strategies += @{
                Name = "BackupIdRoot:$guid"
                Params = @{
                    BackupId = $guid
                    Path     = $BackupPath
                }
            }

            $strategies += @{
                Name = "BackupIdFolder:$guid"
                Params = @{
                    BackupId = $guid
                    Path     = $backupFolderPath
                }
            }
        }

        foreach ($candidateName in @($candidateBackupNames | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)) {
            $strategies += @{
                Name = "BackupNameRoot:$candidateName"
                Params = @{
                    BackupGpoName = [string]$candidateName
                    Path          = $BackupPath
                }
            }
        }

        $importedCurrent = $false
        $lastErrorMessage = ''
        $stopDueToInvalidData = $false

        if ($strategies.Count -eq 0) {
            $lastErrorMessage = "No import strategies were generated for this backup. CandidateBackupIds='$($candidateBackupIds -join ',')'; CandidateBackupNames='$($candidateBackupNames -join ',')'."
            Write-Log -Message "No import strategies available for '$targetGpoName'. $lastErrorMessage" -Level ERROR
        }

        Write-Log -Message "Prepared $($strategies.Count) import strategy(ies) for '$targetGpoName'." -Level INFO

        if ($PSCmdlet.ShouldProcess($targetGpoName, "Import GPO") -and -not $WhatIfPreference) {
            $script:GpoImportStats.ImportAttempted++

            foreach ($strategy in $strategies) {
                $attemptParams = @{} + $importBase + $strategy.Params
                if ($MigrationTablePath) {
                    $attemptParams.MigrationTable = $MigrationTablePath
                }

                try {
                    Import-GPO @attemptParams | Out-Null
                    $script:GpoImportStats.Imported++
                    $importedCurrent = $true
                    if ($script:GpoImportStats.NewlyImportedGpoNames -notcontains $targetGpoName) {
                        $script:GpoImportStats.NewlyImportedGpoNames += $targetGpoName
                    }
                    Write-Log -Message "Successfully imported '$targetGpoName' using strategy '$($strategy.Name)'." -Level INFO
                    break
                } catch {
                    $lastErrorMessage = Get-ErrorMessageText -ErrorRecord $_
                    Write-Log -Message "Strategy '$($strategy.Name)' failed for '$targetGpoName': $lastErrorMessage" -Level WARN

                    if ($strategy.Name -like 'BackupIdRoot:*' -and $lastErrorMessage -match 'The data is invalid|0x8007000D') {
                        $script:GpoImportStats.InvalidBackupDataDetected++
                        if ($script:GpoImportStats.CorruptedBackupNames -notcontains $targetGpoName) {
                            $script:GpoImportStats.CorruptedBackupNames += $targetGpoName
                        }
                        Write-Log -Message "Backup data for '$targetGpoName' appears to be corrupted or incompatible (0x8007000D). Skipping this GPO and continuing with next backup." -Level WARN
                        # Don't abort all imports; just skip this one GPO and continue trying others.
                        # A single corrupt backup doesn't invalidate the entire export of 100+ backups.
                    }

                    if ($MigrationTablePath -and $lastErrorMessage -match 'The data is invalid|0x8007000D') {
                        $retryParams = @{} + $attemptParams
                        $retryParams.Remove('MigrationTable') | Out-Null

                        try {
                            Import-GPO @retryParams | Out-Null
                            $script:GpoImportStats.Imported++
                            $script:GpoImportStats.ImportedWithoutMigTable++
                            $importedCurrent = $true
                            if ($script:GpoImportStats.NewlyImportedGpoNames -notcontains $targetGpoName) {
                                $script:GpoImportStats.NewlyImportedGpoNames += $targetGpoName
                            }
                            Write-Log -Message "Imported '$targetGpoName' using strategy '$($strategy.Name)' without Migration Table after migration table validation failed." -Level WARN
                            break
                        } catch {
                            $lastErrorMessage = Get-ErrorMessageText -ErrorRecord $_
                            Write-Log -Message "Strategy '$($strategy.Name)' retry without Migration Table failed for '$targetGpoName': $lastErrorMessage" -Level WARN
                        }
                    }

                    if ($stopDueToInvalidData) {
                        Write-Log -Message "Stopping additional strategy attempts for '$targetGpoName' after confirmed invalid backup data on primary BackupId strategy." -Level WARN
                        # Skip this GPO but continue with others; don't abort the entire import.
                        # Continue to next GPO in plan instead of aborting all imports.
                        break
                    }
                }
            }
        } elseif ([string]::IsNullOrWhiteSpace($lastErrorMessage)) {
            $lastErrorMessage = "Import was not attempted because ShouldProcess returned false or WhatIf mode is active."
            Write-Log -Message "No import attempt executed for '$targetGpoName': $lastErrorMessage" -Level WARN
        }

        if ([string]::IsNullOrWhiteSpace($lastErrorMessage)) {
            $lastErrorMessage = "Import attempt did not complete successfully, but no specific error message was captured."
        }

        if (-not $importedCurrent) {
            $script:GpoImportStats.ImportFailed++
            if ($script:GpoImportStats.FailedGpoNames -notcontains $targetGpoName) {
                $script:GpoImportStats.FailedGpoNames += $targetGpoName
            }
            Write-Log -Message "Failed to import '$targetGpoName' (source backup '$gpoName'): $lastErrorMessage" -Level ERROR
        }
    }
} -Operation "Import GPOs"

$warningCount =
    $script:GpoImportStats.SkippedDefaultPolicy +
    $script:GpoImportStats.SkippedDuplicateCollision +
    $script:GpoImportStats.SkippedExisting +
    $script:GpoImportStats.ImportFailed

$summary = "Import GPOs summary: BackupsDiscovered=$($script:GpoImportStats.BackupsDiscovered), Attempted=$($script:GpoImportStats.ImportAttempted), Imported=$($script:GpoImportStats.Imported), ImportedWithoutMigTable=$($script:GpoImportStats.ImportedWithoutMigTable), Failed=$($script:GpoImportStats.ImportFailed), InvalidBackupDataDetected=$($script:GpoImportStats.InvalidBackupDataDetected), SkippedExisting=$($script:GpoImportStats.SkippedExisting), SkippedDefaultPolicy=$($script:GpoImportStats.SkippedDefaultPolicy), SkippedDuplicateCollision=$($script:GpoImportStats.SkippedDuplicateCollision)"

if ($script:GpoImportStats.Imported -eq 0 -and $script:GpoImportStats.BackupsDiscovered -gt 0) {
    Write-Host "[!] WARNING: Import-GPOs created 0 new GPOs. See logs for skip/failure reasons." -ForegroundColor Yellow
}

if ($script:GpoImportStats.InvalidBackupDataDetected -gt 0) {
    # Detect if this is a systematic issue (all/nearly all failing with 0x8007000D)
    # or just isolated corruption
    $isSystematicCorruption = ($script:GpoImportStats.InvalidBackupDataDetected -gt 10) -and 
                              ($script:GpoImportStats.InvalidBackupDataDetected / $script:GpoImportStats.BackupsDiscovered -gt 0.9)
    
    if ($isSystematicCorruption) {
        Write-Host "[!] CRITICAL: Nearly ALL backups are showing 0x8007000D errors." -ForegroundColor Red
        Write-Host "[!] This suggests an ENVIRONMENTAL ISSUE, not backup corruption:" -ForegroundColor Red
        Write-Host "[!]   - RSAT/GroupPolicy module may not be installed/loaded" -ForegroundColor Red
        Write-Host "[!]   - Permissions issue (not running as Domain Admin)" -ForegroundColor Red
        Write-Host "[!]   - Network connectivity issue to target domain" -ForegroundColor Red
        Write-Host "[!]   - Backup path permissions issue" -ForegroundColor Red
    } else {
        $corruptList = $script:GpoImportStats.CorruptedBackupNames -join "', '"
        Write-Host "[!] WARNING: $($script:GpoImportStats.InvalidBackupDataDetected) GPO backup(s) appear invalid/incompatible for Import-GPO (0x8007000D): '$corruptList'" -ForegroundColor Yellow
        Write-Host "[!] These backups have been SKIPPED. To resolve, you can:" -ForegroundColor Yellow
        Write-Host "    1. Check if the original GPOs on the source domain are corrupted or inaccessible" -ForegroundColor Yellow
        Write-Host "    2. Try re-exporting these specific GPOs from the source domain" -ForegroundColor Yellow
        Write-Host "    3. Check RSAT version compatibility between environments" -ForegroundColor Yellow
        Write-Host "[!] Other GPO backups have been imported normally." -ForegroundColor Cyan
    }
}

# Show detailed status for re-run scenarios
if ($script:GpoImportStats.ExistingGpoNames.Count -gt 0) {
    Write-Host "`n[*] GPOs ALREADY EXISTED (Skipped in this run):" -ForegroundColor Cyan
    $script:GpoImportStats.ExistingGpoNames | Sort-Object | ForEach-Object {
        Write-Host "    ✓ $_" -ForegroundColor Green
    }
    if ($Force) {
        Write-Host "[*] (Use -Force next time to overwrite these)" -ForegroundColor Yellow
    }
}

if ($script:GpoImportStats.NewlyImportedGpoNames.Count -gt 0) {
    Write-Host "`n[+] GPOs NEWLY IMPORTED (Added in this run):" -ForegroundColor Green
    $script:GpoImportStats.NewlyImportedGpoNames | Sort-Object | ForEach-Object {
        Write-Host "    ✓ $_" -ForegroundColor Green
    }
}

if ($script:GpoImportStats.FailedGpoNames.Count -gt 0) {
    Write-Host "`n[-] GPOs FAILED TO IMPORT (Retry in next run):" -ForegroundColor Red
    $script:GpoImportStats.FailedGpoNames | Sort-Object | ForEach-Object {
        Write-Host "    ✗ $_" -ForegroundColor Red
    }
}

# Determine if this is a fatal failure. Only fail if:
# 1. We attempted imports but NOTHING succeeded, or
# 2. Failure rate is extremely high (>85% of attempted imports failed)
$failureRate = if ($script:GpoImportStats.ImportAttempted -gt 0) {
    [math]::Round(($script:GpoImportStats.ImportFailed / $script:GpoImportStats.ImportAttempted) * 100, 2)
} else {
    0
}

$fatalImportFailure =
    (($script:GpoImportStats.ImportAttempted -gt 0) -and ($script:GpoImportStats.Imported -eq 0)) -or
    ($failureRate -gt 85 -and $script:GpoImportStats.ImportAttempted -ge 10)

if ($warningCount -gt 0) {
    Write-Log -Message "Import GPOs completed with warnings. $summary" -Level WARN
} else {
    Write-Log -Message "Import GPOs completed. $summary" -Level INFO
}

# Provide guidance for re-run scenarios
$isPartialRun = $script:GpoImportStats.ExistingGpoNames.Count -gt 0
$isRetryRun = $script:GpoImportStats.FailedGpoNames.Count -gt 0

if ($isPartialRun -and -not $SkipExistingOnly) {
    Write-Host "`n[*] RE-RUN INFORMATION:" -ForegroundColor Yellow
    Write-Host "[*] This run found $($script:GpoImportStats.ExistingGpoNames.Count) GPO(s) already in the target domain." -ForegroundColor Yellow
    Write-Host "[*] If you need to overwrite existing GPOs, re-run with the -Force flag." -ForegroundColor Yellow
    Write-Host "[*] If you just want to verify what's already been imported, use -SkipExistingOnly flag." -ForegroundColor Yellow
}

if ($isRetryRun) {
    Write-Host "`n[*] RETRY INSTRUCTIONS:" -ForegroundColor Yellow
    Write-Host "[*] $($script:GpoImportStats.FailedGpoNames.Count) GPO(s) failed in this run and require attention." -ForegroundColor Yellow
    Write-Host "[*] To retry the failed imports on the next run:" -ForegroundColor Yellow
    Write-Host "[*]   1. Verify the backup data for failed GPOs (check source domain or re-export)" -ForegroundColor Yellow
    Write-Host "[*]   2. Simply re-run this script - it will skip successful ones and retry failed ones" -ForegroundColor Yellow
    Write-Host "[*]   3. Completed GPOs will be automatically skipped (unless using -Force)" -ForegroundColor Yellow
}

if ($SkipExistingOnly) {
    Write-Host "`n[*] SKIP-EXISTING-ONLY MODE:" -ForegroundColor Cyan
    Write-Host "[*] No imports were attempted in this run. Use this to:" -ForegroundColor Cyan
    Write-Host "[*]   - Audit which GPOs already exist in the target domain" -ForegroundColor Cyan
    Write-Host "[*]   - Verify migration progress without modifying anything" -ForegroundColor Cyan
}

if ($fatalImportFailure) {
    if ($failureRate -gt 85 -and $script:GpoImportStats.ImportAttempted -ge 10) {
        Write-Log -Message "Import GPOs completed but with severe failures: $failureRate% failure rate (>85%). See summary and preceding error details." -Level ERROR
        Write-Host " " -ForegroundColor Red
        Write-Host "[!] SYSTEMATIC IMPORT FAILURE DETECTED" -ForegroundColor Red
        Write-Host "[!] $failureRate% of import attempts failed ($($script:GpoImportStats.ImportFailed) of $($script:GpoImportStats.ImportAttempted))" -ForegroundColor Red
        Write-Host " " -ForegroundColor Red
        Write-Host "[*] Diagnostic Checklist:" -ForegroundColor Yellow
        Write-Host "    1. RSAT/GroupPolicy Module: Is it installed and available on this machine?" -ForegroundColor Yellow
        Write-Host "       - Run: Get-Module GroupPolicy -ListAvailable" -ForegroundColor Yellow
        Write-Host "    2. Import-GPO Cmdlet: Can this machine import GPOs?" -ForegroundColor Yellow
        Write-Host "       - Run: Get-Command Import-GPO" -ForegroundColor Yellow
        Write-Host "    3. Network/Permissions: Can this machine reach domain controllers?" -ForegroundColor Yellow
        Write-Host "       - Test: Test-NetConnection -ComputerName <DC> -Port 445" -ForegroundColor Yellow
        Write-Host "    4. Backup Integrity: Are the backup files actually valid?" -ForegroundColor Yellow
        Write-Host "       - Check: Look at first backup Backup.xml manually" -ForegroundColor Yellow
        Write-Host "    5. Domain Admin Rights: Are you running as Domain Admin?" -ForegroundColor Yellow
        Write-Host "       - Test: whoami /groups | findstr /C:Domain" -ForegroundColor Yellow
        Write-Host " " -ForegroundColor Red
        $msg = "Import-GPOs: Systematic failures detected ($failureRate% failure rate). See diagnostic checklist above. Common causes: missing RSAT, network connectivity, or backup corruption."
        Write-Host "[!] ERROR: $msg" -ForegroundColor Red
        throw $msg
    } else {
        Write-Log -Message "Import GPOs completed but succeeded with 0 imports from $($script:GpoImportStats.ImportAttempted) attempts. All imports failed; check backup integrity and RSAT versions." -Level ERROR
        Write-Host " " -ForegroundColor Red
        Write-Host "[!] ALL IMPORT ATTEMPTS FAILED" -ForegroundColor Red
        Write-Host " " -ForegroundColor Red
        Write-Host "[*] Diagnostic Checklist:" -ForegroundColor Yellow
        Write-Host "    1. RSAT/GroupPolicy Module: Is it installed?" -ForegroundColor Yellow
        Write-Host "       - Run: Get-Module GroupPolicy -ListAvailable" -ForegroundColor Yellow
        Write-Host "    2. Import-GPO Cmdlet: Is it available?" -ForegroundColor Yellow
        Write-Host "       - Run: Get-Command Import-GPO" -ForegroundColor Yellow
        Write-Host "    3. Network Connectivity: Can you reach the target domain?" -ForegroundColor Yellow
        Write-Host "       - Test: Test-NetConnection -ComputerName <DC> -Port 445" -ForegroundColor Yellow
        Write-Host "    4. Backup Paths: Are backup paths accessible?" -ForegroundColor Yellow
        Write-Host "       - Test: Get-ChildItem '$BackupPath' -ErrorAction Stop" -ForegroundColor Yellow
        Write-Host " " -ForegroundColor Red
        $msg = "Import-GPOs: All $($script:GpoImportStats.ImportAttempted) import attempts failed. Check RSAT installation, network connectivity, and backup access."
        Write-Host "[!] ERROR: $msg" -ForegroundColor Red
        throw $msg
    }
}