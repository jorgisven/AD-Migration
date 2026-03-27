<#
.SYNOPSIS
    Validates GPO backup folder structure and manifest before attempting imports.

.DESCRIPTION
    Performs comprehensive validation of backup folders, descriptor XML files, and manifest.xml.
    Identifies issues that would cause Import-GPO to fail before they happen.

.PARAMETER BackupPath
    Path to the GPO_Backups folder to validate.

.EXAMPLE
    .\Validate-GpoBackupFolders.ps1 -BackupPath "C:\Users\Administrator\Documents\ADMigration\Export\GPO_Backups"
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$BackupPath
)

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) { throw "Write-Log unavailable" }

$config = Get-ADMigrationConfig
if ([string]::IsNullOrWhiteSpace($BackupPath)) {
    $BackupPath = Join-Path $config.ExportRoot 'GPO_Backups'
}

Write-Host "`n=== GPO Backup Folder Validation ===" -ForegroundColor Cyan
Write-Host "[*] Backup Path: $BackupPath" -ForegroundColor Cyan

# Check if path exists
if (-not (Test-Path $BackupPath)) {
    Write-Host "[-] ERROR: Backup path does not exist: $BackupPath" -ForegroundColor Red
    exit 1
}

# Initialize stats
$stats = [ordered]@{
    TotalBackupFolders     = 0
    ValidDescriptors       = 0
    WarningDescriptors     = 0
    MissingDescriptors     = 0
    InvalidDescriptors     = 0
    ManifestStatus         = 'Unknown'
    ManifestEntryCount     = 0
    FoldersFoundNoDescriptor = @()
    FoldersWithWarnings    = @()
    FoldersWithIssues      = @()
}

# Find all backup folders (GUID-named directories)
$backupFolders = @(Get-ChildItem -Path $BackupPath -Directory -ErrorAction SilentlyContinue | 
    Where-Object { $_.Name -match '^[\{\}0-9a-fA-F\-]+$' })

$stats.TotalBackupFolders = $backupFolders.Count
Write-Host "`n[*] Found $($stats.TotalBackupFolders) backup folder(s)" -ForegroundColor Cyan

# Validate each backup folder
foreach ($folder in $backupFolders) {
    $folderName = $folder.Name
    
    # Check for descriptor XML
    $descriptorCandidates = @(
        (Join-Path $folder.FullName 'Backup.xml'),
        (Join-Path $folder.FullName 'backup.xml'),
        (Join-Path $folder.FullName 'bkupInfo.xml')
    ) | Where-Object { Test-Path $_ }
    
    if ($descriptorCandidates.Count -eq 0) {
        $stats.MissingDescriptors++
        $stats.FoldersFoundNoDescriptor += $folderName
        Write-Host "    [-] $folderName - NO DESCRIPTOR XML FOUND" -ForegroundColor Red
        continue
    }
    
    $bestResult = $null

    foreach ($descriptorPath in $descriptorCandidates) {
        try {
            [xml]$descriptor = Get-Content -LiteralPath $descriptorPath -ErrorAction Stop

            $gpoNameNode = $descriptor.SelectSingleNode("//*[local-name()='GPODisplayName' or local-name()='DisplayName' or local-name()='GPOName' or local-name()='Name']")
            $backupIdNode = $descriptor.SelectSingleNode("//*[local-name()='BackupId' or local-name()='BackupID' or local-name()='ID']")
            $gpoGuidNode = $descriptor.SelectSingleNode("//*[local-name()='GPOGuid' or local-name()='GpoGuid' or local-name()='Guid' or local-name()='GPOID']")

            $gpoName = if ($gpoNameNode) { $gpoNameNode.InnerText } else { '' }
            $backupId = if ($backupIdNode) { $backupIdNode.InnerText } else { '' }
            $gpoGuid = if ($gpoGuidNode) { $gpoGuidNode.InnerText } else { '' }

            $hasName = -not [string]::IsNullOrWhiteSpace($gpoName)
            $hasIdentifier = (-not [string]::IsNullOrWhiteSpace($backupId)) -or (-not [string]::IsNullOrWhiteSpace($gpoGuid))
            $score = 0
            if ($hasName) { $score += 2 }
            if ($hasIdentifier) { $score += 1 }

            if (($bestResult -eq $null) -or ($score -gt $bestResult.Score)) {
                $bestResult = [pscustomobject]@{
                    Path = $descriptorPath
                    GpoName = $gpoName
                    HasName = $hasName
                    HasIdentifier = $hasIdentifier
                    Score = $score
                }
            }
        } catch {
            # Keep trying other descriptor candidates in this folder.
        }
    }

    if ($bestResult -eq $null) {
        $stats.InvalidDescriptors++
        $stats.FoldersWithIssues += $folderName
        Write-Host "    [!] $folderName - Descriptor parsing failed for all descriptor files" -ForegroundColor Yellow
    } elseif ($bestResult.HasName) {
        $stats.ValidDescriptors++
        Write-Host "    [+] $folderName - Valid: '$($bestResult.GpoName)'" -ForegroundColor Green
    } elseif ($bestResult.HasIdentifier) {
        $stats.WarningDescriptors++
        $stats.FoldersWithWarnings += $folderName
        Write-Host "    [~] $folderName - Descriptor missing display name but has backup identifier/GPO GUID" -ForegroundColor DarkYellow
    } else {
        $stats.InvalidDescriptors++
        $stats.FoldersWithIssues += $folderName
        Write-Host "    [!] $folderName - Descriptor missing both display name and backup identifiers" -ForegroundColor Yellow
    }
}

# Validate manifest.xml
Write-Host "`n[*] Validating manifest.xml..." -ForegroundColor Cyan
$manifestCandidates = @(
    (Join-Path $BackupPath 'manifest.xml'),
    (Join-Path $BackupPath 'Manifest.xml')
)

$manifestFound = $false
$manifestPath = $null

foreach ($candidate in $manifestCandidates) {
    if (Test-Path $candidate) {
        $manifestFound = $true
        $manifestPath = $candidate
        break
    }
}

if (-not $manifestFound) {
    Write-Host "    [-] No manifest.xml found" -ForegroundColor Red
    $stats.ManifestStatus = 'Missing'
} else {
    try {
        [xml]$manifest = Get-Content -LiteralPath $manifestPath -ErrorAction Stop
        $backupInstNodes = $manifest.SelectNodes("//*[local-name()='BackupInst']")
        $stats.ManifestEntryCount = $backupInstNodes.Count
        
        if ($backupInstNodes.Count -eq 0) {
            Write-Host "    [!] manifest.xml found but contains NO BackupInst entries" -ForegroundColor Yellow
            $stats.ManifestStatus = 'Empty'
        } else {
            Write-Host "    [+] manifest.xml contains $($backupInstNodes.Count) BackupInst entries" -ForegroundColor Green
            $stats.ManifestStatus = 'Valid'
        }
    } catch {
        Write-Host "    [!] manifest.xml parsing failed: $($_.Exception.Message)" -ForegroundColor Yellow
        $stats.ManifestStatus = 'Invalid'
    }
}

# Summary and recommendations
Write-Host "`n=== VALIDATION SUMMARY ===" -ForegroundColor Cyan
Write-Host "Total Backup Folders:           $($stats.TotalBackupFolders)"
Write-Host "Valid Descriptors:              $($stats.ValidDescriptors)" -ForegroundColor Green
Write-Host "Warning Descriptors:            $($stats.WarningDescriptors)" -ForegroundColor DarkYellow
Write-Host "Missing Descriptors:            $($stats.MissingDescriptors)" -ForegroundColor Red
Write-Host "Invalid Descriptors:            $($stats.InvalidDescriptors)" -ForegroundColor Yellow
Write-Host "Manifest Status:                $($stats.ManifestStatus)" -ForegroundColor Cyan
Write-Host "Manifest BackupInst Entries:    $($stats.ManifestEntryCount)"

Write-Host "`n=== IMPORT READINESS ===" -ForegroundColor Cyan

$validatedDescriptors = $stats.ValidDescriptors + $stats.WarningDescriptors
$isReady = ($validatedDescriptors -eq $stats.TotalBackupFolders) -and 
           ($stats.ManifestStatus -eq 'Valid') -and 
           ($stats.ManifestEntryCount -gt 0)

if ($isReady) {
    Write-Host "[OK] READY FOR IMPORT" -ForegroundColor Green
    Write-Host "[+] All backup folders have valid descriptor metadata" -ForegroundColor Green
    Write-Host "[+] manifest.xml is properly populated" -ForegroundColor Green
    if ($stats.WarningDescriptors -gt 0) {
        Write-Host "[~] Some descriptors omit display name, but include valid identifiers" -ForegroundColor DarkYellow
    }
} else {
    Write-Host "[X] NOT READY FOR IMPORT - Issues Detected" -ForegroundColor Red
    
    if ($stats.MissingDescriptors -gt 0) {
        $missingCount = $stats.MissingDescriptors
        Write-Host ("`n[!] Missing Descriptors ({0} folders):" -f $missingCount) -ForegroundColor Red
        $stats.FoldersFoundNoDescriptor | ForEach-Object {
            Write-Host "    - $_" -ForegroundColor Red
        }
        Write-Host "[*] Solution: Check if Backup.xml files exist in these folders," -ForegroundColor Yellow
        Write-Host "    or re-export these specific GPOs from the source domain." -ForegroundColor Yellow
    }
    
    if ($stats.InvalidDescriptors -gt 0) {
        $invalidCount = $stats.InvalidDescriptors
        Write-Host ("`n[!] Invalid Descriptors ({0} folders):" -f $invalidCount) -ForegroundColor Yellow
        $stats.FoldersWithIssues | ForEach-Object {
            Write-Host "    - $_" -ForegroundColor Yellow
        }
        Write-Host "[*] Solution: Review descriptor XML files for corrupt/missing data," -ForegroundColor Yellow
        Write-Host "    or re-export from source domain." -ForegroundColor Yellow
    }

    if ($stats.WarningDescriptors -gt 0) {
        $warningCount = $stats.WarningDescriptors
        Write-Host ("`n[~] Descriptor Warnings ({0} folders):" -f $warningCount) -ForegroundColor DarkYellow
        Write-Host "    Missing display name fields, but identifiers were found." -ForegroundColor DarkYellow
        Write-Host "    This usually indicates alternate backup XML schema and is often safe." -ForegroundColor DarkYellow
    }
    
    if ($stats.ManifestStatus -eq 'Missing') {
        Write-Host "`n[!] Manifest File Missing:" -ForegroundColor Red
        Write-Host "    Expected: $($manifestCandidates[0])" -ForegroundColor Red
        Write-Host "[*] Solution: Run Export-GPOReports with -Force to regenerate manifest," -ForegroundColor Yellow
        Write-Host "    or Import-GPOs will attempt folder-based fallback discovery." -ForegroundColor Yellow
    } elseif ($stats.ManifestStatus -eq 'Empty') {
        Write-Host "`n[!] Manifest File is EMPTY (no BackupInst entries):" -ForegroundColor Yellow
        Write-Host "    File: $manifestPath" -ForegroundColor Yellow
        Write-Host "[*] Solution: Regenerate manifest via Export-GPOReports -Force," -ForegroundColor Yellow
        Write-Host "    or the Import-GPOs folder-based fallback will attempt to recover." -ForegroundColor Yellow
    } elseif ($stats.ManifestStatus -eq 'Invalid') {
        Write-Host "`n[!] Manifest File is INVALID (parse error):" -ForegroundColor Yellow
        Write-Host "    File: $manifestPath" -ForegroundColor Yellow
        Write-Host "[*] Solution: Delete and regenerate via Export-GPOReports -Force" -ForegroundColor Yellow
    }
}

Write-Host "`n[*] For detailed logs, review the Import-GPOs log file." -ForegroundColor Cyan
Write-Host ""

exit 0
