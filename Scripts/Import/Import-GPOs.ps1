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
    [string]$DefaultDomainPolicyMode = 'Prompt'
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

Write-Log -Message "Starting GPO Import to $TargetDomain..." -Level INFO

$script:GpoImportStats = [ordered]@{
    BackupsDiscovered          = 0
    SkippedDefaultPolicy       = 0
    SkippedDuplicateCollision  = 0
    SkippedExisting            = 0
    ImportAttempted            = 0
    Imported                   = 0
    ImportFailed               = 0
}

Invoke-Safely -ScriptBlock {
    # 1. Discover Backups directly from bkupInfo.xml files
    # We avoid the CSV and old manifest files because they may incorrectly map the GPO ID instead of the Backup ID.
    $backupXmlFiles = Get-ChildItem -Path $BackupPath -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "^(?:bkupInfo|backup)\.xml$" }
    $backups = @()

    foreach ($backupXml in $backupXmlFiles) {
        try {
            $name = ""

            # Use Select-Xml with XPath for robust, namespace-aware parsing.
            $ns = @{
                gpo = "http://www.microsoft.com/GroupPolicy/Settings"
                gpm = "http://www.microsoft.com/GroupPolicy/BackupManifest"
            }
            $nameNode = Select-Xml -Path $backupXml.FullName -XPath "//gpo:Name | //gpm:GPODisplayName" -Namespace $ns | Select-Object -First 1
            if ($nameNode) {
                $name = $nameNode.Node.InnerText.Trim()
            }

            # The Backup ID is strictly the name of the folder containing the XML
            $id = $backupXml.Directory.Name

            # Ensure folder name is actually a valid GUID format
            if ($id -match '^\{?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\}?$' -and -not [string]::IsNullOrWhiteSpace($name)) {
                $backups += [PSCustomObject]@{
                    ID             = $id
                    GPODisplayName = $name
                }
            } else {
                Write-Log -Message "Skipped folder '$id': Could not extract valid GPO Name or Backup ID." -Level WARN
            }
        } catch {
            Write-Log -Message "Failed parsing backup descriptor '$($backupXml.FullName)': $($_.Exception.Message)" -Level WARN
        }
    }

        # De-duplicate by backup ID in case of repeated reads.
        $backups = @($backups | Group-Object ID | ForEach-Object { $_.Group | Select-Object -First 1 })

    if (-not $backups -or $backups.Count -eq 0) {
        throw "No valid GPO backups were found in '$BackupPath'."
    }
    $script:GpoImportStats.BackupsDiscovered = @($backups).Count

    # 2. Always regenerate manifest.xml with verified Backup IDs
    $manifestTarget = Join-Path $BackupPath "manifest.xml"
    $manifestTargetAlt = Join-Path $BackupPath "Manifest.xml"
    
    if (Test-Path $manifestTarget) { Remove-Item -Path $manifestTarget -Force }
    if (Test-Path $manifestTargetAlt) { Remove-Item -Path $manifestTargetAlt -Force }
    
    Write-Log -Message "Generating verified manifest.xml to satisfy Import-GPO cmdlet requirements..." -Level INFO
        $sb = New-Object System.Text.StringBuilder
        $sb.AppendLine('<?xml version="1.0" encoding="utf-8"?>') | Out-Null
        $sb.AppendLine('<Backups xmlns="http://www.microsoft.com/GroupPolicy/BackupManifest">') | Out-Null
        foreach ($b in $backups) {
            $guidStr = $b.ID
            if ($guidStr -notmatch '^\{') { $guidStr = "{$guidStr}" }
            $escapedName = $b.GPODisplayName.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;").Replace('"',"&quot;").Replace("'","&apos;")
            $sb.AppendLine("  <BackupInst>") | Out-Null
            $sb.AppendLine("    <ID>$guidStr</ID>") | Out-Null
            $sb.AppendLine("    <GPODisplayName>$escapedName</GPODisplayName>") | Out-Null
            $sb.AppendLine("  </BackupInst>") | Out-Null
        }
        $sb.AppendLine('</Backups>') | Out-Null
        Set-Content -Path $manifestTarget -Value $sb.ToString() -Encoding UTF8

    # Pre-compute target names and detect collisions before importing.
    $plan = @()
    foreach ($b in $backups) {
        $sourceName = [string]$b.GPODisplayName
        $backupId = [string]$b.ID
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
    
    foreach ($entry in $plan) {
        $gpoName = $entry.SourceName
        $backupId = $entry.BackupId
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
        if ($existing -and -not $Force) {
            $script:GpoImportStats.SkippedExisting++
            Write-Log -Message "GPO '$targetGpoName' already exists. Skipping (Use -Force to overwrite)." -Level WARN
            continue
        }

        Write-Log -Message "Importing GPO: '$gpoName' as '$targetGpoName' (ID: $backupId)" -Level INFO
        
        $params = @{
            BackupId       = $backupId
            Path           = $BackupPath
            TargetName     = $targetGpoName
            Domain         = $TargetDomain
            CreateIfNeeded = $true
            ErrorAction    = 'Stop'
        }
        
        if ($MigrationTablePath) {
            $params.MigrationTable = $MigrationTablePath
        }
        
        try {
            if ($PSCmdlet.ShouldProcess($targetGpoName, "Import GPO") -and -not $WhatIfPreference) {
                $script:GpoImportStats.ImportAttempted++
                Import-GPO @params | Out-Null
                $script:GpoImportStats.Imported++
                Write-Log -Message "Successfully imported '$targetGpoName'" -Level INFO
            }
        } catch {
            $script:GpoImportStats.ImportFailed++
            Write-Log -Message "Failed to import '$targetGpoName' (source backup '$gpoName'): $_" -Level ERROR
        }
    }
} -Operation "Import GPOs"

$warningCount =
    $script:GpoImportStats.SkippedDefaultPolicy +
    $script:GpoImportStats.SkippedDuplicateCollision +
    $script:GpoImportStats.SkippedExisting +
    $script:GpoImportStats.ImportFailed

$summary = "Import GPOs summary: BackupsDiscovered=$($script:GpoImportStats.BackupsDiscovered), Attempted=$($script:GpoImportStats.ImportAttempted), Imported=$($script:GpoImportStats.Imported), Failed=$($script:GpoImportStats.ImportFailed), SkippedExisting=$($script:GpoImportStats.SkippedExisting), SkippedDefaultPolicy=$($script:GpoImportStats.SkippedDefaultPolicy), SkippedDuplicateCollision=$($script:GpoImportStats.SkippedDuplicateCollision)"

if ($script:GpoImportStats.Imported -eq 0 -and $script:GpoImportStats.BackupsDiscovered -gt 0) {
    Write-Host "[!] WARNING: Import-GPOs created 0 new GPOs. See logs for skip/failure reasons." -ForegroundColor Yellow
}

if ($warningCount -gt 0) {
    Write-Log -Message "Import GPOs completed with warnings. $summary" -Level WARN
} else {
    Write-Log -Message "Import GPOs completed. $summary" -Level INFO
}