<#
.SYNOPSIS
    Generate a GPO Migration Table for mapping security principals.

.DESCRIPTION
    Reads exported Users and Groups and creates a .migtable file.
    Used by Import-GPOs to map Source\User to Target\User dynamically.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDomain, # NetBIOS or FQDN

    [Parameter(Mandatory = $true)]
    [string]$TargetDomain  # NetBIOS or FQDN
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO

$config = Get-ADMigrationConfig
$SourceSecurityPath = Join-Path $config.ExportRoot 'Security'
$TransformPath = $config.TransformRoot

# Ensure transform directory exists
if (-not (Test-Path $TransformPath)) { New-Item -ItemType Directory -Path $TransformPath -Force | Out-Null }

# Check for optional UserMap.csv for custom username mappings
# Format: SourceSam,TargetSam
$UserMapPath = Join-Path $TransformPath 'UserMap.csv'
$UserMap = @{}
if (Test-Path $UserMapPath) {
    Import-Csv $UserMapPath | ForEach-Object {
        if ($_.SourceSam -and $_.TargetSam) { $UserMap[$_.SourceSam] = $_.TargetSam }
    }
    Write-Log -Message "Loaded $($UserMap.Count) custom user mappings from $UserMapPath" -Level INFO
}

Write-Log -Message "Generating Migration Table mapping $SourceDomain -> $TargetDomain" -Level INFO

$Mappings = @()

Function Add-Mapping {
    param($Pattern, $Type)
    $files = Get-ChildItem -Path $SourceSecurityPath -Filter $Pattern | Sort-Object LastWriteTime -Descending
    if ($files) {
        $data = Import-Csv $files[0].FullName
        foreach ($row in $data) {
            $targetSam = $row.SamAccountName
            if ($UserMap.ContainsKey($row.SamAccountName)) {
                $targetSam = $UserMap[$row.SamAccountName]
            }

            $Mappings += [PSCustomObject]@{
                Type        = $Type
                Source      = "$SourceDomain\$($row.SamAccountName)"
                Destination = "$TargetDomain\$targetSam"
            }
        }
        Write-Log -Message "Added $($data.Count) $Type mappings" -Level INFO
    }
}

# Map Groups (Global, DomainLocal, Universal)
# GPMC uses specific types, but 'GlobalGroup' is a safe default for most domain groups in migration tables
Add-Mapping "Groups_*.csv" "GlobalGroup"

# Map Users
Add-Mapping "Users_*.csv" "User"
Add-Mapping "ServiceAccounts_*.csv" "User"

# Map Computers (Less common in GPOs but possible)
Add-Mapping "Computers_*.csv" "Computer"

# Also map the Domain Root itself (for UNC paths like \\Source\SysVol)
$Mappings += [PSCustomObject]@{
    Type        = "Unknown"
    Source      = $SourceDomain
    Destination = $TargetDomain
}

# Build XML
$xmlHeader = '<?xml version="1.0" encoding="utf-8"?>'
$xmlStart  = '<MigrationTable xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
$xmlEnd    = '</MigrationTable>'

$xmlContent = New-Object System.Text.StringBuilder
$xmlContent.AppendLine($xmlHeader) | Out-Null
$xmlContent.AppendLine($xmlStart) | Out-Null

foreach ($map in $Mappings) {
    $xmlContent.AppendLine("  <Mapping>") | Out-Null
    $xmlContent.AppendLine("    <Type>$($map.Type)</Type>") | Out-Null
    $xmlContent.AppendLine("    <Source>$($map.Source)</Source>") | Out-Null
    $xmlContent.AppendLine("    <Destination>$($map.Destination)</Destination>") | Out-Null
    $xmlContent.AppendLine("  </Mapping>") | Out-Null
}

$xmlContent.AppendLine($xmlEnd) | Out-Null

$outputFile = Join-Path $TransformPath "GPO_MigrationTable.migtable"
Set-Content -Path $outputFile -Value $xmlContent.ToString() -Encoding UTF8

Write-Log -Message "Migration Table generated at $outputFile" -Level INFO
Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "Migration Table Created: $outputFile"
Write-Host "Contains $($Mappings.Count) mappings."
Write-Host "Import-GPOs.ps1 will now automatically use this file." -ForegroundColor Yellow
Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan