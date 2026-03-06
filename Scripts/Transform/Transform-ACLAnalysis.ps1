<#
.SYNOPSIS
    Analyze ACLs for Tier-0/Tier-1 boundary validation.

.DESCRIPTION
    Compares exported ACLs to expected inheritance and privilege boundaries.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO

$config = Get-ADMigrationConfig
$SourceSecurityPath = Join-Path $config.ExportRoot 'Security'
$AnalysisPath = Join-Path $config.TransformRoot 'ACL_Analysis'

# Ensure transform directory exists
if (-not (Test-Path $AnalysisPath)) { New-Item -ItemType Directory -Path $AnalysisPath -Force | Out-Null }

Write-Log -Message "Starting ACL Analysis..." -Level INFO

# 1. Load Identity Map (SIDs -> Names) from Account Exports
$SidMap = @{}

Function Load-SidMap {
    param($Pattern, $Type)
    $files = Get-ChildItem -Path $SourceSecurityPath -Filter $Pattern | Sort-Object LastWriteTime -Descending
    if ($files) {
        $data = Import-Csv $files[0].FullName
        foreach ($row in $data) {
            if ($row.SID) { $SidMap[$row.SID] = "$($row.SamAccountName) ($Type)" }
        }
        Write-Log -Message "Loaded $($data.Count) $Type SIDs" -Level INFO
    }
}

Load-SidMap "Users_*.csv" "User"
Load-SidMap "Groups_*.csv" "Group"
Load-SidMap "Computers_*.csv" "Computer"
Load-SidMap "ServiceAccounts_*.csv" "Svc"

# 2. Load latest ACL Export
$aclFiles = Get-ChildItem -Path $SourceSecurityPath -Filter "ACLs_OUs_*.csv" | Sort-Object LastWriteTime -Descending
if (-not $aclFiles) {
    Write-Log -Message "No ACL export found. Run Export-ACLs.ps1 first." -Level ERROR
    return
}

$ACLs = Import-Csv $aclFiles[0].FullName
Write-Log -Message "Analyzing $($ACLs.Count) ACEs from $($aclFiles[0].Name)" -Level INFO

$AnalysisResults = [System.Collections.Generic.List[PSObject]]::new()

foreach ($ace in $ACLs) {
    $identity = $ace.IdentityReference
    $resolvedName = $identity
    $status = "OK"
    $notes = ""

    # Resolve SID if IdentityReference looks like a SID
    if ($identity -match "^S-1-5") {
        if ($SidMap.ContainsKey($identity)) {
            $resolvedName = $SidMap[$identity]
            $notes = "Resolved from SID"
        } else {
            $status = "UnknownSID"
            $notes = "SID not found in export"
        }
    }

    # Check for Risky Permissions
    $rights = $ace.ActiveDirectoryRights
    
    if ($rights -match "WriteDacl") {
        if ($status -eq "OK") { $status = "Risky" }
        $notes += "; CRITICAL: Can change permissions (WriteDacl)"
    }
    if ($rights -match "WriteOwner") {
        if ($status -eq "OK") { $status = "Risky" }
        $notes += "; CRITICAL: Can take ownership (WriteOwner)"
    }
    if ($rights -match "GenericAll|FullControl") {
        if ($status -eq "OK") { $status = "Risky" }
        $notes += "; HIGH: Full Control"
    }

    # Filter out boring default permissions (optional)
    if ($status -ne "OK" -or $ace.IsInherited -eq $false) {
        $AnalysisResults.Add([PSCustomObject]@{
            OU                    = $ace.DistinguishedName
            OriginalIdentity      = $identity
            ResolvedIdentity      = $resolvedName
            Access                = $ace.AccessControlType
            Rights                = $ace.ActiveDirectoryRights
            IsInherited           = $ace.IsInherited
            Status                = $status
            Notes                 = $notes.Trim("; ")
        })
    }
}

$reportFile = Join-Path $AnalysisPath "ACL_Analysis_Report.csv"
$AnalysisResults | Export-Csv -Path $reportFile -NoTypeInformation

# Generate Guide
$guideContent = @"
ACL Analysis Guide
==================

1. Status: UnknownSID
   - These are SIDs found in the ACLs that do not match any User, Group, or Computer exported from the source domain.
   - Cause: Deleted accounts, or security principals from external trusted domains.
   - Action: Generally safe to ignore/remove during migration unless they represent critical external trusts.

2. Status: Risky
   - These are permissions that allow an object to control other objects or the OU itself.
   - WriteDacl: The ability to write the DACL (change permissions). This allows a user to grant themselves Full Control.
   - WriteOwner: The ability to take ownership. Owners can implicitly change permissions.
   - GenericAll / FullControl: Complete control over the object.
   - Action: Review these carefully. 
     - If the Identity is "Domain Admins" or "Enterprise Admins", this is normal.
     - If the Identity is a standard user or non-admin group, this is a potential security hole (Privilege Escalation path).
"@
$guideFile = Join-Path $AnalysisPath "README_Analysis_Guide.txt"
Set-Content -Path $guideFile -Value $guideContent

Write-Log -Message "Analysis complete. Report saved to $reportFile" -Level INFO

Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
Write-Host "ACL Analysis Complete"
Write-Host "Report: $reportFile"
Write-Host "Found $(($AnalysisResults | Where-Object {$_.Status -eq 'UnknownSID'}).Count) Unknown SIDs" -ForegroundColor Yellow
Write-Host "Found $(($AnalysisResults | Where-Object {$_.Status -eq 'Risky'}).Count) Risky/High-Privilege ACEs" -ForegroundColor Red
Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan