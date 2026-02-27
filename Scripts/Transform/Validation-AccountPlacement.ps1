<#
.SYNOPSIS
    Validate account placement against target naming conventions.

.DESCRIPTION
    Analyzes the OU Map and Account Exports to flag accounts that would be moved
    to Target OUs where they don't match the expected naming convention.
    Useful for detecting manual mapping errors or accounts that need renaming.
#>

param(
    [Parameter(Mandatory = $false)]
    [hashtable]$Conventions = @{
        # Target OU Keyword (Regex) = Expected Object Name Pattern (Regex)
        "OU=Workstations" = "^WKS-"
        "OU=Servers"      = "^SRV-"
        "OU=Service"      = "^svc_"
        "OU=Admin"        = "^adm_"
        "OU=Laptops"      = "^LPT-"
    }
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO

$config = Get-ADMigrationConfig
$MapPath = Join-Path $config.TransformRoot 'Mapping'
$ExportPath = Join-Path $config.ExportRoot 'Security'
$AnalysisPath = Join-Path $config.TransformRoot 'Placement_Analysis'

if (-not (Test-Path $AnalysisPath)) { New-Item -ItemType Directory -Path $AnalysisPath -Force | Out-Null }

# 1. Load OU Map
$mapFiles = Get-ChildItem -Path $MapPath -Filter "OU_Map_*.csv" | Sort-Object LastWriteTime -Descending
if (-not $mapFiles) { throw "Missing OU Map. Run Transform-OUMap.ps1 first." }

$OUMap = @{}
Import-Csv $mapFiles[0].FullName | ForEach-Object { $OUMap[$_.SourceDN] = $_.TargetDN }
Write-Log -Message "Loaded OU Map with $($OUMap.Count) entries" -Level INFO

# 2. Define Account Types to Check
$AccountTypes = @(
    @{ Type="User";     FilePattern="Users_*.csv";     NameCol="SamAccountName" },
    @{ Type="Computer"; FilePattern="Computers_*.csv"; NameCol="ComputerName" },
    @{ Type="Group";    FilePattern="Groups_*.csv";    NameCol="Name" }
)

$Violations = @()

Invoke-Safely -ScriptBlock {
    foreach ($accType in $AccountTypes) {
        $files = Get-ChildItem -Path $ExportPath -Filter $accType.FilePattern | Sort-Object LastWriteTime -Descending
        if (-not $files) {
            Write-Log -Message "No export found for $($accType.Type)" -Level WARN
            continue
        }

        $accounts = Import-Csv $files[0].FullName
        Write-Log -Message "Analyzing $($accounts.Count) $($accType.Type) records..." -Level INFO

        foreach ($row in $accounts) {
            $dn = $row.DistinguishedName
            $name = $row.$($accType.NameCol)
            
            # Extract Parent DN (Source OU)
            # Regex removes the first RDN (CN=Name,) to get the parent
            $parentDN = $dn -replace "^(?:CN|OU)=[^,]+?,", ""
            
            if ($OUMap.ContainsKey($parentDN)) {
                $targetOU = $OUMap[$parentDN]
                
                # Check against conventions
                foreach ($key in $Conventions.Keys) {
                    if ($targetOU -match $key) {
                        $pattern = $Conventions[$key]
                        if ($name -notmatch $pattern) {
                            $Violations += [PSCustomObject]@{
                                Name        = $name
                                Type        = $accType.Type
                                SourceOU    = $parentDN
                                TargetOU    = $targetOU
                                Expected    = $pattern
                                Rule        = $key
                            }
                        }
                    }
                }
            }
        }
    }

    # Output Report
    $reportFile = Join-Path $AnalysisPath "Placement_Violations_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($Violations.Count -gt 0) {
        $Violations | Export-Csv -Path $reportFile -NoTypeInformation
        Write-Log -Message "Found $($Violations.Count) naming convention violations." -Level WARN
        
        Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
        Write-Host "Placement Analysis Complete"
        Write-Host "Found $($Violations.Count) accounts violating naming conventions." -ForegroundColor Red
        Write-Host "Report: $reportFile"
        Write-Host "Review the report and either:" -ForegroundColor Yellow
        Write-Host "  1. Update the OU Map to send these accounts elsewhere." -ForegroundColor Yellow
        Write-Host "  2. Rename the accounts in the source (or during migration)." -ForegroundColor Yellow
        Write-Host "----------------------------------------------------------------" -ForegroundColor Cyan
    } else {
        Write-Log -Message "No naming convention violations found." -Level INFO
        Write-Host "✅ All accounts match target OU naming conventions." -ForegroundColor Green
    }
} -Operation "Validate Account Placement"