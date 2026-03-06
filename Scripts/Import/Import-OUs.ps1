<#
.SYNOPSIS
    Rebuild OU structure in target domain.

.DESCRIPTION
    Creates OUs based on the mapping document and applies delegation boundaries.
    Reads from the transformed mapping CSV to ensure correct hierarchy.
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain
)

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
Write-Log -Message "Loaded ADMigration module from $ModulePath" -Level INFO
if (-not (Get-Command Invoke-Safely -ErrorAction SilentlyContinue)) { throw "Invoke-Safely unavailable" }

$config = Get-ADMigrationConfig
$MapPath = Join-Path $config.TransformRoot 'Mapping'

# Find mapping file
$mapFiles = Get-ChildItem -Path $MapPath -Filter "OU_Map_*.csv" | Sort-Object LastWriteTime -Descending
if (-not $mapFiles) {
    Write-Log -Message "No OU mapping file found in $MapPath. Run Transform-OUMap.ps1 first." -Level ERROR
    throw "Missing OU Map"
}

$Mapping = Import-Csv $mapFiles[0].FullName
Write-Log -Message "Loaded OU map from $($mapFiles[0].Name)" -Level INFO

if (-not $TargetDomain) {
    $TargetDomain = (Get-ADDomain).DNSRoot
    Write-Log -Message "Target domain not specified, using current: $TargetDomain" -Level INFO
}

Invoke-Safely -ScriptBlock {
    foreach ($row in $Mapping) {
        if ($row.Action -eq 'Skip') { continue }

        $targetDN = $row.TargetDN
        
        # Parse DN to get Name and Parent Path
        # Regex looks for: OU=Name,ParentPath...
        if ($targetDN -match "^OU=([^,]+),(.*)$") {
            $name = $matches[1]
            $parentPath = $matches[2]
            
            try {
                # Try to create the OU
                if ($PSCmdlet.ShouldProcess($targetDN, "Create OU")) {
                    New-ADOrganizationalUnit -Name $name -Path $parentPath -Description $row.Description -Server $TargetDomain -ProtectedFromAccidentalDeletion $true -ErrorAction Stop
                    Write-Log -Message "Created OU: $targetDN" -Level INFO
                }
            } catch [Microsoft.ActiveDirectory.Management.ADIdentityAlreadyExistsException] {
                Write-Log -Message "OU already exists: $targetDN" -Level INFO
            } catch {
                Write-Log -Message "Failed to create $targetDN : $_" -Level ERROR
            }
        } else {
            Write-Log -Message "Skipping invalid TargetDN format: $targetDN" -Level WARN
        }
    }
} -Operation "Import OUs to $TargetDomain"