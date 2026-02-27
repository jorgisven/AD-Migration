# Determine module root
$ScriptDir = $PSScriptRoot

# Load Private functions first (dependencies for Logging and Public)
$PrivatePath = Join-Path $ScriptDir 'Private'
if (Test-Path $PrivatePath) {
    $PrivateFiles = Get-ChildItem -Path $PrivatePath -Filter *.ps1
    
    if (-not ($PrivateFiles.Name -contains 'Invoke-Safely.ps1')) {
        Write-Warning "CRITICAL: Invoke-Safely.ps1 is missing from $PrivatePath"
    }

    foreach ($file in $PrivateFiles) {
        Write-Verbose "Loading private function: $($file.Name)"
        . $file.FullName
    }
} else {
    Write-Warning "Private directory not found at $PrivatePath"
}

# Then load Logging functions
$LoggingPath = Join-Path $ScriptDir 'Logging'
if (Test-Path $LoggingPath) {
    Get-ChildItem -Path $LoggingPath -Filter *.ps1 -ErrorAction SilentlyContinue |
        ForEach-Object { . $_.FullName }
}

# Finally load Public functions
$PublicPath = Join-Path $ScriptDir 'Public'
if (Test-Path $PublicPath) {
    Get-ChildItem -Path $PublicPath -Filter *.ps1 -ErrorAction SilentlyContinue |
        ForEach-Object { . $_.FullName }
}

# Optional: module initializer - only if the function was loaded
if (Get-Command Initialize-ADMigration -ErrorAction SilentlyContinue) {
    Initialize-ADMigration
}
