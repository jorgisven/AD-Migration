# Determine module root
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load Private functions first (dependencies for Logging and Public)
$PrivatePath = Join-Path $ScriptDir 'Private'
if (Test-Path $PrivatePath) {
    Get-ChildItem -Path $PrivatePath -Filter *.ps1 -ErrorAction SilentlyContinue |
        ForEach-Object { . $_.FullName }
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
