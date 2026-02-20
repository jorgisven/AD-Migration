function Write-log {
        param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    $config = Get-ADMigrationConfig
    $LogRoot = $config.LogRoot
    if (-not (Test-Path $LogRoot)) {
        New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    }

    $DateStamp = (Get-Date).ToString('yyyy-MM-dd')
    $TimeStamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $LogFile   = Join-Path $LogRoot "$DateStamp.log"

    $Entry = "[$TimeStamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $Entry
}