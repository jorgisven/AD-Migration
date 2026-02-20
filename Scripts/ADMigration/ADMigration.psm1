# Determine module root
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load all .ps1 files recursively (Public, Private, Logging, etc.)
Get-ChildItem -Path $ScriptDir -Recurse -Filter *.ps1 |
    Where-Object { $_.FullName -ne $MyInvocation.MyCommand.Path } |
    ForEach-Object {
        . $_.FullName
    }