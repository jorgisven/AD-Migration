function Invoke-Safely {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ScriptBlock]$ScriptBlock,

        [string]$Operation = "Operation"
    )

    try {
        & $ScriptBlock
    }
    catch {
        Write-Log -Level ERROR -Message "$Operation failed: $($_.Exception.Message)"
        throw
    }
}