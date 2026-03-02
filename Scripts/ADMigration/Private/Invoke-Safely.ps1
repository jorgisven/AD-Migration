function Invoke-Safely {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ScriptBlock]$ScriptBlock,

        [string]$Operation = "Operation"
    )

    try {
        & $ScriptBlock
        Write-Log -Message "$Operation succeeded" -Level INFO
    }
    catch {
        Write-Log -Message "$Operation failed: $($_.Exception.Message)" -Level ERROR
        throw
    }
}
