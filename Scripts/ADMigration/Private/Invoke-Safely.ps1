function Invoke-Safely {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ScriptBlock]$ScriptBlock,

        [Parameter(Mandatory)]
        [string]$Operation
    )

    try {
        & $ScriptBlock
        Write-Log -Message "$Operation succeeded" -Level INFO
    }
    catch {
        Write-Log -Message "$Operation failed: $_" -Level ERROR
        throw
    }
}
