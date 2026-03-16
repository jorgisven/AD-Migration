<#
.SYNOPSIS
    Orchestrate the full transform process.

.DESCRIPTION
    Guides the user through the transform phase, which involves both automated script execution
    and required manual review and editing of mapping files. This is an interactive wizard.
#>

# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force

Add-Type -AssemblyName System.Windows.Forms

# --- Portable Data Detection & Staging ---
# If running from an extracted Migration Package, data is in ..\..\Export relative to this script.
# The module expects data in Documents\ADMigration. We must sync it if detected.
$PortableRoot = Split-Path -Parent (Split-Path -Parent $ScriptRoot)
$PortableExport = Join-Path $PortableRoot "Export"
$PortableTransform = Join-Path $PortableRoot "Transform"

if (Test-Path $PortableExport) {
    $config = Get-ADMigrationConfig
    $LocalExport = $config.ExportRoot
    $LocalTransform = $config.TransformRoot
    
    # Check if we need to sync
    $shouldSync = $false
    
    if (-not (Test-Path $LocalExport)) {
        $shouldSync = $true
        Write-Host "Portable data detected. Initializing local environment..." -ForegroundColor Cyan
    } else {
        # Data exists in both places. Ask user.
        $msg = "Portable migration data detected in the script directory.`n`nDo you want to import this data to your local workspace?`n`nSource: $PortableExport`nTarget: $LocalExport`n`nClick YES to overwrite local data.`nClick NO to use existing local data."
        $result = [System.Windows.Forms.MessageBox]::Show($msg, "Portable Data Detected", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -eq 'Yes') { $shouldSync = $true }
    }
    
    if ($shouldSync) {
        Write-Host "Syncing data to workspace..." -ForegroundColor Yellow
        if (-not (Test-Path $LocalExport)) { New-Item -ItemType Directory -Path $LocalExport -Force | Out-Null }
        if (-not (Test-Path $LocalTransform)) { New-Item -ItemType Directory -Path $LocalTransform -Force | Out-Null }
        
        Copy-Item -Path "$PortableExport\*" -Destination $LocalExport -Recurse -Force
        if (Test-Path $PortableTransform) {
            Copy-Item -Path "$PortableTransform\*" -Destination $LocalTransform -Recurse -Force
        }
        Write-Host "Data sync complete." -ForegroundColor Green
    }
}

# --- Helper Function for User Prompts ---
function Show-ManualActionPrompt {
    param(
        [string]$Title,
        [string]$Message,
        [string]$ToolPath = $null,
        [string]$ToolButtonText = "Launch Tool"
    )
    
    if (-not $ToolPath) {
        $result = [System.Windows.Forms.MessageBox]::Show($Message, $Title, [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -eq 'Cancel') {
            Write-Host "Transform process cancelled by user." -ForegroundColor Yellow
            throw "User cancelled operation."
        }
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(480, 260)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $lblMessage = New-Object System.Windows.Forms.Label
    $lblMessage.Text = $Message
    $lblMessage.Location = New-Object System.Drawing.Point(20, 20)
    $lblMessage.Size = New-Object System.Drawing.Size(420, 100)
    $form.Controls.Add($lblMessage)

    $btnLaunch = New-Object System.Windows.Forms.Button
    $btnLaunch.Text = $ToolButtonText
    $btnLaunch.Location = New-Object System.Drawing.Point(20, 130)
    $btnLaunch.Size = New-Object System.Drawing.Size(150, 30)
    $btnLaunch.Add_Click({
        if (Test-Path $ToolPath) {
            Start-Process powershell.exe -ArgumentList "-Command", "& '$ToolPath'"
        } else {
            [System.Windows.Forms.MessageBox]::Show("Script not found at: $ToolPath", "Error", "OK", "Error")
        }
    })
    $form.Controls.Add($btnLaunch)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK (Done)"
    $btnOK.Location = New-Object System.Drawing.Point(220, 170)
    $btnOK.Size = New-Object System.Drawing.Size(100, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(340, 170)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)

    $form.AcceptButton = $btnOK
    $form.CancelButton = $btnCancel

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Host "Transform process cancelled by user." -ForegroundColor Yellow
        throw "User cancelled operation."
    }
}

function Invoke-TransformStep {
    param(
        [string]$ScriptName
    )
    $scriptPath = Join-Path $ScriptRoot $ScriptName
    if (Test-Path $scriptPath) {
        Write-Host "`n[+] Running $ScriptName..." -ForegroundColor Green
        & $scriptPath
    } else {
        throw "Script not found: $scriptPath"
    }
}

# --- Main Transform Workflow ---
try {
    Write-Host "=== Starting Interactive Transform Phase ===" -ForegroundColor Cyan
    [System.Windows.Forms.MessageBox]::Show("This wizard will guide you through the Transform phase. This phase requires manual input to map your source environment to the target.`n`nClick OK to begin.", "Transform Phase Wizard", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    # 1. Validate Exports
    Write-Host "`n--- Step 1: Validate Exported Data ---" -ForegroundColor Yellow
    Invoke-TransformStep -ScriptName "Validation-Exports.ps1"
    Show-ManualActionPrompt -Title "Exports Validated" -Message "Export validation complete. Check the output above for any warnings.`n`nClick OK to proceed to OU Mapping."

    # 2. OU Mapping
    Write-Host "`n--- Step 2: OU Mapping ---" -ForegroundColor Yellow
    Invoke-TransformStep -ScriptName "Transform-OUMap.ps1"
    $guiMapperPath = Join-Path $ScriptRoot "Transform-OUMap-GUI.ps1"
    Show-ManualActionPrompt -Title "Manual Action: Edit OU Map" -Message "The script has created 'OU_Map_Draft.csv'.`n`nPlease edit this file to define your target OU structure. You can use the GUI mapper (`.\Scripts\Transform\Transform-OUMap-GUI.ps1`) for this.`n`nClick OK when you have finished editing and saved the file." -ToolPath $guiMapperPath -ToolButtonText "Launch GUI Mapper"
    Invoke-TransformStep -ScriptName "Validation-OUMap.ps1"
    Show-ManualActionPrompt -Title "OU Map Validated" -Message "OU Map validation complete. Check the output above for any errors.`n`nClick OK to proceed to Account Mapping."

    # 3. Account Mapping
    Write-Host "`n--- Step 3: Account Mapping ---" -ForegroundColor Yellow
    Invoke-TransformStep -ScriptName "Transform-AccountMapping.ps1"
    Show-ManualActionPrompt -Title "Manual Action: Edit Account Map" -Message "Account mapping files have been generated.`n`nPlease review and edit these files to define how accounts will be migrated and placed in the new OUs.`n`nClick OK when you are finished."
    Invoke-TransformStep -ScriptName "Validation-AccountPlacement.ps1"
    Show-ManualActionPrompt -Title "Account Placement Validated" -Message "Account placement validation complete. Check the output above for any errors.`n`nClick OK to proceed."

    # 4. WMI, GPO, DNS Transforms
    Write-Host "`n--- Step 4: WMI, GPO, and DNS Transforms ---" -ForegroundColor Yellow
    Invoke-TransformStep -ScriptName "Transform-WMIFilters.ps1"
    Invoke-TransformStep -ScriptName "Transform-GPOSettings.ps1"
    Invoke-TransformStep -ScriptName "Transform-DNS.ps1"
    Show-ManualActionPrompt -Title "Automated Transforms Complete" -Message "WMI, GPO, and DNS transforms are complete.`n`nClick OK to generate the final migration summary."

    # 5. Final Migration Table
    Write-Host "`n--- Step 5: Generate Final Migration Table ---" -ForegroundColor Yellow
    Invoke-TransformStep -ScriptName "Transform-GenerateMigrationTable.ps1"

    [System.Windows.Forms.MessageBox]::Show("The Transform phase is complete. All necessary mapping and rebuild files have been generated in the 'Transform' directory.`n`nYou can now proceed to the Import phase.", "Transform Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    Write-Host "`n=== Transform Sequence Complete ===" -ForegroundColor Cyan

} catch {
    Write-Host "`n=== Transform Sequence Failed or Cancelled ===" -ForegroundColor Red
    Write-Log -Message "Transform orchestrator failed or was cancelled by user: $_" -Level ERROR
}