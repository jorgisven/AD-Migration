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
        $msg = "Portable migration data detected in the script directory.`n`nDo you want to import this data to your local workspace?`n`nSource: $PortableExport`nTarget: $LocalExport`n`nClick YES to overwrite the files in your local workspace folder.`nClick NO to keep your existing workspace files.`n`n(Note: This only affects the migration files in your Documents folder, NOT your Active Directory domain.)"
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
    $form.Size = New-Object System.Drawing.Size(480, 220)
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
    $btnOK.Location = New-Object System.Drawing.Point(230, 130)
    $btnOK.Size = New-Object System.Drawing.Size(100, 30)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(340, 130)
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
    $exportsValidated = $false
    $exportsValidationBypassed = $false
    $exportsValidationAttempt = 0
    while (-not $exportsValidated) {
        $exportsValidationAttempt++
        Write-Log -Message "Export validation attempt #$exportsValidationAttempt started." -Level INFO
        try {
            Invoke-TransformStep -ScriptName "Validation-Exports.ps1"
            $exportsValidated = $true
            Write-Log -Message "Export validation attempt #$exportsValidationAttempt succeeded." -Level INFO
        } catch {
            $errText = $_.Exception.Message
            Write-Log -Message "Export validation attempt #$exportsValidationAttempt failed: $errText" -Level ERROR
            $decision = [System.Windows.Forms.MessageBox]::Show("Export validation failed.`n`nError: $errText`n`nYES = Retry validation`nNO = Continue Anyway`nCancel = Abort", "Export Validation Failed", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($decision -eq [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log -Message "Export validation retry requested by user after attempt #$exportsValidationAttempt." -Level WARN
            } elseif ($decision -eq [System.Windows.Forms.DialogResult]::No) {
                $exportsValidated = $true
                $exportsValidationBypassed = $true
                Write-Log -Message "User selected Continue Anyway after failed export validation attempt #$exportsValidationAttempt." -Level WARN
            } else {
                Write-Log -Message "User cancelled during export validation after failed attempt #$exportsValidationAttempt." -Level WARN
                throw "User cancelled operation."
            }
        }
    }

    if ($exportsValidationBypassed) {
        Show-ManualActionPrompt -Title "Exports Validation Bypassed" -Message "Export validation was bypassed by user selection (Continue Anyway).`n`nProceed with caution and review logs for details before import."
    } else {
        Show-ManualActionPrompt -Title "Exports Validated" -Message "Export validation complete. Check the output above for any warnings.`n`nClick OK to proceed to OU Mapping."
    }

    # 2. OU Mapping
    Write-Host "`n--- Step 2: OU Mapping ---" -ForegroundColor Yellow
    Invoke-TransformStep -ScriptName "Transform-OUMap.ps1"
    $guiMapperPath = Join-Path $ScriptRoot "Transform-OUMap-GUI.ps1"

    $ouValid = $false
    $ouValidationBypassed = $false
    $ouValidationAttempt = 0
    $ouMsg = "The script has created 'OU_Map_Draft.csv'.`n`nPlease edit this file to define your target OU structure. You can use the GUI mapper for this.`n`nClick OK when you have finished editing and saved the file."
    while (-not $ouValid) {
        Show-ManualActionPrompt -Title "Manual Action: Edit OU Map" -Message $ouMsg -ToolPath $guiMapperPath -ToolButtonText "Launch GUI Mapper"
        $ouValidationAttempt++
        Write-Log -Message "OU Map validation attempt #$ouValidationAttempt started." -Level INFO
        try {
            Invoke-TransformStep -ScriptName "Validation-OUMap.ps1"
            $ouValid = $true
            Write-Log -Message "OU Map validation attempt #$ouValidationAttempt succeeded." -Level INFO
        } catch {
            $errText = $_.Exception.Message
            Write-Log -Message "OU Map validation attempt #$ouValidationAttempt failed: $errText" -Level ERROR
            [System.Windows.Forms.MessageBox]::Show("Validation failed! Please check the console behind this window for details.", "Validation Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            $decision = [System.Windows.Forms.MessageBox]::Show("OU Map validation failed.`n`nError: $errText`n`nYES = Retry after fixing`nNO = Continue Anyway`nCancel = Abort", "OU Map Validation", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($decision -eq [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log -Message "OU Map validation retry requested by user after attempt #$ouValidationAttempt." -Level WARN
            } elseif ($decision -eq [System.Windows.Forms.DialogResult]::No) {
                $ouValid = $true
                $ouValidationBypassed = $true
                Write-Log -Message "User selected Continue Anyway after failed OU Map validation attempt #$ouValidationAttempt." -Level WARN
            } else {
                Write-Log -Message "User cancelled during OU Map validation after failed attempt #$ouValidationAttempt." -Level WARN
                throw "User cancelled operation."
            }
            $ouMsg = "VALIDATION FAILED! Check the console for errors.`n`nClick 'Launch GUI Mapper' to fix the issues in your mapping, save it, and click OK to try again."
        }
    }
    if ($ouValidationBypassed) {
        [System.Windows.Forms.MessageBox]::Show("OU Map validation was bypassed (Continue Anyway). Please review logs and proceed carefully.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    } else {
        [System.Windows.Forms.MessageBox]::Show("OU Map validation passed!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }

    # 3. Account Mapping
    Write-Host "`n--- Step 3: Account Mapping ---" -ForegroundColor Yellow
    Invoke-TransformStep -ScriptName "Transform-AccountMapping.ps1"
    $guiAccountMapperPath = Join-Path $ScriptRoot "Transform-AccountMap-GUI.ps1"

    $accValid = $false
    $accValidationBypassed = $false
    $accValidationAttempt = 0
    $accMsg = "Account mapping files have been generated.`n`nPlease review and edit these files to define how accounts will be migrated and placed in the new OUs.`n`nClick OK when you are finished."
    while (-not $accValid) {
        Show-ManualActionPrompt -Title "Manual Action: Edit Account Map" -Message $accMsg -ToolPath $guiAccountMapperPath -ToolButtonText "Launch GUI Mapper"
        $accValidationAttempt++
        Write-Log -Message "Account Placement validation attempt #$accValidationAttempt started." -Level INFO
        try {
            Invoke-TransformStep -ScriptName "Validation-AccountPlacement.ps1"
            $accValid = $true
            Write-Log -Message "Account Placement validation attempt #$accValidationAttempt succeeded." -Level INFO
        } catch {
            $errText = $_.Exception.Message
            Write-Log -Message "Account Placement validation attempt #$accValidationAttempt failed: $errText" -Level ERROR
            [System.Windows.Forms.MessageBox]::Show("Validation failed! Please check the console behind this window for details.", "Validation Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            $decision = [System.Windows.Forms.MessageBox]::Show("Account Placement validation failed.`n`nError: $errText`n`nYES = Retry after fixing`nNO = Continue Anyway`nCancel = Abort", "Account Validation", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($decision -eq [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log -Message "Account Placement validation retry requested by user after attempt #$accValidationAttempt." -Level WARN
            } elseif ($decision -eq [System.Windows.Forms.DialogResult]::No) {
                $accValid = $true
                $accValidationBypassed = $true
                Write-Log -Message "User selected Continue Anyway after failed Account Placement validation attempt #$accValidationAttempt." -Level WARN
            } else {
                Write-Log -Message "User cancelled during Account Placement validation after failed attempt #$accValidationAttempt." -Level WARN
                throw "User cancelled operation."
            }
            $accMsg = "VALIDATION FAILED! Check the console for errors.`n`nClick 'Launch GUI Mapper' to fix the 'TargetOU_DN' entries in your account mappings, save them, and click OK to try again."
        }
    }
    if ($accValidationBypassed) {
        [System.Windows.Forms.MessageBox]::Show("Account placement validation was bypassed (Continue Anyway). Please review logs and proceed carefully.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    } else {
        [System.Windows.Forms.MessageBox]::Show("Account placement validation passed!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }
    # 4. WMI, GPO, DNS Transforms
    Write-Host "`n--- Step 4: WMI, GPO, and DNS Transforms ---" -ForegroundColor Yellow
    Invoke-TransformStep -ScriptName "Transform-WMIFilters.ps1"
    Invoke-TransformStep -ScriptName "Transform-GPOSettings.ps1"
    Invoke-TransformStep -ScriptName "Transform-DNS.ps1"
    Write-Log -Message "Automated WMI, GPO, and DNS transforms completed successfully." -Level INFO
    Show-ManualActionPrompt -Title "Automated Transforms Complete" -Message "WMI, GPO, and DNS transforms are complete.`n`nClick OK to generate the final migration summary."

    # Optional: Policy Analyzer Integration
    $analyzerMsg = @"
Would you like to run the optional GPO conflict review before proceeding?

This step launches Validation-GPOConflicts.ps1, which opens Microsoft Policy Analyzer and loads your exported GPO backups for visual comparison.

If Policy Analyzer is not installed, you will be prompted to locate PolicyAnalyzer.exe (download from https://aka.ms/sct).

Click YES to launch Policy Analyzer via Validation-GPOConflicts.ps1.
Click NO to skip this check and continue to the next step.
"@
    $analyzerResult = [System.Windows.Forms.MessageBox]::Show($analyzerMsg, "Check GPO Conflicts", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($analyzerResult -eq 'Yes') {
        Invoke-TransformStep -ScriptName "Validation-GPOConflicts.ps1"
        Write-Log -Message "GPO Conflict Validation script executed." -Level INFO
    } else {
        Write-Log -Message "GPO Conflict Validation skipped by user." -Level INFO
    }

    # 5. Final Migration Table
    Write-Host "`n--- Step 5: Generate Final Migration Table ---" -ForegroundColor Yellow
    Invoke-TransformStep -ScriptName "Transform-GenerateMigrationTable.ps1"
    $guiMigTablePath = Join-Path $ScriptRoot "Transform-MigTable-GUI.ps1"

    $migValid = $false
    $migValidationBypassed = $false
    $migValidationAttempt = 0
    $migMsg = "A GPO Migration Table (.migtable) has been generated based on your account mappings and GPO analysis.`n`nPlease review the table and fill in any blank Destination paths or accounts.`n`nClick OK when you are finished."
    while (-not $migValid) {
        Show-ManualActionPrompt -Title "Manual Action: Edit Migration Table" -Message $migMsg -ToolPath $guiMigTablePath -ToolButtonText "Launch GUI Editor"
        $migValidationAttempt++
        Write-Log -Message "Migration Table validation attempt #$migValidationAttempt started." -Level INFO
        try {
            Invoke-TransformStep -ScriptName "Validation-MigTable.ps1"
            $migValid = $true
            Write-Log -Message "Migration Table validation attempt #$migValidationAttempt succeeded." -Level INFO
        } catch {
            $errText = $_.Exception.Message
            Write-Log -Message "Migration Table validation attempt #$migValidationAttempt failed: $errText" -Level ERROR
            [System.Windows.Forms.MessageBox]::Show("Validation failed! Please check the console behind this window for details.", "Validation Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
            $decision = [System.Windows.Forms.MessageBox]::Show("Migration Table validation failed.`n`nError: $errText`n`nYES = Retry after fixing`nNO = Continue Anyway`nCancel = Abort", "Migration Table Validation", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($decision -eq [System.Windows.Forms.DialogResult]::Yes) {
                Write-Log -Message "Migration Table validation retry requested by user after attempt #$migValidationAttempt." -Level WARN
            } elseif ($decision -eq [System.Windows.Forms.DialogResult]::No) {
                $migValid = $true
                $migValidationBypassed = $true
                Write-Log -Message "User selected Continue Anyway after failed Migration Table validation attempt #$migValidationAttempt." -Level WARN
            } else {
                Write-Log -Message "User cancelled during Migration Table validation after failed attempt #$migValidationAttempt." -Level WARN
                throw "User cancelled operation."
            }
            $migMsg = "VALIDATION FAILED! Check the console for errors.`n`nEvery item in the Migration Table must have a Destination mapped. Click 'Launch GUI Editor' to fix the empty fields, save, and click OK to try again."
        }
    }
    if ($migValidationBypassed) {
        [System.Windows.Forms.MessageBox]::Show("Migration Table validation was bypassed (Continue Anyway). Please review logs and proceed carefully.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    } else {
        [System.Windows.Forms.MessageBox]::Show("Migration Table validation passed!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
    }

    [System.Windows.Forms.MessageBox]::Show("The Transform phase is complete. All necessary mapping and rebuild files have been generated in the 'Transform' directory.`n`nYou can now proceed to the Import phase.", "Transform Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    Write-Host "`n=== Transform Sequence Complete ===" -ForegroundColor Cyan
    Write-Log -Message "Transform Sequence Complete." -Level INFO

} catch {
    Write-Host "`n=== Transform Sequence Failed or Cancelled ===" -ForegroundColor Red
    Write-Log -Message "Transform orchestrator failed or was cancelled by user: $_" -Level ERROR
}