# AD-Migration Copilot Instructions

You are an expert PowerShell developer assisting with an Active Directory Migration project.
The project follows a strict **Export -> Transform -> Import** pipeline to migrate from a Source Domain to a Target Domain without a trust relationship.

## 1. Project Architecture

### Phases
1.  **Export**: Read-only extraction from Source AD. Saves to `%USERPROFILE%\Documents\ADMigration\Export`.
2.  **Transform**: Offline manipulation of exported data (mapping OUs, rewriting GPOs). Saves to `%USERPROFILE%\Documents\ADMigration\Transform`.
3.  **Import**: Creation of objects in Target AD using Transformed data. Reads from `Transform` folder.

### Directory Structure
- `Scripts\ADMigration\`: The core module (`ADMigration.psd1`, `psm1`).
- `Scripts\Export\`: Scripts for Phase 1.
- `Scripts\Transform\`: Scripts for Phase 2.
- `Scripts\Import\`: Scripts for Phase 3.
- `Scripts\Validation\`: Scripts for Phase 4 (Verification).
- `Scripts\Rollback\`: Scripts for cleanup and reversal.
- `Docs\`: Documentation.

## 2. Coding Standards & Patterns

### Module Loading (Boilerplate)
All scripts must begin by dynamically locating and importing the `ADMigration` module. Do not hardcode paths.

```powershell
# Import module and config
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ModulePath = Join-Path (Join-Path (Split-Path -Parent (Split-Path -Parent $ScriptRoot)) 'Scripts') 'ADMigration\ADMigration.psd1'
if (-not (Test-Path $ModulePath)) { throw "Module manifest missing." }
Import-Module $ModulePath -Force
```

### Configuration & Paths
Never assume the location of data files. Use `Get-ADMigrationConfig`.

```powershell
$config = Get-ADMigrationConfig
$ExportPath = $config.ExportRoot
$TransformPath = $config.TransformRoot
# Example: $MapPath = Join-Path $config.TransformRoot 'Mapping'
```

### Logging
Use `Write-Log` for all status updates.
- Levels: `INFO`, `WARN`, `ERROR`, `DEBUG`.
- Logs are automatically written to `%USERPROFILE%\Documents\ADMigration\Logs\`.

```powershell
Write-Log -Message "Starting operation..." -Level INFO
```

### Error Handling & Safety
Wrap critical AD operations in `Invoke-Safely`.

```powershell
Invoke-Safely -ScriptBlock {
    # AD Operations here
    New-ADOrganizationalUnit ...
} -Operation "Create OUs"
```

### Input/Output
- **CSV**: Use `Import-Csv` and `Export-Csv -NoTypeInformation -Encoding UTF8`.
- **Validation**: Check if input files exist before processing. Throw clear errors if missing.

## 3. Specific Implementation Details

### OU Mapping
- OUs are mapped via CSV generated in Transform phase.
- Target DNs are constructed in `Transform-OUMap.ps1`.
- Import scripts must respect the `Action` column (Migrate, Skip, Merge).

### GPO Migration
- GPOs are backed up using `Backup-GPO`.
- Reports (XML) are used for analysis.
- Import involves `Import-GPO` (from backup) or `New-GPO` depending on strategy.
- Links are reconstructed separately after GPOs exist.

### WMI Filters
- WMI Filters are complex because they use GUIDs.
- We map them by Name.
- Queries must be sanitized (replace Source Domain references with Target Domain).

## 4. User Interaction
- If a critical decision is needed (e.g., skipping empty GPOs), use `System.Windows.Forms.MessageBox` or `Read-Host` if appropriate, but prefer automated defaults with logging.