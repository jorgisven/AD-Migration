# AD-Migration Codebase Instructions

## Project Overview
**SAICPRINT → CRIT.AD Migration**: A PowerShell-based Active Directory migration system that orchestrates a three-phase pipeline: **Export** (legacy domain) → **Transform** (mapping & rewrites) → **Import** (target domain).

## Architecture & Critical Workflows

### Three-Phase Pipeline Architecture
```
01_SAICPRINT_Exports  →  02_Transform  →  03_CRITAD_Imports
(GPO_Backups,OU_Struture, (ACL_Analysis,  (GPO_Restores,
 Security,WMI_Filters)  Mapping,WMI_Rebuild) Link_Rebuild)
```
- **Export Phase**: `Scripts/Export/*.ps1` pull data from source domain (saicprint.local)
- **Transform Phase**: `Scripts/Transform/*.ps1` normalize/rewrite data for target domain compatibility
- **Import Phase**: `Scripts/Import/*.ps1` rebuild structure in target domain (crit.ad)

### Path Strategy: Code vs. Data
- **GitHub Repo** (this codebase): Can be anywhere, syncs with GitHub (e.g., `C:\Users\jorgi\OneDrive\Documents\GitHub\AD-Migration`)
- **Sensitive Data** (logs, exports, transforms, imports): Always saved to local Documents folder **without OneDrive**: `C:\Users\<username>\Documents\ADMigration\{Logs,Export,Transform,Import}`
- **Why separate?**: Sensitive AD data (users, ACLs, GPO configs) stays local; code is version-controlled in GitHub

## Key Patterns & Conventions

### Logging Pattern (Required for all scripts)
```powershell
Write-Log -Message "operation completed" -Level INFO
Write-Log -Message "connection failed" -Level ERROR
# Logs to: C:\Users\<username>\Documents\ADMigration\Logs\YYYY-MM-DD.log
# Levels: INFO, WARN, ERROR
```

### Error Handling Pattern
```powershell
Invoke-Safely -ScriptBlock { /* operation */ } -Operation "descriptive name"
# Wraps with try-catch, logs errors, re-throws
```

### File Organization Rules
- **Public Functions**: `Scripts/ADMigration/Public/*.ps1` - exported via PSD1
- **Private Helpers**: `Scripts/ADMigration/Private/*.ps1` - internal utilities only
- **Phase Scripts**: Follow naming: `{Export|Transform|Import}-{Domain|GPO|OU|WMIFilters|ACL}.ps1`
- Each script has `.SYNOPSIS` and `.DESCRIPTION` blocks at top

### Export Script Pattern (all TODO - incomplete)
```powershell
# Template in Scripts/Export/Export-*.ps1
$ExportPath = ".\01_SAICPRINT_Exports\{GPO_Backups|OU_Structure|...}"
# Return CSV/XML to workspace folder, use Get-ADMigrationConfig for paths
```

## Critical Developer Workflows

### Adding a New Migration Function
1. Create in appropriate phase folder: `Scripts/{ADMigration|Export|Transform|Import}/`
2. If helper function → place in `Scripts/ADMigration/Private/`
3. Add `Write-Log` calls for progress tracking
4. Wrap AD operations in `Invoke-Safely -Operation "description"`
5. Export from `ADMigration.psd1` if public-facing
6. Use `Get-ADMigrationConfig` to reference base paths (don't hardcode paths)

### Running/Testing Scripts
- Module auto-loads on import: `Import-Module .\Scripts\ADMigration\ADMigration.psd1`
- All paths resolve via `Get-ADMigrationConfig` to `C:\Users\<username>\Documents\ADMigration`
- Check logs at `C:\Users\<username>\Documents\ADMigration\Logs\{today}.log`

## Dependencies & Integration Points

### Active Directory Integration
- Requires **AD Module** (part of RSAT) for `Get-ADUser`, `Get-GPO`, `New-ADOrganizationalUnit`, etc.
- Source domain: `saicprint.local` (export phase)
- Target domain: `crit.ad` (import phase)
- No explicit connection logic yet—assumes domain trusts or agent runs with appropriate permissions

### WMI Filters & Group Policy
- GPO backups: XML format in `01_SAICPRINT_Exports/GPO_Backups/`
- WMI filters: Stored separately in `01_SAICPRINT_Exports/WMI_Filters/` (requires rewrite for new domain)
- GPO links: Rebuilt by `Import-GPOLinks.ps1` after GPOs imported

### CSV/Data Flow
- Export outputs → CSV files in phase folders
- Transform processes CSVs + rewrites → prepared CSVs in phase 2
- Import reads prepared CSVs → creates AD objects

## Important Implementation Notes

- **Path Strategy**: GitHub repo location is flexible; sensitive data always goes to `C:\Users\<username>\Documents\ADMigration` via `Get-ADMigrationConfig`
- **ACL Analysis**: `Transform-ACLAnalysis.ps1` compares source vs target ACLs (incomplete—add diff logic)
- **All major scripts are TODO placeholders** — implement export/transform/import logic incrementally
- **Module Version**: Currently 1.0.0; update PSD1 when adding breaking changes

## When Adding/Modifying Code
- Start with `Get-ADMigrationConfig` to understand path structure
- Reference `Write-Log` and `Invoke-Safely` for error handling pattern
- Check `02_Transform/` and `01_SAICPRINT_Exports/` folder structure to understand expected data formats
- All public functions must appear in `ADMigration.psd1` FunctionsToExport array
- Test module loading: `Import-Module .\Scripts\ADMigration\ADMigration.psd1 -Force`
