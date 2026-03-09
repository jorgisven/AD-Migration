# AD-Migration

A PowerShell-based Active Directory migration framework for orchestrating clean, trustless migrations from a **source domain** to a **target domain**. This system is completely agnostic to domains, users, and systems—making it reusable across any migration scenario.

## 📋 Overview

This repository contains scripts, documentation, and mapping artifacts to execute a three-phase Active Directory migration pipeline, orchestrated by a central GUI:

```
Export (Source)  →  Transform (Workstation)  →  Import (Target)
(Extract Legacy)      (Map & Rewrite)          (Rebuild Target)
```

The goal is to rebuild the OU hierarchy, GPOs, WMI filters, and account structure cleanly—without establishing a trust—while preserving functionality and improving security boundaries.

## 📁 Repository Structure

```
AD-Migration/
|
|-- Scripts/                    # PowerShell module + migration scripts
|   |-- Export/                 # Export phase scripts
|   |   |-- Export-OUs.ps1
|   |   |-- Export-GPOReports.ps1
|   |   |-- Export-WMIFilters.ps1
|   |   |-- Export-AccountData.ps1
|   |   |-- Export-ACLs.ps1
|   |   |-- Export-DNS.ps1
|   |   |-- Export-MigrationPackage.ps1
|   |   `-- Run-AllExports.ps1
|   |
|   |-- Transform/              # Transform phase scripts
|   |   |-- Transform-OUMap.ps1
|   |   |-- Transform-OUMap-GUI.ps1
|   |   |-- Transform-AccountMapping.ps1
|   |   |-- Transform-GenerateMigrationTable.ps1
|   |   |-- Transform-GPOSettings.ps1
|   |   |-- Transform-WMIFilters.ps1
|   |   |-- Transform-ACLAnalysis.ps1
|   |   |-- Transform-DNS.ps1
|   |   |-- Validation-Exports.ps1
|   |   |-- Validation-OUMap.ps1
|   |   |-- Validation-AccountPlacement.ps1
|   |   |-- Validation-GPOApplication.ps1
|   |   `-- Run-AllTransforms.ps1
|   |
|   |-- Import/                 # Import phase scripts
|   |   |-- Import-OUs.ps1
|   |   |-- Import-GPOs.ps1
|   |   |-- Import-WMIFilters.ps1
|   |   |-- Import-GPOLinks.ps1
|   |   `-- Import-DNS.ps1
|   |
|   `-- ADMigration/            # PowerShell module (core utilities)
|       |-- ADMigration.psd1    # Module manifest
|       |-- ADMigration.psm1    # Module implementation
|       |-- Public/             # Exported functions
|       |   `-- Export-ADUsers.ps1
|       |
|       |-- Private/            # Internal helper functions
|       |   |-- Get-ADMigrationConfig.ps1
|       |   |-- Initialize-ADMigration.ps1
|       |   `-- Invoke-Safely.ps1
|       |
|       `-- Logging/
|           `-- Write-Log.ps1
|
|-- Docs/                       # Project documentation
|   |-- ProjectOutline.md       # Detailed project planning
|   |-- OU_Mapping.md           # OU mapping documentation
|   |-- AccountMapping.md       # Account reconciliation guide
|   |-- GPO_Notes.md            # GPO analysis and strategy
|   `-- ValidationChecklist.md  # Post-migration validation steps
|
|-- .github/
|   `-- copilot-instructions.md # AI assistant context and patterns
|
`-- README.md                   # This file
```

## 🧭 Migration Workstreams

### **1. Export Source Domain Data**
Extract all necessary information from the source domain:
- OU structure and attributes
- GPO backups (stored locally, not committed to GitHub)
- GPO XML reports (links, WMI filters, applied settings)
- WMI filter definitions and queries
- ACL and security information
- User and service account attributes for reconciliation

**Output:** CSV/XML files in `%USERPROFILE%\Documents\ADMigration\Export\`

### **2. Normalize & Map to Target Domain**
Analyze source data and prepare transformation rules:
- OU mapping and consolidation strategy
- GPO dependency analysis
- WMI filter rewrite planning
- Account reconciliation (merge, create, retire, disable)
- Naming and attribute standardization
- Domain-specific path and reference rewriting

**Output:** Mapping artifacts and rewrite rules in `%USERPROFILE%\Documents\ADMigration\Transform\`

### **3. Rebuild Target Domain Structure**
Pre-create the target domain hierarchy:
- Create OU hierarchy with consistent naming
- Apply tier/boundary structure (if applicable)
- Apply delegation at OU level
- Pre-stage computer accounts
- Create service accounts

**Execution:** Runs `Scripts/Import/Import-OUs.ps1` (Reads from `Transform\Mapping`)

### **4. Import & Reconstruct GPOs**
Migrate and validate Group Policy objects:
- Restore GPOs from backups
- Apply domain-specific setting rewrites
- Recreate WMI filters with target domain syntax
- Validate security filtering (if applicable)
- Document GPO dependencies and inheritance models

**Execution:** Runs `Scripts/Import/Import-GPOs.ps1` (Reads from `Export\GPO_Reports` & `Transform`)

### **5. Rebuild GPO Links**
Establish policies in the target domain:
- Recreate link locations (site, domain, OU)
- Apply link order and inheritance settings
- Re-enable enforcement flags
- Validate policy application (gpupdate, gpresult)
- Compare policy results vs. source domain

**Execution:** Runs `Scripts/Import/Import-GPOLinks.ps1` (Reads from `Export\GPO_Reports` & `Transform\Mapping`)

### **6. Validate Post-Migration**
Comprehensive testing to ensure functionality:
- ✅ GPO application on test machines
- ✅ ACL inheritance and permissions
- ✅ WMI filter evaluation
- ✅ User/service account authentication and permissions
- ✅ No privilege escalation or unintended access
- ✅ Naming consistency and standards compliance

**Output:** Validation report in logs and documentation

### **7. Pre-Cutover Preparation**
Final steps before production cutover:
- Freeze changes in source domain
- Finalize account mapping (last reconciliation)
- Prepare domain-join scripts and runbooks
- Prepare user profile migration strategy (if needed)
- Create rollback and recovery procedures
- Document known issues and workarounds

**Output:** Runbook and cutover checklists in `Docs/`

---

## 🔐 Sensitive Data Handling

This repository **deliberately excludes**:
- ❌ GPO backup files (stored locally only)
- ❌ SYSVOL-derived content
- ❌ Passwords, secrets, or credentials
- ❌ Raw domain controller logs
- ❌ Production AD objects or configurations

**Sensitive exports are stored locally** (not cloud-synced):
```
%USERPROFILE%\Documents\ADMigration\
├── Logs\
├── Export\
├── Transform\
└── Import\
```

For example:
- `C:\Users\<user>\Documents\ADMigration\`
- `D:\Projects\<user>_ADMigration\` (local disk, non-synced)

**GitHub repo contains only:**
- ✅ Scripts (reusable, non-sensitive)
- ✅ Documentation
- ✅ Mapping templates
- ✅ Non-sensitive export artifacts (CSV with non-sensitive attributes)

---

## 🛠 Requirements

- **Windows Server** with Active Directory Administrative Center or RSAT (or a workstation with RSAT enabled)
- **PowerShell 5.1+** (or PowerShell 7+ for cross-platform compatibility)
- **GPMC** (Group Policy Management Console) installed
- **AD Module** (part of RSAT) for `Get-ADUser`, `Get-GPO`, `New-ADOrganizationalUnit`, etc.
- Network access to both source and target domains (at appropriate phases)
- Appropriate permissions in source domain (read-only for exports)

> **Tip:** After cloning the repo and opening it in VS Code, run the helper
> `Initialize-ADMigration` from the module. It will create the required
> `%USERPROFILE%\Documents\ADMigration` folders and warn if prerequisites
> such as the ActiveDirectory module or GPMC cmdlets are missing. Installing
> RSAT on the machine will satisfy these dependencies; on Windows 10/11 you can
> enable it via **Settings → Apps → Optional features** or by running:
>
> ```powershell
> Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
> ```
>
> Once RSAT/GPMC are installed you should be able to import the module with:
>
> ```powershell
> Import-Module .\Scripts\ADMigration\ADMigration.psd1 -Force
> Get-Command -Module ADMigration
> ```
>
> and proceed with exports/transforms/imports as described below.
- Object Creator permissions in target domain (write for imports)

---

## 🚀 Quick Start

### 1. Clone the Repository
```powershell
git clone https://github.com/jorgisven/AD-Migration.git
cd AD-Migration
```

### 2. Initialize the Module
```powershell
Import-Module .\Scripts\ADMigration\ADMigration.psd1 -Force
Initialize-ADMigration
```

This creates the local data directory at `%USERPROFILE%\Documents\ADMigration\` with Logs, Export, Transform, and Import subdirectories.

### 3. Run Export Phase
```powershell
# Run the orchestration script to export all data
.\Scripts\Export\Run-AllExports.ps1 -SourceDomain "source.local"
```

### 4. Review & Process Transforms
Analyze the exported data in `%USERPROFILE%\Documents\ADMigration\Export\` and create mapping artifacts in the Transform phase folders.

### 5. Run Import Phase
```powershell
# Create OU hierarchy
.\Scripts\Import\Import-OUs.ps1

# Restore GPOs
.\Scripts\Import\Import-GPOs.ps1

# Recreate WMI filters
.\Scripts\Import\Import-WMIFilters.ps1

# Rebuild GPO links
.\Scripts\Import\Import-GPOLinks.ps1
```

### 6. Validate
Check logs at `%USERPROFILE%\Documents\ADMigration\Logs\` and run validation tests.

---

## 📖 Documentation

For detailed guidance on specific topics:
- **[Copilot Instructions](.github/copilot-instructions.md)** - Architecture, patterns, and developer workflows
- **[Docs/ProjectOutline.md](Docs/ProjectOutline.md)** - Detailed project planning and scope
- **[Docs/OU_Mapping.md](Docs/OU_Mapping.md)** - OU consolidation and mapping strategy
- **[Docs/AccountMapping.md](Docs/AccountMapping.md)** - Account reconciliation guide
- **[Docs/GPO_Notes.md](Docs/GPO_Notes.md)** - GPO analysis and rewrite strategy
- **[Docs/ValidationChecklist.md](Docs/ValidationChecklist.md)** - Post-migration validation steps

---

## 🔧 Key Patterns & Conventions

### Logging
All scripts log to `%USERPROFILE%\Documents\ADMigration\Logs\YYYY-MM-DD.log`:
```powershell
Write-Log -Message "OU exported successfully" -Level INFO
Write-Log -Message "Failed to export GPO" -Level ERROR
```

### Error Handling
All AD operations use safe invocation:
```powershell
Invoke-Safely -ScriptBlock { Get-ADOrganizationalUnit -Filter * } -Operation "Retrieve OUs"
```

### Path Resolution
Never hardcode paths—use `Get-ADMigrationConfig`:
```powershell
$config = Get-ADMigrationConfig
$ExportPath = Join-Path $config.ExportRoot 'OU_Structure'
```

---

## 📊 Status

**Current Phase:** ⏳ Development (Phase 1: Export scripts are templates)

- [x] Folder structure established
- [x] PowerShell module framework created
- [x] Logging and error handling utilities
- [ ] Export scripts (in progress)
- [ ] Transform scripts (planned)
- [ ] Import scripts (planned)
- [ ] Validation scripts (planned)
- [ ] Documentation (in progress)

---

## 👤 Author

**GitHub:** [jorgisven](https://github.com/jorgisven)

---

## 📝 Notes

This repository is designed to support a **clean, trustless migration** where the target domain is built intentionally rather than cloned blindly. The structure emphasizes:

- **Auditability** – Track every change and decision
- **Repeatability** – Scripts can be re-run safely
- **Data Privacy** – Sensitive exports stay local, not synced to cloud
- **Reusability** – Works across any source/target domain pair
- **Documentation** – Every step is documented for future reference

---

## ⚠️ Disclaimer

This tool is provided as-is for Active Directory migrations. Thoroughly test all scripts in a non-production environment first. Always maintain backups of source and target domains before executing import operations.
