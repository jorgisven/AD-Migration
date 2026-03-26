# Group Policy Migration Strategy

## Overview

## The Process

### 1. Backup (Export)
We use `Backup-GPO` to create a portable backup of all GPOs.
- **Location**: `Export\GPO_Backups\`
- **Manifest**: `manifest.xml` maps the GUIDs to friendly names.

### 2. Analysis (Transform)
We generate XML reports (`Export-GPOReports.ps1`) and analyze them (`Transform-GPOSettings.ps1`) to find:
- **Hardcoded UNC Paths**: `\\source-dc\share` will break in the target.
- **Security Principals**: `SOURCE\Domain Admins` in "User Rights Assignment" needs to become `TARGET\Domain Admins`.

### 3. Migration Table
The script `Transform-GenerateMigrationTable.ps1` creates a `.migtable` file.
- This maps Source SIDs -> Target SIDs.
- This maps Source UNC Paths -> Target UNC Paths.
- **Manual Step**: You must review `MigrationTable.migtable` and fill in any blank destination fields before import.

### 4. Import
We use `Import-GPO` which reads the backup and applies it to a new GPO in the target.
- **Idempotency**: If the GPO exists, we skip it (unless `-Force` is used).
- **Migration Table**: Applied during import to rewrite security principals inside the policy settings.

### 5. Linking
Links are not part of the GPO object itself; they are attributes of the OU (SOM - Scope of Management).
- `Import-GPOLinks.ps1` reads the XML reports to find where GPOs were linked.
- It uses the **OU Map** to translate `OU=Sales,DC=source` to `OU=Sales,DC=target`.
- It restores:
    - Link Enabled status
    - Enforced (No Override) status
    - WMI Filter links

## WMI Filters
WMI Filters are separate objects.
- They are exported to `Export\WMI_Filters\`.
- `Import-WMIFilters.ps1` recreates them.
- **Note**: The query content is text. If your WMI query references the domain name (e.g., `SELECT * FROM Win32_NTDomain WHERE DomainName = 'SOURCE'`), you must manually update the CSV in the Transform phase.

## Detecting GPO Conflicts
The best way to detect internal setting conflicts (e.g., two GPOs trying to configure the same registry key with different values) across your migrated GPOs is to use the **Microsoft Policy Analyzer**.

1. Download the **Security Compliance Toolkit** from Microsoft: [https://aka.ms/sct](https://aka.ms/sct) (Select `PolicyAnalyzer.zip` from the download list).
2. Extract the ZIP file.
3. Run `PolicyAnalyzer.exe` and point it to your exported backups located at `Export\GPO_Backups\`.
4. Select the policies you wish to compare and click **View/Compare** to generate a conflict matrix. Any conflicting settings will be highlighted in Yellow.

## Known Limitations
- **Passwords in GPP**: Group Policy Preferences passwords (cPasswords) are deprecated and insecure. They are typically stripped by modern GPMC tools during backup/restore.
- **Software Installation**: MSI packages assigned via GPO rely on UNC paths. Ensure the software share is migrated and the Migration Table is updated with the new path.
- **WMI Filters**: WMI filter migration logic is included and handles basic domain name substitution, but it has not been exhaustively tested against highly complex environments. Always review the generated `WMI_Filters_Ready.csv` carefully before running the import phase.
```
