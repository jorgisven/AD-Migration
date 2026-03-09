# Project Outline & Migration Strategy

## 1. Executive Summary
This project facilitates a **trustless Active Directory migration**. Unlike traditional migrations that rely on AD Trusts and SID History, this framework assumes the source and target domains are air-gapped or logically separated. It rebuilds the environment from the ground up using exported data, ensuring a clean, secure, and "greenfield-like" result.

## 2. Migration Philosophy
- **Clean Slate**: We do not migrate "garbage". Stale accounts, empty OUs, and unlinked GPOs are filtered out during the Transform phase.
- **Security First**: High-privilege groups (Domain Admins, etc.) are stripped during export/import to prevent privilege escalation in the new environment.
- **Auditability**: Every mapping decision (OU placement, account status) is recorded in CSV files before any write operation occurs in the target.

## 3. Roles & Responsibilities

### Source Administrator
- **Responsibility**: Run read-only export scripts.
- **Access**: Domain Admin (or Backup Operator) in Source.
- **Output**: A portable ZIP package containing CSVs and XML reports.

### Migration Analyst
- **Responsibility**: Map OUs, review GPO settings, and decide on account placement.
- **Access**: Workstation with Excel/PowerShell (No Domain Admin required).
- **Output**: Transformed CSV mapping files (`Transform\Mapping\`).

### Target Administrator
- **Responsibility**: Execute import scripts to build the new environment.
- **Access**: Domain Admin in Target.
- **Input**: The Transformed mapping files.

## 4. Phased Execution

### Phase 1: Discovery & Export
- **Goal**: Snapshot the current state.
- **Tools**: `Run-AllExports.ps1`
- **Artifacts**: Raw CSV dumps of Users, Groups, OUs, DNS, and GPO Backups.

### Phase 2: Transformation (The "Air Gap")
- **Goal**: Cleanse and restructure data.
- **Tools**: `Run-AllTransforms.ps1`, `Transform-OUMap-GUI.ps1`.
- **Actions**:
    - Map legacy OUs to new hierarchy.
    - Generate GPO Migration Tables (map old SIDs to new users).
    - Define User/Computer placement.

### Phase 3: Implementation
- **Goal**: Hydrate the target domain.
- **Tools**: `Run-AllImports.ps1`.
- **Sequence**:
    1. OUs (Skeleton)
    2. Accounts & Groups (Identity)
    3. GPOs & WMI Filters (Policy)
    4. DNS (Infrastructure)

### Phase 4: Validation
- **Goal**: Verify functionality.
- **Tools**: `Validation-*.ps1` scripts and manual checklists.