# Migration Validation Checklist

Use this checklist to verify the success of the migration phases.

## Phase 1: Export Validation
- [ ] **File Existence**: Run `Validation-Exports.ps1`. Ensure no critical files (Users, OUs, GPOs) are missing.
- [ ] **Size Check**: Ensure CSV files are not 0KB (indicating a failed export or empty domain).
- [ ] **GPO Backups**: Verify `Export\GPO_Backups` contains a `manifest.xml` and subfolders with GUIDs.

## Phase 2: Transform Validation
- [ ] **OU Map**: Run `Validation-OUMap.ps1`.
    - [ ] No duplicate Target DNs.
    - [ ] No orphaned OUs.
- [ ] **Account Placement**: Run `Validation-AccountPlacement.ps1`.
    - [ ] All users mapped to valid OUs.
- [ ] **Migration Table**: Open `Transform\Mapping\MigrationTable.migtable`.
    - [ ] Ensure all `<Destination>` fields are filled for SIDs and UNC paths you intend to keep.

## Phase 3: Import Validation (Target Domain)

### Structure
- [ ] **OUs**: Open ADUC. Verify the OU hierarchy matches the design.
- [ ] **Delegation**: Check if standard users can read/write where expected (if delegation was part of the scope).

### GPOs
- [ ] **Existence**: Verify all GPOs appear in GPMC.
- [ ] **Settings**: Open a complex GPO (e.g., "Default Domain Policy"). Check "User Rights Assignment".
    - [ ] Verify principals are `TARGET\Group` and not SIDs (S-1-5-...).
- [ ] **Links**: Click on an OU in GPMC. Verify the correct GPOs are linked and the Link Order is correct.
- [ ] **WMI Filters**: Verify WMI filters exist and are linked to the correct GPOs.

### Accounts
- [ ] **Users**: Verify user counts match the export (minus skipped accounts).
- [ ] **Groups**: Check a few groups (e.g., "HR_Users"). Verify members are present.
- [ ] **Logon**: Attempt to log on with a migrated test user.
    - [ ] Verify "User must change password" prompt appears.

### Infrastructure
- [ ] **DNS**: Verify Forward Lookup Zones are populated.
- [ ] **Replication**: Run `repadmin /showrepl` to ensure the new objects are replicating to other DCs in the target.

## Rollback Plan
If validation fails critically, delete the created OUs and GPOs in the target domain, adjust the Transform maps, and re-run the Import scripts.