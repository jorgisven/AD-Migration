# OU Mapping Strategy

## Overview
Organizational Unit (OU) mapping is the foundation of the migration. Since we are not simply cloning the domain, we have the opportunity to flatten deep hierarchies or consolidate disparate OUs.

## The Mapping File
The transformation process generates a file: `Transform\Mapping\OU_Map_Draft.csv`.

### Columns
- **Action**:
    - `Migrate`: Create this OU in the target.
    - `Skip`: Ignore this OU (and likely its contents, unless they are re-mapped).
    - `Merge`: (Advanced) Logic to merge contents into another OU.
- **SourceDN**: The DistinguishedName in the source domain (Reference only).
- **TargetDN**: The calculated DistinguishedName for the target domain.
- **Description**: Carried over from source.

## Workflow

1. **Export**: `Export-OUs.ps1` dumps the raw hierarchy.
2. **Draft**: `Transform-OUMap.ps1` creates the initial CSV, defaulting to a 1:1 map under a `Migrated` root OU.
3. **Design**: Use `Transform-OUMap-GUI.ps1` to visually drag-and-drop source OUs into the target structure.
    - *Tip*: You can load the live structure of the Target Domain to map legacy OUs into an existing brownfield environment.
4. **Import**: `Import-OUs.ps1` reads the CSV and builds the structure.

## Common Scenarios

### Scenario A: Lift and Shift
Keep the structure exactly the same.
- **Action**: Leave defaults.
- **TargetDN**: Ensure the root matches the new domain (e.g., `DC=target,DC=local`).

### Scenario B: Flattening
Moving `OU=HR,OU=Users,OU=London,DC=source` to `OU=HR,DC=target`.
- Edit the **TargetDN** column in the CSV.
- Change `OU=HR,OU=Users,OU=London,DC=target,DC=local` to `OU=HR,DC=target,DC=local`.

### Scenario C: Pruning
Removing empty or obsolete OUs.
- Set **Action** to `Skip`.
- *Note*: Ensure no active users or computers are mapped to a skipped OU in the Account Mapping phase, or the validation scripts will flag an error.

## Validation
Run `Validation-OUMap.ps1` to check for:
- Duplicate Target DNs.
- Orphaned OUs (an OU whose parent is marked as Skip or does not exist).