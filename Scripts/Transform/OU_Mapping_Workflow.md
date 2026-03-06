# OU Mapping Workflow

## Overview
The OU Mapping phase is the architectural blueprint for your new Active Directory. This document outlines how to use the `Transform-OUMap.ps1` script and the resulting CSV to define your target structure.

## Workflow Steps

### 1. Generate the Draft Map
Run the transform script to analyze your exported source OUs and generate a baseline CSV.

```powershell
.\Scripts\Transform\Transform-OUMap.ps1
```

This creates: `%USERPROFILE%\Documents\ADMigration\Transform\Mapping\OU_Map_Draft.csv`

### 2. Edit the Map
You have two options for editing the map:

#### Option A: GUI Mapper (Recommended)
Use the visual drag-and-drop tool to design your target structure.

```powershell
.\Scripts\Transform\Transform-OUMap-GUI.ps1
```
1. **Left Pane**: Shows your Source OUs.
2. **Right Pane**: Shows your Target Domain.
3. **Drag & Drop**: Move OUs from Source to Target.
4. **Right-Click**: Rename or create new OUs in the Target pane.
5. **Save**: Writes to `OU_Map_Draft.csv`.

#### Option B: Manual CSV Editing
Open `OU_Map_Draft.csv` in Excel or a text editor.

- **Action**:
  - `Migrate`: (Default) Create this OU.
  - `Skip`: Do not create this OU. Useful for pruning old structures.
  - `Merge`: Map this source OU to an existing target OU.
- **SourceOU**: The name of the OU in the source (for reference).
- **TargetDN**: **(Critical)** The Distinguished Name in the Target Domain.
  - Default: `OU=<Name>,OU=Migrated,DC=Target...`
  - **Edit this** to restructure. e.g., change `OU=Users,OU=Migrated...` to `OU=Users,OU=Corp,DC=Target...`
- **SourceDN**: The unique identifier from the source. **Do not modify**.

### 3. Save and Validate
Save the CSV. Ensure the file encoding remains UTF-8 if possible.

### 4. Import
When `Scripts\Import\Import-OUs.ps1` runs, it reads this map. It will:
1. Filter out `Skip` rows.
2. Create OUs defined in `TargetDN`.
3. Apply descriptions.
