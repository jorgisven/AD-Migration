# Account Mapping & Reconciliation

## Overview
Migrating accounts (Users, Computers, Groups) in a trustless scenario requires recreating the objects with new SIDs. This document details how identities are preserved and mapped.

## Identity Map
The core artifact for account migration is `Identity_Map_Final.csv`. This file links the old identity to the new one and is used by GPO migration scripts to fix security filtering.

| SourceSam | SourceSID | TargetSam | TargetDN |
|-----------|-----------|-----------|----------|
| jdoe      | S-1-5-..  | jdoe      | CN=jdoe,OU=Users,DC=target... |

### Mapping Actions
When configuring the `User_Account_Map.csv` or `Computer_Account_Map.csv` during the Transform phase, you will define an `Action` for each source account:
- **`Create` (Move)**: The account is net-new. The import script will create it in the designated `TargetOU_DN`, generate a new password, and map the identity in the migration table.
- **`Merge`**: The user already exists in the target domain (Brownfield). The import script will **skip** creating or modifying the AD object, but will add the identity mapping to the `.migtable` so their old GPO/file permissions translate to their existing target account.
- **`Skip`**: The account is dead/disabled/obsolete and should not be brought over in any capacity.

## Handling Passwords
Since we cannot decrypt the NTLM hashes from the source domain without specialized tools (like Mimikatz, which we avoid for safety), **passwords cannot be migrated**.

This limitation aligns perfectly with the logical expectations of a "clean break" migration:
- **For "Moved" Users (Net-New to Target):** They are receiving a fundamentally new identity (e.g., `user@newdomain.com` instead of `user@olddomain.com`). It is standard practice and logically sound to issue a new initial password along with their new digital identity.
- **For "Merged" Users (Already exist in Target):** The user already has an active identity and password in the target domain. They simply continue using their existing target credentials for everything, and their old domain credentials are abandoned.

### Strategy
1. **Users**: All migrated users are created with a temporary initial password (e.g., defined in config or random).
2. **Flag**: Accounts are set to "User must change password at next logon".
3. **Communication**: Users must be notified of their new credentials via out-of-band methods.

## Handling SIDs (Security Identifiers)
In a trustless migration, **SID History is NOT used**.
- New objects get new SIDs.
- **Impact**: ACLs on file servers and GPO permissions referencing old SIDs will break.
- **Mitigation**: The `Import-GPOs` process uses a Migration Table to map old SIDs to new SIDs during import.

## High-Privilege Accounts
By default, the export scripts flag and exclude memberships for:
- Domain Admins
- Enterprise Admins
- Schema Admins
- Administrators

**Why?** To prevent accidental pollution of the new domain's administrative tier. You should manually re-assign administrative privileges in the target domain after import.

## Service Accounts
Service Accounts (identified by `svc_` prefix or similar) are exported to a separate list.
- **Managed Service Accounts (gMSA)**: Cannot be directly migrated. They must be re-created and re-installed on the target hosts.
- **Standard Service Accounts**: Migrated as disabled users. Passwords must be reset and updated on the services running them.

## Validation
Run `Validation-AccountPlacement.ps1` to ensure:
- Every account is mapped to a valid Target OU.
- No accounts are mapped to a `Skip`ped OU.