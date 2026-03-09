# Account Mapping & Reconciliation

## Overview
Migrating accounts (Users, Computers, Groups) in a trustless scenario requires recreating the objects with new SIDs. This document details how identities are preserved and mapped.

## Identity Map
The core artifact for account migration is `Identity_Map_Final.csv`. This file links the old identity to the new one and is used by GPO migration scripts to fix security filtering.

| SourceSam | SourceSID | TargetSam | TargetDN |
|-----------|-----------|-----------|----------|
| jdoe      | S-1-5-..  | jdoe      | CN=jdoe,OU=Users,DC=target... |

## Handling Passwords
Since we cannot decrypt the NTLM hashes from the source domain without specialized tools (like Mimikatz, which we avoid for safety), **passwords cannot be migrated**.

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