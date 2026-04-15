# SortedGroups

A PowerShell script that reads a server inventory Excel file and organizes servers into a hash table grouped by department/team. Built to help cloud administrators send bulk upgrade notification emails by group instead of server-by-server.

## What it does

- Parses an `.xlsx` server list via Excel COM object
- Builds a hash table keyed by group name
- Deduplicates contacts across `AppOwner`, `AppOwner_2`, and `Manager` columns
- Outputs a clean summary, email-ready `To:` lines, or raw JSON

## Requirements

- Windows with Microsoft Excel installed
- PowerShell 5.1+

## Usage

Place `SortedGroups.ps1` in the same directory as your Excel file, then run:

```powershell
# Full summary of all groups
.\SortedGroups.ps1

# Ready-to-paste email targets per group
.\SortedGroups.ps1 -Email

# Drill into a single group
.\SortedGroups.ps1 -Group "IT-DA"

# JSON dump of the full hash table
.\SortedGroups.ps1 -Json

# Custom file path
.\SortedGroups.ps1 -XlsxPath "C:\path\to\your\file.xlsx"
```

If you get a script execution error, run this first:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

## Expected Excel columns

| Column | Description |
|---|---|
| `Server` | Server hostname |
| `Provisioned Space` | Total allocated storage |
| `Used Space` | Current storage usage |
| `Host Mem` | Host memory |
| `Site` | Site/datacenter location |
| `Group` | Department or team name (used as hash table key) |
| `AppOwner` | Primary contact email |
| `AppOwner_2` | Secondary contact email (optional) |
| `Manager` | Manager email (optional) |
| `Status` | Current server status |

## Example output

```
============================================================
  GROUP : IT-DA  (6 server(s))
============================================================
  Contacts (3):
    - admin1@company.com
    - admin2@company.com
    - manager@company.com

  Servers:
    [MDF] server-01       Provisioned: 1.57 TB     Used: 1.48 TB
    [SQ]  server-02       Provisioned: 6.3 TB      Used: 3.82 TB
```
