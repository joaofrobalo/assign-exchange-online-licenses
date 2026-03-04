# Assign Exchange Online Licenses

A PowerShell script that reads a list of email addresses from a CSV file and bulk-assigns a Microsoft 365 license to each user via Microsoft Graph.

## What it does

1. Verifies you are connected to Microsoft Graph
2. Reads email addresses from a CSV file
3. Fetches all active license SKUs in your tenant and lets you pick one interactively
4. For each user, sets `UsageLocation` to `PT` (Portugal) if it is not already set — this is required by Microsoft before a license can be assigned
5. Assigns the selected license to each user
6. Reports success/failure per user and saves a results CSV next to the input file

## Prerequisites

- PowerShell 5.1 or later
- Microsoft Graph PowerShell module:
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```
- Connected to Microsoft Graph before running the script:
  ```powershell
  Connect-MgGraph -Scopes "User.ReadWrite.All", "Organization.Read.All"
  ```

## CSV format

The input CSV must have a column named `EmailAddress`, `Email`, or `UserPrincipalName`. Values may optionally be wrapped in quotes.

```
EmailAddress
user1@contoso.com
user2@contoso.com
```

## Usage

```powershell
.\Assign-Licenses.ps1 -CsvPath "C:\path\to\users.csv"
```

The script will list available licenses in your tenant, prompt you to pick one, ask for confirmation, and then process each user.

A `_results.csv` file is saved alongside the input file with the outcome for each address.

## Notes

- License assignment is handled through Microsoft Graph, not the Exchange Online module. You do not need to be connected to EXO to run this script.
- The script only sets `UsageLocation` on users that are missing it. Users that already have a location set are not modified.
