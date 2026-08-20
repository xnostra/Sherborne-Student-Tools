# Sherborne Qatar Student Tools

Scripts for creating Office 365 student accounts for Sherborne Qatar, built on Microsoft Graph PowerShell.

## Quick Start

Run this single command in PowerShell on any computer:

```powershell
irm https://raw.githubusercontent.com/xnostra/Sherborne-Student-Tools/master/invoke-studenttoolkit.ps1 | iex
```

That's it - it downloads the toolkit scripts to `Desktop\StudentToolkit` and opens the GUI. Sign in with your Global Admin account when prompted.

## What's in the GUI

| Button | Mode | What it does |
|---|---|---|
| **BULK - Add New Students (XLSX)** | Bulk | Reads the school MIS export, creates missing student accounts, assigns the A5 for Students license, and highlights results in a copy of the file |

The button opens a console window where you sign in / pick a license - the GUI itself just collects your inputs.

## New Student File Format (XLSX, from the school MIS export)

Expected columns (case-insensitive):

```
Forename, Full Name, Middle Names, Preferred Name, Surname, Pupil Email Address,
Form, Form Tutor, Form Tutor Initials, Year (NC), Year Code
```

Run it with:

```powershell
.\New-M365Students.ps1 -XlsxPath ".\student account creation for bh.xlsx" -UsageLocation "QA"
```

What it does:

- **Duplicates**: if the same `Full Name` appears more than once in the sheet, only the first row is processed - later rows are skipped with a warning.
- **Existing accounts**: if `Pupil Email Address` is already filled in, it's verified against the tenant. Confirmed accounts are left alone and the cell is highlighted **green**. If the listed email doesn't actually exist, the student is treated as new. A name-match check against the tenant is also run as an extra safety net before creating anything.
- **New email generation**: you're asked once for a number to append (e.g. `26`). New addresses are built as the first 4 letters of the student's `Forename` + that number (e.g. `omar26@sherborneqatar.org`). If that address is taken, one more letter from the name is added and it tries again, repeating until a free address is found.
- **Password**: first-initial + last-initial + `student@123` (all lowercase), e.g. `oastudent@123`.
- **License**: you pick the A5 for Students SKU from the tenant's available licenses (the script suggests a likely match).
- **Output**: nothing is written to your original file - a copy named `<file> - processed.xlsx` is created with new accounts highlighted **yellow** (plus the generated UPN/password in two extra columns) and existing accounts highlighted **green**.

Add `-WhatIfOnly` to preview without creating any accounts.

## Requirements

- Windows PowerShell with the [Microsoft.Graph](https://aka.ms/graph/sdk/powershell) module:
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```
- `.xlsx` input requires the `ImportExcel` module - installed automatically on first use if missing.
- A Global Admin / User Administrator account to sign in with.

## Repository

**One-Liner**: `irm https://raw.githubusercontent.com/xnostra/Sherborne-Student-Tools/master/invoke-studenttoolkit.ps1 | iex`

**Repository**: https://github.com/xnostra/Sherborne-Student-Tools
