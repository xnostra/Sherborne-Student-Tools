<#
Creates M365 student accounts from the school MIS export (e.g. "student account creation for bh.xlsx")
and assigns the A5 for Students license. Run manually by a Global Admin / User Administrator.
Requires the Microsoft.Graph module:
    Install-Module Microsoft.Graph -Scope CurrentUser
(.xlsx input also requires the ImportExcel module - installed automatically if missing)

Expected columns (case-insensitive), matching the MIS export:
    Forename, Full Name, Middle Names, Preferred Name, Surname, Pupil Email Address,
    Form, Form Tutor, Form Tutor Initials, Year (NC), Year Code

What it does:
  - Skips rows whose "Full Name" is a duplicate elsewhere in the sheet (warns, processes the first occurrence only).
  - If "Pupil Email Address" is already filled in, verifies that account really exists in the tenant.
      - Exists and matches       -> leave alone, highlight the cell GREEN.
      - Blank, or doesn't exist  -> treated as a new student (falls through to creation below).
  - For new students: asks once for a number to append, then builds the email as
        <first 4 letters of Forename><number>@sherborneqatar.org
    If that UPN is already taken (in the tenant, or already used earlier in this run), it takes one
    more letter from the name and tries again, repeating until a free address is found.
  - Creates the account, assigns the A5 for Students license you pick, and highlights the cell YELLOW.
  - Writes the results to a COPY of the input file (so your original export is never touched) with
    the highlighting applied, plus a UPN/Password column for new accounts.

Usage:
    .\New-M365Students.ps1 -XlsxPath ".\student account creation for bh.xlsx" -UsageLocation "QA"

Add -WhatIfOnly to preview without creating any accounts or writing to Graph.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$XlsxPath,

    [Parameter(Mandatory = $true)]
    [string]$UsageLocation,

    [string]$Domain = "sherborneqatar.org",

    [string]$OutputPath,

    [switch]$WhatIfOnly
)

function Test-SherborneToolAccess {
    try {
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/xnostra/Sherborne-Student-Tools/master/README.md" -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop | Out-Null
    } catch {
        Write-Host ""
        Write-Host "This tool could not verify access to its source repository and cannot continue." -ForegroundColor Red
        Write-Host "(github.com/xnostra/Sherborne-Student-Tools)" -ForegroundColor Red
        exit 1
    }
}
Test-SherborneToolAccess

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module (needed to read/write .xlsx files)..."
    Install-Module ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

if (-not (Test-Path $XlsxPath)) {
    throw "File not found: $XlsxPath"
}

if (-not $OutputPath) {
    $dir  = Split-Path -Parent (Resolve-Path $XlsxPath)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($XlsxPath)
    $OutputPath = Join-Path $dir "$name - processed.xlsx"
}
Copy-Item -Path $XlsxPath -Destination $OutputPath -Force
Write-Host "Working on a copy: $OutputPath" -ForegroundColor Cyan

function Get-CleanLetters {
    param([string]$Text)
    return ($Text -replace '[^a-zA-Z]', '').ToLower()
}

function Test-UpnTaken {
    param([string]$Upn, [System.Collections.Generic.HashSet[string]]$UsedThisRun)

    if ($UsedThisRun.Contains($Upn)) { return $true }
    try {
        $existing = Get-MgUser -UserId $Upn -ErrorAction Stop
        return $null -ne $existing
    } catch {
        return $false
    }
}

function New-StudentUpn {
    param(
        [string]$Forename,
        [string]$Surname,
        [string]$Number,
        [System.Collections.Generic.HashSet[string]]$UsedThisRun
    )

    $foreLetters = Get-CleanLetters $Forename
    $surLetters  = Get-CleanLetters $Surname

    $take = [Math]::Min(4, $foreLetters.Length)
    if ($take -eq 0) { throw "Cannot build an email for a row with no usable Forename letters." }

    for ($extra = 0; $extra -le ($foreLetters.Length + $surLetters.Length - $take); $extra++) {
        $len = $take + $extra
        if ($len -le $foreLetters.Length) {
            $prefix = $foreLetters.Substring(0, $len)
        } else {
            $fromSurname = $len - $foreLetters.Length
            if ($fromSurname -gt $surLetters.Length) { break }
            $prefix = $foreLetters + $surLetters.Substring(0, $fromSurname)
        }

        $candidateUpn = "$prefix$Number@$Domain"
        if (-not (Test-UpnTaken -Upn $candidateUpn -UsedThisRun $UsedThisRun)) {
            return $candidateUpn
        }
    }

    throw "Could not find a free email address for '$Forename $Surname' with number '$Number' - ran out of letters to try."
}

function New-StudentPassword {
    param([string]$Forename, [string]$Surname)
    $f = (Get-CleanLetters $Forename).Substring(0, 1)
    $s = (Get-CleanLetters $Surname).Substring(0, 1)
    return "${f}${s}student@123"
}

# --- Load rows ---
$rows = Import-Excel -Path $XlsxPath

# --- Duplicate full-name detection within the sheet ---
$nameCounts = @{}
foreach ($row in $rows) {
    $key = ($row.'Full Name'.Trim().ToLower())
    if (-not $nameCounts.ContainsKey($key)) { $nameCounts[$key] = 0 }
    $nameCounts[$key]++
}

# --- Ask for the license and the email number up front ---
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "Organization.Read.All"

$skus = Get-MgSubscribedSku | Where-Object { ($_.PrepaidUnits.Enabled - $_.ConsumedUnits) -gt 0 }
if (-not $skus) { throw "No licenses with available seats were found in this tenant." }

Write-Host "`nAvailable licenses in this tenant:"
for ($i = 0; $i -lt $skus.Count; $i++) {
    $s = $skus[$i]
    $free = $s.PrepaidUnits.Enabled - $s.ConsumedUnits
    Write-Host "  [$i] $($s.SkuPartNumber)  (available: $free)"
}
$suggested = 0..($skus.Count - 1) | Where-Object { $skus[$_].SkuPartNumber -match 'A5' -and $skus[$_].SkuPartNumber -match 'STU' }
if ($suggested) { Write-Host "`n  (Looks like the A5 for Students SKU might be [$($suggested[0])])" -ForegroundColor Yellow }

$choice = Read-Host "`nEnter the number of the A5 Student license to assign"
$sku = $skus[[int]$choice]
if (-not $sku) { throw "Invalid selection." }
$available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
Write-Host "Using license: $($sku.SkuPartNumber)"

$emailNumber = Read-Host "`nWhat number should be appended to new student email addresses? (e.g. 26)"
if ($emailNumber -notmatch '^\d+$') { throw "The number must be digits only." }

$usedUpns = New-Object 'System.Collections.Generic.HashSet[string]'

# --- Process rows, tracking outcomes per Excel row number (header = row 1) ---
$results = @{}   # rowIndex (1-based, matching sheet row = index+2) -> 'created' | 'existing' | 'skipped'
$summaryCreated = 0
$summaryExisting = 0
$summarySkippedDup = 0

for ($i = 0; $i -lt $rows.Count; $i++) {
    $row = $rows[$i]
    $sheetRow = $i + 2
    $fullName = $row.'Full Name'.Trim()
    $key = $fullName.ToLower()

    Write-Host "`n--- Row $sheetRow`: $fullName ---"

    if ($nameCounts[$key] -gt 1) {
        $firstIndexForName = 0..($rows.Count - 1) | Where-Object { $rows[$_].'Full Name'.Trim().ToLower() -eq $key } | Select-Object -First 1
        if ($i -ne $firstIndexForName) {
            Write-Warning "Duplicate name '$fullName' - already handled at row $($firstIndexForName + 2). Skipping this row."
            $results[$sheetRow] = 'skipped'
            $summarySkippedDup++
            continue
        }
    }

    $existingEmail = ($row.'Pupil Email Address' | ForEach-Object { $_.ToString().Trim() })
    if ($existingEmail) {
        try {
            $found = Get-MgUser -UserId $existingEmail -ErrorAction Stop
        } catch {
            $found = $null
        }
        if ($found) {
            Write-Host "Existing account confirmed: $existingEmail" -ForegroundColor Green
            $usedUpns.Add($existingEmail) | Out-Null
            $results[$sheetRow] = 'existing'
            $summaryExisting++
            continue
        } else {
            Write-Warning "Sheet lists '$existingEmail' but no such account exists in the tenant - treating as a new student."
        }
    }

    # Extra duplicate safety net: exact display-name match already in the tenant
    $byName = Get-MgUser -Filter "displayName eq '$($fullName.Replace("'", "''"))'" -ErrorAction SilentlyContinue
    if ($byName) {
        Write-Host "Found existing account by name match: $($byName.UserPrincipalName)" -ForegroundColor Green
        $usedUpns.Add($byName.UserPrincipalName) | Out-Null
        $results[$sheetRow] = 'existing'
        $summaryExisting++
        continue
    }

    $forename = $row.Forename
    $surname  = $row.Surname
    $upn = New-StudentUpn -Forename $forename -Surname $surname -Number $emailNumber -UsedThisRun $usedUpns
    $usedUpns.Add($upn) | Out-Null
    $password = New-StudentPassword -Forename $forename -Surname $surname
    $mailNickname = $upn.Split('@')[0]

    Write-Host "New UPN:  $upn"
    Write-Host "Password: $password"
    Write-Host "License:  $($sku.SkuPartNumber)"

    if (-not $WhatIfOnly) {
        if ($available -le 0) {
            Write-Warning "No available '$($sku.SkuPartNumber)' licenses left - creating account without a license."
        }

        $newUser = New-MgUser -DisplayName $fullName `
            -GivenName $forename `
            -Surname $surname `
            -UserPrincipalName $upn `
            -MailNickname $mailNickname `
            -JobTitle "Student" `
            -Department $row.Form `
            -UsageLocation $UsageLocation `
            -AccountEnabled `
            -PasswordProfile @{
                Password                      = $password
                ForceChangePasswordNextSignIn = $false
            }

        if ($available -gt 0) {
            Set-MgUserLicense -UserId $newUser.Id -AddLicenses @{ SkuId = $sku.SkuId } -RemoveLicenses @()
            $available--
        }
    }

    $results[$sheetRow] = @{ Status = 'created'; Upn = $upn; Password = $password }
    $summaryCreated++
}

# --- Write results + highlighting back into the copy ---
$pkg = Open-ExcelPackage -Path $OutputPath
$ws = $pkg.Workbook.Worksheets[1]

$headerCols = @{}
for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
    $headerCols[$ws.Cells[1, $c].Text] = $c
}
$emailCol = $headerCols['Pupil Email Address']

$upnColIndex = $ws.Dimension.End.Column + 1
$pwColIndex  = $upnColIndex + 1
$ws.Cells[1, $upnColIndex].Value = "Created UPN"
$ws.Cells[1, $pwColIndex].Value  = "Created Password"

foreach ($sheetRow in $results.Keys) {
    $result = $results[$sheetRow]
    $cell = $ws.Cells[$sheetRow, $emailCol]

    if ($result -eq 'existing') {
        $cell.Style.Fill.PatternType = 'Solid'
        $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen)
    } elseif ($result -is [hashtable] -and $result.Status -eq 'created') {
        $cell.Value = $result.Upn
        $cell.Style.Fill.PatternType = 'Solid'
        $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
        $ws.Cells[$sheetRow, $upnColIndex].Value = $result.Upn
        $ws.Cells[$sheetRow, $pwColIndex].Value = $result.Password
    }
}

Close-ExcelPackage $pkg

Write-Host "`n===================================="
Write-Host "Created:            $summaryCreated" -ForegroundColor Yellow
Write-Host "Already existing:   $summaryExisting" -ForegroundColor Green
Write-Host "Skipped (dup rows): $summarySkippedDup"
Write-Host "Results written to: $OutputPath"
Write-Host "===================================="
