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
  - Skips rows that are true duplicates within the sheet - same "Full Name" AND same "Form" as an
    earlier row. Same name but a DIFFERENT form is treated as two different students (e.g. two
    unrelated students who happen to share a name) and both are processed independently.
  - If "Pupil Email Address" is already filled in, verifies that account really exists in the tenant
    AND that its display name (and Form, when known) actually matches this row - not just that some
    account with that address exists.
      - Confirmed match          -> leave alone, highlight the cell GREEN.
      - Exists but doesn't match -> looks like it belongs to a different student with a similar
                                     name/address. Nothing is created; highlighted ORANGE for you
                                     to check manually.
      - Blank, or doesn't exist  -> treated as a new student (falls through to creation below).
  - As a second safety net (for rows with no listed email), it also searches the tenant by exact
    display name. A single confirmed match (name + Form agree) is treated as existing (GREEN).
    Multiple accounts sharing that name, or a match whose Form doesn't line up, is too risky to
    guess - flagged ORANGE for manual review instead of silently creating or skipping.
  - For genuinely new students: asks once for a number to append, then builds the email as
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

function Get-NormalizedName {
    param([string]$Name)
    return (($Name -replace '\s+', ' ').Trim().ToLower())
}

function Get-UsernameFromUpn {
    param([string]$Upn)
    return ($Upn -replace "@$([regex]::Escape($Domain))$", '')
}

function Get-LevenshteinDistance {
    param([string]$A, [string]$B)
    $lenA = $A.Length; $lenB = $B.Length
    $d = New-Object 'int[,]' ($lenA + 1), ($lenB + 1)
    for ($i = 0; $i -le $lenA; $i++) { $d[$i, 0] = $i }
    for ($j = 0; $j -le $lenB; $j++) { $d[0, $j] = $j }
    for ($i = 1; $i -le $lenA; $i++) {
        for ($j = 1; $j -le $lenB; $j++) {
            $cost = if ($A[$i - 1] -eq $B[$j - 1]) { 0 } else { 1 }
            $deleteCost     = $d[($i - 1), $j] + 1
            $insertCost     = $d[$i, ($j - 1)] + 1
            $substituteCost = $d[($i - 1), ($j - 1)] + $cost
            $d[$i, $j] = [Math]::Min([Math]::Min($deleteCost, $insertCost), $substituteCost)
        }
    }
    return $d[$lenA, $lenB]
}

# Compares two already-normalized (lowercase, single-spaced) names and returns how confident a match is:
#   'exact' - identical
#   'fuzzy' - almost certainly the same person: one name is missing/has extra middle name(s) compared to the
#             other (in the same order), or the two strings are a close spelling/spacing match overall
#   'none'  - not similar enough to trust automatically
function Get-NameMatchQuality {
    param([string]$NormA, [string]$NormB)

    if ($NormA -eq $NormB) { return 'exact' }
    if (-not $NormA -or -not $NormB) { return 'none' }

    $tokensA = @($NormA -split ' ' | Where-Object { $_ })
    $tokensB = @($NormB -split ' ' | Where-Object { $_ })
    $shorter = if ($tokensA.Count -le $tokensB.Count) { $tokensA } else { $tokensB }
    $longer  = if ($tokensA.Count -le $tokensB.Count) { $tokensB } else { $tokensA }

    # Missing/extra middle name(s): every token in the shorter name appears in the longer name, in the same
    # order (allow up to 2 tokens missing, e.g. a dropped middle name or two).
    if ($shorter.Count -ge 2 -and ($longer.Count - $shorter.Count) -ge 1 -and ($longer.Count - $shorter.Count) -le 2) {
        $isSubsequence = $true
        $pos = 0
        foreach ($tok in $shorter) {
            $foundAt = -1
            for ($k = $pos; $k -lt $longer.Count; $k++) {
                if ($longer[$k] -eq $tok) { $foundAt = $k; break }
            }
            if ($foundAt -lt 0) { $isSubsequence = $false; break }
            $pos = $foundAt + 1
        }
        if ($isSubsequence) { return 'fuzzy' }
    }

    # Otherwise, fall back to overall edit-distance similarity (catches typos, missing/extra spaces, etc.)
    $maxLen = [Math]::Max($NormA.Length, $NormB.Length)
    if ($maxLen -ge 6) {
        $distance = Get-LevenshteinDistance -A $NormA -B $NormB
        $ratio = 1 - ($distance / $maxLen)
        if ($ratio -ge 0.82) { return 'fuzzy' }
    }

    return 'none'
}

# --- Load rows ---
$rows = Import-Excel -Path $XlsxPath

# --- Duplicate full-name detection within the sheet ---
# Same Full Name + same Form is treated as a genuine duplicate row (skipped).
# Same Full Name but a DIFFERENT Form is treated as two different students who happen to
# share a name - both are processed, just flagged here so you know to sanity-check them.
$nameCounts = @{}
foreach ($row in $rows) {
    $key = Get-NormalizedName $row.'Full Name'
    if (-not $nameCounts.ContainsKey($key)) { $nameCounts[$key] = 0 }
    $nameCounts[$key]++
}
foreach ($name in ($nameCounts.Keys | Where-Object { $nameCounts[$_] -gt 1 })) {
    $forms = $rows | Where-Object { (Get-NormalizedName $_.'Full Name') -eq $name } | ForEach-Object { $_.Form } | Select-Object -Unique
    if ($forms.Count -gt 1) {
        Write-Host "Note: the name '$name' appears in more than one form ($($forms -join ', ')) - treating these as different students, not duplicates." -ForegroundColor Cyan
    }
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

$emailNumber = Read-Host "`nWhat number should be appended to new student email addresses? (e.g. 26 or 26q)"
if ($emailNumber -notmatch '^[a-zA-Z0-9]+$') { throw "The value can only contain letters and digits (e.g. 26 or 26q)." }

$usedUpns = New-Object 'System.Collections.Generic.HashSet[string]'

# --- Process rows, tracking outcomes per Excel row number (header = row 1) ---
function New-StudentAccount {
    param([string]$FullName, $Row, [System.Collections.Generic.HashSet[string]]$UsedUpns)

    $forename = $Row.Forename
    $surname  = $Row.Surname
    $upn = New-StudentUpn -Forename $forename -Surname $surname -Number $emailNumber -UsedThisRun $UsedUpns
    $UsedUpns.Add($upn) | Out-Null
    $password = New-StudentPassword -Forename $forename -Surname $surname
    $mailNickname = $upn.Split('@')[0]

    Write-Host "New UPN:  $upn"
    Write-Host "Password: $password"
    Write-Host "License:  $($sku.SkuPartNumber)"

    if (-not $WhatIfOnly) {
        if ($script:available -le 0) {
            Write-Warning "No available '$($sku.SkuPartNumber)' licenses left - creating account without a license."
        }

        $newUserParams = @{
            DisplayName       = $FullName
            GivenName         = $forename
            Surname           = $surname
            UserPrincipalName = $upn
            MailNickname      = $mailNickname
            JobTitle          = "Student"
            UsageLocation     = $UsageLocation
            AccountEnabled    = $true
            PasswordProfile   = @{
                Password                      = $password
                ForceChangePasswordNextSignIn = $false
            }
        }
        if ($Row.Form) { $newUserParams.Department = $Row.Form.ToString().Trim() }

        try {
            $newUser = New-MgUser @newUserParams -ErrorAction Stop
        } catch {
            Write-Warning "Failed to create account for '$FullName': $($_.Exception.Message)"
            return [pscustomobject]@{ Status = 'failed' }
        }

        if ($script:available -gt 0) {
            Set-MgUserLicense -UserId $newUser.Id -AddLicenses @{ SkuId = $sku.SkuId } -RemoveLicenses @()
            $script:available--
        }
    }

    return [pscustomobject]@{ Status = 'created'; Upn = $upn; Password = $password }
}

function Get-StudentMatchInfo {
    param($TenantUser, $Row, [string]$FullName)

    $quality = Get-NameMatchQuality -NormA (Get-NormalizedName $TenantUser.DisplayName) -NormB (Get-NormalizedName $FullName)

    # If both sides know the Form, it should agree too - a mismatch usually means the tenant record is stale
    # (e.g. the student moved up a Form) rather than a genuinely different student.
    $formMatches = $true
    if ($Row.Form -and $TenantUser.Department -and ($Row.Form.ToString().Trim() -ne $TenantUser.Department.Trim())) {
        $formMatches = $false
    }

    return [pscustomobject]@{ NameQuality = $quality; FormMatches = $formMatches }
}

$results = @{}   # rowIndex (1-based, matching sheet row = index+2) -> 'created' | 'existing' | 'skipped' | 'review'
$summaryCreated = 0
$summaryExisting = 0
$summarySkippedDup = 0
$summaryFailed = 0
$summaryReview = 0
$summaryFuzzy = 0
$summaryFormMismatch = 0
$seenRowKeys = @{}   # "name|form" -> first sheet row that used it, for true in-sheet duplicate rows

for ($i = 0; $i -lt $rows.Count; $i++) {
    $row = $rows[$i]
    $sheetRow = $i + 2
    $fullName = $row.'Full Name'.Trim()
    $key = Get-NormalizedName $fullName

    Write-Host "`n--- Row $sheetRow`: $fullName ---"

    $rowKey = "$key|$($row.Form)"
    if ($seenRowKeys.ContainsKey($rowKey)) {
        Write-Warning "Duplicate row: '$fullName' in form '$($row.Form)' already handled at row $($seenRowKeys[$rowKey]). Skipping this row."
        $results[$sheetRow] = 'skipped'
        $summarySkippedDup++
        continue
    }
    $seenRowKeys[$rowKey] = $sheetRow

    $existingEmail = if ($row.'Pupil Email Address') { $row.'Pupil Email Address'.ToString().Trim() } else { '' }
    if ($existingEmail) {
        try {
            $found = Get-MgUser -UserId $existingEmail -Property DisplayName, UserPrincipalName, Department -ErrorAction Stop
        } catch {
            $found = $null
        }
        if ($found) {
            $match = Get-StudentMatchInfo -TenantUser $found -Row $row -FullName $fullName

            if ($match.NameQuality -eq 'exact' -and $match.FormMatches) {
                Write-Host "Existing account confirmed: $existingEmail" -ForegroundColor Green
                $usedUpns.Add($existingEmail) | Out-Null
                $results[$sheetRow] = @{ Status = 'existing'; Upn = $existingEmail; MatchType = 'exact' }
                $summaryExisting++
                continue
            } elseif ($match.NameQuality -eq 'exact' -and -not $match.FormMatches) {
                Write-Warning "Existing account confirmed by name for '$fullName' ($existingEmail), but tenant Form '$($found.Department)' differs from sheet Form '$($row.Form)' - likely outdated, treating as existing."
                $usedUpns.Add($existingEmail) | Out-Null
                $results[$sheetRow] = @{ Status = 'existing'; Upn = $existingEmail; MatchType = 'form-mismatch' }
                $summaryExisting++
                $summaryFormMismatch++
                continue
            } elseif ($match.NameQuality -eq 'fuzzy') {
                $noteSuffix = if (-not $match.FormMatches) { " (tenant Form '$($found.Department)' also differs from sheet Form '$($row.Form)')" } else { "" }
                Write-Host "Existing account matched by close name: $existingEmail (tenant name '$($found.DisplayName)')$noteSuffix - please spot-check." -ForegroundColor Cyan
                $usedUpns.Add($existingEmail) | Out-Null
                $results[$sheetRow] = @{ Status = 'existing'; Upn = $existingEmail; MatchType = 'fuzzy' }
                $summaryExisting++
                $summaryFuzzy++
                continue
            } else {
                Write-Warning "Sheet lists '$existingEmail' for '$fullName' (form '$($row.Form)'), but that account belongs to '$($found.DisplayName)' (form '$($found.Department)') - this looks like a different student. Flagging for manual review instead of assuming they're the same."
                $results[$sheetRow] = 'review'
                $summaryReview++
                continue
            }
        } else {
            Write-Warning "Sheet lists '$existingEmail' but no such account exists in the tenant - treating as a new student."
        }
    }

    # Extra duplicate safety net: display-name match already in the tenant
    $nameMatches = @(Get-MgUser -Filter "displayName eq '$($fullName.Replace("'", "''"))'" -Property DisplayName, UserPrincipalName, Department -All -ErrorAction SilentlyContinue)
    $confirmedMatches = @($nameMatches | Where-Object { (Get-StudentMatchInfo -TenantUser $_ -Row $row -FullName $fullName).FormMatches })

    if ($confirmedMatches.Count -eq 1) {
        $m = $confirmedMatches[0]
        Write-Host "Found existing account by name match: $($m.UserPrincipalName)" -ForegroundColor Green
        $usedUpns.Add($m.UserPrincipalName) | Out-Null
        $results[$sheetRow] = @{ Status = 'existing'; Upn = $m.UserPrincipalName; MatchType = 'exact' }
        $summaryExisting++
        continue
    } elseif ($nameMatches.Count -eq 1) {
        # Exact display-name match, but the Form doesn't line up - most likely the tenant record is just
        # stale (student moved up a Form), not a different student. Treat as existing, but flag it.
        $m = $nameMatches[0]
        Write-Warning "Found existing account '$($m.UserPrincipalName)' by exact name match for '$fullName', but tenant Form '$($m.Department)' differs from sheet Form '$($row.Form)' - likely outdated, treating as existing."
        $usedUpns.Add($m.UserPrincipalName) | Out-Null
        $results[$sheetRow] = @{ Status = 'existing'; Upn = $m.UserPrincipalName; MatchType = 'form-mismatch' }
        $summaryExisting++
        $summaryFormMismatch++
        continue
    } elseif ($nameMatches.Count -gt 0) {
        # Multiple accounts share this exact name and none has a matching Form - too risky to guess,
        # so leave it for a human to check.
        $candidates = ($nameMatches | ForEach-Object { "$($_.UserPrincipalName) (form '$($_.Department)')" }) -join '; '
        Write-Warning "Found tenant account(s) named '$fullName' but couldn't confirm a match for form '$($row.Form)': $candidates. Flagging for manual review; no account will be created automatically."
        $results[$sheetRow] = 'review'
        $summaryReview++
        continue
    }

    $creation = New-StudentAccount -FullName $fullName -Row $row -UsedUpns $usedUpns
    if ($creation.Status -eq 'failed') {
        $results[$sheetRow] = 'skipped'
        $summaryFailed++
        continue
    }

    $results[$sheetRow] = @{ Status = 'created'; Upn = $creation.Upn; Password = $creation.Password }
    $summaryCreated++
}

# --- Optional: force-create accounts for rows flagged for manual review ---
$reviewRowsList = @()
foreach ($sheetRow in $results.Keys) {
    if ($results[$sheetRow] -eq 'review') {
        $rowIdx = $sheetRow - 2
        $reviewRowsList += [pscustomobject]@{ SheetRow = $sheetRow; Row = $rows[$rowIdx]; FullName = $rows[$rowIdx].'Full Name'.Trim() }
    }
}

if ($reviewRowsList.Count -gt 0) {
    Write-Host "`n===================================="
    Write-Host "The following $($reviewRowsList.Count) row(s) were flagged for manual review:" -ForegroundColor DarkYellow
    foreach ($r in ($reviewRowsList | Sort-Object SheetRow)) { Write-Host "  Row $($r.SheetRow): $($r.FullName)" }

    if ($WhatIfOnly) {
        Write-Host "`n(Running with -WhatIfOnly, so skipping the force-create step.)" -ForegroundColor Cyan
    } else {
        Write-Host "`nIf any of these should just be created as brand-new accounts anyway (ignoring the match check above),"
        Write-Host "paste their Full Names below - one per line, exactly as shown above. Press Enter on a blank line when done."
        Write-Host "(Leave blank and press Enter immediately to skip this step.)" -ForegroundColor Cyan

        $pastedNames = New-Object System.Collections.Generic.List[string]
        while ($true) {
            $line = Read-Host
            if ([string]::IsNullOrWhiteSpace($line)) { break }
            $pastedNames.Add($line.Trim())
        }

        foreach ($pastedName in $pastedNames) {
            $match = $reviewRowsList | Where-Object { (Get-NormalizedName $_.FullName) -eq (Get-NormalizedName $pastedName) } | Select-Object -First 1
            if (-not $match) {
                Write-Warning "Couldn't find a flagged row matching '$pastedName' - skipping."
                continue
            }

            Write-Host "`nForce-creating: $($match.FullName) (Row $($match.SheetRow))" -ForegroundColor Yellow
            $creation = New-StudentAccount -FullName $match.FullName -Row $match.Row -UsedUpns $usedUpns
            if ($creation.Status -eq 'failed') {
                continue
            }

            $results[$match.SheetRow] = @{ Status = 'created'; Upn = $creation.Upn; Password = $creation.Password }
            $summaryCreated++
            $summaryReview--
        }
    }
}

# --- Write results + highlighting back into the copy ---
$pkg = Open-ExcelPackage -Path $OutputPath
$ws = $pkg.Workbook.Worksheets[1]

$headerCols = @{}
for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
    $headerCols[$ws.Cells[1, $c].Text] = $c
}
$emailCol = $headerCols['Pupil Email Address']

function Get-OrAddColumn {
    param($Worksheet, $HeaderLookup, [string]$Title, [ref]$NextFreeCol)

    if ($HeaderLookup.ContainsKey($Title)) { return $HeaderLookup[$Title] }

    $col = $NextFreeCol.Value
    $Worksheet.Cells[1, $col].Value = $Title
    $HeaderLookup[$Title] = $col
    $NextFreeCol.Value++
    return $col
}

$nextFreeCol = $ws.Dimension.End.Column + 1
$upnColIndex  = Get-OrAddColumn -Worksheet $ws -HeaderLookup $headerCols -Title "Created UPN" -NextFreeCol ([ref]$nextFreeCol)
$userColIndex = Get-OrAddColumn -Worksheet $ws -HeaderLookup $headerCols -Title "Username" -NextFreeCol ([ref]$nextFreeCol)
$pwColIndex   = Get-OrAddColumn -Worksheet $ws -HeaderLookup $headerCols -Title "Created Password" -NextFreeCol ([ref]$nextFreeCol)

# --- Legend explaining the "Pupil Email Address" cell highlight colors ---
$legendCol = Get-OrAddColumn -Worksheet $ws -HeaderLookup $headerCols -Title "Legend (Pupil Email Address highlight)" -NextFreeCol ([ref]$nextFreeCol)
$ws.Cells[1, $legendCol].Style.Font.Bold = $true

$legendRows = @(
    @{ Text = "Yellow = new account created this run"; Color = [System.Drawing.Color]::Yellow }
    @{ Text = "Green  = existing account confirmed (already had one)"; Color = [System.Drawing.Color]::LightGreen }
    @{ Text = "Light Blue = existing account matched by a close/fuzzy name - please spot-check"; Color = [System.Drawing.Color]::LightSkyBlue }
    @{ Text = "Khaki  = existing account confirmed by name, but tenant Form looks outdated"; Color = [System.Drawing.Color]::Khaki }
    @{ Text = "Orange = needs manual review (name/email mismatch found)"; Color = [System.Drawing.Color]::Orange }
    @{ Text = "No fill = duplicate row, skipped"; Color = $null }
)
for ([int]$legendIdx = 0; $legendIdx -lt $legendRows.Count; $legendIdx++) {
    $legendRowNum = 2 + $legendIdx
    $legendCell = $ws.Cells[$legendRowNum, $legendCol]
    $legendCell.Value = $legendRows[$legendIdx].Text
    if ($legendRows[$legendIdx].Color) {
        $legendCell.Style.Fill.PatternType = 'Solid'
        $legendCell.Style.Fill.BackgroundColor.SetColor($legendRows[$legendIdx].Color)
    }
}
$ws.Column($legendCol).Width = 55

foreach ($sheetRow in $results.Keys) {
    $result = $results[$sheetRow]
    $cell = $ws.Cells[$sheetRow, $emailCol]

    if ($result -is [hashtable] -and $result.Status -eq 'existing') {
        $cell.Value = $result.Upn
        $cell.Style.Fill.PatternType = 'Solid'
        $color = switch ($result.MatchType) {
            'fuzzy'         { [System.Drawing.Color]::LightSkyBlue }
            'form-mismatch' { [System.Drawing.Color]::Khaki }
            default         { [System.Drawing.Color]::LightGreen }
        }
        $cell.Style.Fill.BackgroundColor.SetColor($color)
        $ws.Cells[$sheetRow, $userColIndex].Value = Get-UsernameFromUpn -Upn $result.Upn
    } elseif ($result -eq 'review') {
        $cell.Style.Fill.PatternType = 'Solid'
        $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Orange)
    } elseif ($result -is [hashtable] -and $result.Status -eq 'created') {
        $cell.Value = $result.Upn
        $cell.Style.Fill.PatternType = 'Solid'
        $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
        $ws.Cells[$sheetRow, $upnColIndex].Value = $result.Upn
        $ws.Cells[$sheetRow, $userColIndex].Value = Get-UsernameFromUpn -Upn $result.Upn
        $ws.Cells[$sheetRow, $pwColIndex].Value = $result.Password
    }
}

Close-ExcelPackage $pkg

Write-Host "`n===================================="
Write-Host "Created:            $summaryCreated" -ForegroundColor Yellow
Write-Host "Already existing:   $summaryExisting" -ForegroundColor Green
Write-Host "  (of which, fuzzy name match - spot-check): $summaryFuzzy" -ForegroundColor Cyan
Write-Host "  (of which, Form looked outdated):          $summaryFormMismatch" -ForegroundColor DarkYellow
Write-Host "Skipped (dup rows): $summarySkippedDup"
Write-Host "Needs manual review: $summaryReview" -ForegroundColor DarkYellow
Write-Host "Failed to create:   $summaryFailed" -ForegroundColor Red
Write-Host "Results written to: $OutputPath"
Write-Host "===================================="
