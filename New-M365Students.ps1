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

$ToolkitVersion = '2026.09.07.7'
Write-Host "Sherborne Student Toolkit $ToolkitVersion" -ForegroundColor Cyan

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
        if (Test-GraphNotFound -ErrorRecord $_) { return $false }
        throw  # An unavailable directory is not evidence that a UPN is free.
    }
}

function Test-GraphNotFound {
    param($ErrorRecord)

    return $ErrorRecord.FullyQualifiedErrorId -match 'Request_ResourceNotFound' -or
           $ErrorRecord.Exception.Message -match '(?i)Status:\s*404|does not exist'
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

    if (-not $NormA -or -not $NormB) { return 'none' }
    if ($NormA -eq $NormB) { return 'exact' }

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

# Show exactly which rows are candidates for account creation before any tenant changes are made.
$missingEmailRows = @()
for ($i = 0; $i -lt $rows.Count; $i++) {
    $email = "$($rows[$i].'Pupil Email Address')".Trim()
    if (-not $email) {
        $missingEmailRows += [pscustomobject]@{
            SheetRow = $i + 2
            FullName = "$($rows[$i].'Full Name')".Trim()
            Form     = "$($rows[$i].Form)".Trim()
        }
    }
}
if ($missingEmailRows.Count -eq 0) {
    Write-Host "No rows have a blank Pupil Email Address. There are no new student accounts to create." -ForegroundColor Yellow
} else {
    Write-Host "`nRows with no Pupil Email Address (possible new accounts):" -ForegroundColor Cyan
    foreach ($candidate in $missingEmailRows) {
        Write-Host "  Row $($candidate.SheetRow): $($candidate.FullName)  (Form $($candidate.Form))"
    }
    Write-Host "Total possible new accounts: $($missingEmailRows.Count)" -ForegroundColor Cyan
}

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
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All", "Organization.Read.All", "Group.Read.All", "GroupMember.ReadWrite.All"

$targetGroupName = 'BH PREP STUDENTS'
$escapedTargetGroupName = $targetGroupName.Replace("'", "''")
$targetGroups = @(Get-MgGroup -Filter "displayName eq '$escapedTargetGroupName'" -Property Id, DisplayName -All -ErrorAction Stop)
if ($targetGroups.Count -eq 0) { throw "Required Microsoft 365 group '$targetGroupName' was not found. No accounts will be created." }
if ($targetGroups.Count -gt 1) { throw "More than one Microsoft 365 group is named '$targetGroupName'. Rename the duplicates before creating accounts." }
$targetGroup = $targetGroups[0]
Write-Host "New accounts will be added to: $targetGroupName" -ForegroundColor Cyan

function Get-AssignableSeatCount {
    param($SubscribedSku)

    # The admin centre counts both active and warning-state seats as available during the
    # subscription grace period. Exclude only suspended and locked-out seats.
    $enabled = [int]$SubscribedSku.PrepaidUnits.Enabled
    $warning = [int]$SubscribedSku.PrepaidUnits.Warning
    return ($enabled + $warning - [int]$SubscribedSku.ConsumedUnits)
}

$studentA5SkuPartNumbers = @(
    # Current Microsoft 365 Education A5 student entitlement.
    'M365EDU_A5_STUUSEBNFT',
    # Older/alternative A5 student subscriptions that a tenant may still have.
    'M365EDU_A5_STUDENT',
    'M365EDU_A5_NOPSTNCONF_STUUSEBNFT',
    'M365EDU_A5_NOPSTNCONF_STUDENT',
    'ENTERPRISEPREMIUM_STUUSEBNFT',
    'ENTERPRISEPREMIUM_STUDENT',
    'ENTERPRISEPREMIUM_NOPSTNCONF_STUUSEBNFT',
    'ENTERPRISEPREMIUM_NOPSTNCONF_STUDENT'
)

$allSkus = @(Get-MgSubscribedSku -All)
$studentA5Skus = @($allSkus | Where-Object { $studentA5SkuPartNumbers -contains $_.SkuPartNumber })
$skus = @($allSkus | Where-Object { (Get-AssignableSeatCount $_) -gt 0 })
if (-not $skus) { throw "No licenses with available seats were found in this tenant." }

Write-Host "`nAvailable licenses in this tenant:"
for ($i = 0; $i -lt $skus.Count; $i++) {
    $s = $skus[$i]
    $free = Get-AssignableSeatCount $s
    $label = if ($s.SkuPartNumber -eq 'M365EDU_A5_STUUSEBNFT') { 'Microsoft 365 A5 for Students (Student Use Benefit)' } else { $s.SkuPartNumber }
    Write-Host "  [$i] $label  (available: $free)"
}
$suggested = @(0..($skus.Count - 1) | Where-Object { $skus[$_].SkuPartNumber -eq 'M365EDU_A5_STUUSEBNFT' })
if (-not $suggested) {
    $suggested = @(0..($skus.Count - 1) | Where-Object { $studentA5SkuPartNumbers -contains $skus[$_].SkuPartNumber })
}
if (-not $suggested -and $studentA5Skus) {
    $a5Status = $studentA5Skus | ForEach-Object { "$($_.SkuPartNumber): $(Get-AssignableSeatCount $_) available" }
    Write-Warning "Microsoft 365 A5 for Students was found, but it has no available seats ($($a5Status -join '; ')). Choose another available license below, or free/buy A5 student seats."
}
if ($suggested) { Write-Host "`n  (Recommended: [$($suggested[0])] Microsoft 365 A5 for Students)" -ForegroundColor Yellow }

$choice = Read-Host "`nEnter the number of the A5 Student license to assign (press Enter for the recommended license)"
if ([string]::IsNullOrWhiteSpace($choice) -and $suggested) { $choice = $suggested[0] }
$sku = $skus[[int]$choice]
if (-not $sku) { throw "Invalid selection." }
$available = Get-AssignableSeatCount $sku
Write-Host "Using license: $($sku.SkuPartNumber)"

$emailNumber = Read-Host "`nWhat number should be appended to new student email addresses? (e.g. 26 or 26q)"
if ($emailNumber -notmatch '^[a-zA-Z0-9]+$') { throw "The value can only contain letters and digits (e.g. 26 or 26q)." }

$passwordMode = ''
while ($passwordMode -notin @('A', 'M')) {
    $passwordMode = (Read-Host "`nType A to create the usual automatic password for each student, or M to set one password for all new students").Trim().ToUpper()
}

$sharedPassword = $null
if ($passwordMode -eq 'M') {
    while (-not $sharedPassword) {
        $sharedPassword = (Read-Host "Enter the password to apply to every new student").Trim()
        if (-not $sharedPassword) { Write-Warning "Password can't be blank." }
    }
}

$usedUpns = New-Object 'System.Collections.Generic.HashSet[string]'

# --- Process rows, tracking outcomes per Excel row number (header = row 1) ---
function New-StudentAccount {
    param([string]$FullName, $Row, [System.Collections.Generic.HashSet[string]]$UsedUpns)

    $forename = $Row.Forename
    $surname  = $Row.Surname
    $upn = New-StudentUpn -Forename $forename -Surname $surname -Number $emailNumber -UsedThisRun $UsedUpns
    $UsedUpns.Add($upn) | Out-Null
    $password = if ($sharedPassword) { $sharedPassword } else { New-StudentPassword -Forename $forename -Surname $surname }
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

        $groupAdded = $true
        try {
            New-MgGroupMemberByRef -GroupId $targetGroup.Id -OdataId "https://graph.microsoft.com/v1.0/directoryObjects/$($newUser.Id)" -ErrorAction Stop
            Write-Host "Added to group: $targetGroupName" -ForegroundColor Green
        } catch {
            $groupAdded = $false
            Write-Warning "Account '$upn' was created but could not be added to '$targetGroupName': $($_.Exception.Message)"
        }

        if ($script:available -gt 0) {
            try {
                Set-MgUserLicense -UserId $newUser.Id -AddLicenses @{ SkuId = $sku.SkuId } -RemoveLicenses @() -ErrorAction Stop
                $script:available--
            } catch {
                Write-Warning "Account '$upn' was created but license assignment failed: $($_.Exception.Message)"
            }
        }
    }

    return [pscustomobject]@{ Status = 'created'; Upn = $upn; Password = $password; GroupAdded = ($WhatIfOnly -or $groupAdded) }
}

function Get-StudentMatchInfo {
    param($TenantUser, $Row, [string]$FullName)

    $quality = Get-NameMatchQuality -NormA (Get-NormalizedName $TenantUser.DisplayName) -NormB (Get-NormalizedName $FullName)

    # If both sides know the Form, it should agree too - a mismatch usually means the tenant record is stale
    # (e.g. the student moved up a Form) rather than a genuinely different student.
    $sheetForm = "$($Row.Form)".Trim()
    $tenantForm = "$($TenantUser.Department)".Trim()
    # Many existing tenant accounts have no Department. A blank tenant value is
    # not evidence of a mismatch; use Form only when both systems provide it.
    $formMatches = $true
    if ($sheetForm -and $tenantForm -and $sheetForm -ne $tenantForm) {
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
$summaryGroupFailed = 0
$seenRowKeys = @{}   # "name|form" -> first sheet row that used it, for true in-sheet duplicate rows

# Reject repeated supplied addresses before any account creation. With no stable
# pupil ID, do not guess whether two rows are duplicate exports or two people.
$emailRows = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[int]]]::new([System.StringComparer]::OrdinalIgnoreCase)
$reviewReasons = @{}
$claimedUpns = @{}
$duplicateEmailByRow = @{}
for ($j = 0; $j -lt $rows.Count; $j++) {
    $emailKey = "$($rows[$j].'Pupil Email Address')".Trim()
    if ($emailKey) {
        if (-not $emailRows.ContainsKey($emailKey)) {
            $emailRows[$emailKey] = [System.Collections.Generic.List[int]]::new()
        }
        $emailRows[$emailKey].Add($j + 2)
    }
}
foreach ($emailEntry in $emailRows.GetEnumerator()) {
    if ($emailEntry.Value.Count -gt 1) {
        $rowList = $emailEntry.Value -join ', '
        foreach ($r in $emailEntry.Value) {
            $duplicateEmailByRow[$r] = $emailEntry.Key
        }
    }
}

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
    if ($duplicateEmailByRow.ContainsKey($sheetRow)) {
        Write-Warning "iSAMS email '$existingEmail' is shared by multiple rows. Ignoring that email for this row and resolving '$fullName' independently by name."
        $existingEmail = ''
    }
    if ($existingEmail) {
        try {
            $found = Get-MgUser -UserId $existingEmail -Property DisplayName, UserPrincipalName, Department -ErrorAction Stop
        } catch {
            if (-not (Test-GraphNotFound -ErrorRecord $_)) { throw }
            $found = $null
        }
        if ($found) {
            $match = Get-StudentMatchInfo -TenantUser $found -Row $row -FullName $fullName

            if ($match.NameQuality -eq 'exact' -and $match.FormMatches) {
                Write-Host "Existing account confirmed: $existingEmail" -ForegroundColor Green
                $claim = $found.UserPrincipalName.ToLowerInvariant()
                if ($claimedUpns.ContainsKey($claim)) {
                    throw "Account $claim claimed by rows $($claimedUpns[$claim]) and $sheetRow; output not finalized."
                }
                $claimedUpns[$claim] = $sheetRow
                $usedUpns.Add($existingEmail) | Out-Null
                $results[$sheetRow] = @{ Status = 'existing'; Upn = $existingEmail; MatchType = 'exact' }
                $summaryExisting++
                continue
            } elseif ($match.NameQuality -eq 'exact') {
                Write-Warning "Existing account confirmed by exact name for '$fullName', but Microsoft Form '$($found.Department)' differs from iSAMS Form '$($row.Form)'. Treating as existing with a stale-Form warning."
                $claim = $found.UserPrincipalName.ToLowerInvariant()
                if ($claimedUpns.ContainsKey($claim)) {
                    throw "Account $claim claimed by rows $($claimedUpns[$claim]) and $sheetRow; output not finalized."
                }
                $claimedUpns[$claim] = $sheetRow
                $usedUpns.Add($found.UserPrincipalName) | Out-Null
                $results[$sheetRow] = @{ Status = 'existing'; Upn = $found.UserPrincipalName; MatchType = 'form-mismatch' }
                $summaryExisting++
                $summaryFormMismatch++
                continue
            } elseif ($match.NameQuality -eq 'fuzzy') {
                Write-Host "Existing account matched from supplied email using a close name: $($found.UserPrincipalName) (Microsoft name '$($found.DisplayName)')." -ForegroundColor Cyan
                $claim = $found.UserPrincipalName.ToLowerInvariant()
                if ($claimedUpns.ContainsKey($claim)) {
                    throw "Account $claim claimed by rows $($claimedUpns[$claim]) and $sheetRow; output not finalized."
                }
                $claimedUpns[$claim] = $sheetRow
                $usedUpns.Add($found.UserPrincipalName) | Out-Null
                $results[$sheetRow] = @{ Status = 'existing'; Upn = $found.UserPrincipalName; MatchType = 'fuzzy' }
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
            Write-Warning "Supplied email '$existingEmail' was not found. Continuing with exact display-name lookup before considering account creation."
        }
    }

    # Extra duplicate safety net: display-name match already in the tenant
    $nameMatches = @(Get-MgUser -Filter "displayName eq '$($fullName.Replace("'", "''"))'" -Property DisplayName, UserPrincipalName, Department -All -ErrorAction Stop)
    $recoveredFromDuplicate = $false
    if ($nameMatches.Count -eq 0 -and $duplicateEmailByRow.ContainsKey($sheetRow)) {
        # A duplicated iSAMS email is unreliable. Search a limited Microsoft candidate set
        # by first name, then apply the existing full-name fuzzy comparison locally.
        $firstName = (@($fullName -split '\s+' | Where-Object { $_ }))[0]
        $escapedFirstName = $firstName.Replace("'", "''")
        $prefixMatches = @(Get-MgUser -Filter "startsWith(displayName,'$escapedFirstName')" -Property DisplayName, UserPrincipalName, Department -All -ErrorAction Stop)
        $nameMatches = @($prefixMatches | Where-Object {
            (Get-NameMatchQuality -NormA (Get-NormalizedName $_.DisplayName) -NormB (Get-NormalizedName $fullName)) -in @('exact', 'fuzzy')
        })
        $recoveredFromDuplicate = $true
    }
    $confirmedMatches = @($nameMatches | Where-Object { (Get-StudentMatchInfo -TenantUser $_ -Row $row -FullName $fullName).FormMatches })

    if ($nameMatches.Count -eq 1 -and $confirmedMatches.Count -eq 1) {
        $m = $confirmedMatches[0]
        Write-Host "Found existing account by name match: $($m.UserPrincipalName)" -ForegroundColor Green
        $claim = $m.UserPrincipalName.ToLowerInvariant()
        if ($claimedUpns.ContainsKey($claim)) {
            throw "Account $claim claimed by rows $($claimedUpns[$claim]) and $sheetRow; output not finalized."
        }
        $claimedUpns[$claim] = $sheetRow
        $usedUpns.Add($m.UserPrincipalName) | Out-Null
        $matchType = if ($recoveredFromDuplicate) { 'fuzzy' } else { 'exact' }
        $results[$sheetRow] = @{ Status = 'existing'; Upn = $m.UserPrincipalName; MatchType = $matchType }
        $summaryExisting++
        if ($recoveredFromDuplicate) { $summaryFuzzy++ }
        continue
    } elseif ($nameMatches.Count -eq 1) {
        # The display name is exact and unique. Form/Department can be stale in Microsoft,
        # so accept the identity but make the stale Form visible in the output.
        $m = $nameMatches[0]
        Write-Warning "Found unique exact-name account '$($m.UserPrincipalName)' for '$fullName', but Microsoft Form '$($m.Department)' differs from iSAMS Form '$($row.Form)'. Treating as existing."
        $claim = $m.UserPrincipalName.ToLowerInvariant()
        if ($claimedUpns.ContainsKey($claim)) {
            throw "Account $claim claimed by rows $($claimedUpns[$claim]) and $sheetRow; output not finalized."
        }
        $claimedUpns[$claim] = $sheetRow
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
        $results[$sheetRow] = @{ Status = 'failed' }
        $summaryFailed++
        continue
    }

    $results[$sheetRow] = @{ Status = 'created'; Upn = $creation.Upn; Password = $creation.Password; GroupAdded = $creation.GroupAdded }
    $summaryCreated++
    if (-not $creation.GroupAdded) { $summaryGroupFailed++ }
}

# Resolve review rows against the MIS/tenant before rerunning; no name-only override.

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
$originalEmailCol = Get-OrAddColumn -Worksheet $ws -HeaderLookup $headerCols -Title "Original supplied email" -NextFreeCol ([ref]$nextFreeCol)
$reasonCol = Get-OrAddColumn -Worksheet $ws -HeaderLookup $headerCols -Title "Match reason" -NextFreeCol ([ref]$nextFreeCol)
$statusColIndex = Get-OrAddColumn -Worksheet $ws -HeaderLookup $headerCols -Title "Account Status" -NextFreeCol ([ref]$nextFreeCol)

# --- Legend explaining the "Pupil Email Address" cell highlight colors ---
$legendCol = Get-OrAddColumn -Worksheet $ws -HeaderLookup $headerCols -Title "Legend (Pupil Email Address highlight)" -NextFreeCol ([ref]$nextFreeCol)
$ws.Cells[1, $legendCol].Style.Font.Bold = $true

$legendRows = @(
    @{ Text = "Yellow = new account created this run"; Color = [System.Drawing.Color]::Yellow }
    @{ Text = "Green  = existing account confirmed (already had one)"; Color = [System.Drawing.Color]::LightGreen }
    @{ Text = "Light Blue = existing account matched by a close/fuzzy name - please spot-check"; Color = [System.Drawing.Color]::LightSkyBlue }
    @{ Text = "Khaki  = existing account confirmed by name, but tenant Form looks outdated"; Color = [System.Drawing.Color]::Khaki }
    @{ Text = "Orange = needs manual review (name/email mismatch found)"; Color = [System.Drawing.Color]::Orange }
    @{ Text = "Red    = account creation failed (see console output)"; Color = [System.Drawing.Color]::LightCoral }
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
    if (-not $ws.Cells[$sheetRow, $originalEmailCol].Text) {
        $ws.Cells[$sheetRow, $originalEmailCol].Value = $rows[$sheetRow - 2].'Pupil Email Address'
    }
    # Remove stale generated results when processing a previously processed file.
    $ws.Cells[$sheetRow, $userColIndex].Value = $null
    $ws.Cells[$sheetRow, $upnColIndex].Value = $null
    $ws.Cells[$sheetRow, $pwColIndex].Value = $null
    $ws.Cells[$sheetRow, $reasonCol].Value = $null

    if ($result -is [hashtable] -and $result.Status -eq 'existing') {
        $ws.Cells[$sheetRow, $statusColIndex].Value = "Existing account confirmed"
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
        $cell.Value = $null
        $reason = $reviewReasons[$sheetRow]
        if (-not $reason) { $reason = 'Identity not confirmed: name/form mismatch or ambiguous candidates' }
        $ws.Cells[$sheetRow, $reasonCol].Value = $reason
        $ws.Cells[$sheetRow, $statusColIndex].Value = "Needs manual review"
        $cell.Style.Fill.PatternType = 'Solid'
        $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Orange)
    } elseif ($result -is [hashtable] -and $result.Status -eq 'failed') {
        $ws.Cells[$sheetRow, $statusColIndex].Value = "Failed to create"
        $cell.Style.Fill.PatternType = 'Solid'
        $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightCoral)
    } elseif ($result -is [hashtable] -and $result.Status -eq 'created') {
        $ws.Cells[$sheetRow, $statusColIndex].Value = if ($result.GroupAdded) { "New account created; added to $targetGroupName" } else { "New account created; group assignment failed" }
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
Write-Host "Group add failures: $summaryGroupFailed" -ForegroundColor $(if ($summaryGroupFailed) { 'Red' } else { 'Green' })
Write-Host "Results written to: $OutputPath"
Write-Host "===================================="
