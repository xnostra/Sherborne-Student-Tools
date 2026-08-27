<#
Resets a student's password - one account at a time, or in bulk for a whole class. Built for Prep
and SEN ("special child") groups where a whole class often needs a fresh, easy password at once.
Run manually by a Global Admin / User Administrator. Requires the Microsoft.Graph module:
    Install-Module Microsoft.Graph -Scope CurrentUser
(.xlsx input also requires the ImportExcel module - installed automatically if missing)

Usage - single account:
    .\Set-M365StudentPasswords.ps1 -Email pupil@sherborneqatar.org

Usage - bulk from CSV or XLSX:
    .\Set-M365StudentPasswords.ps1 -CsvPath ".\class-list.csv"
    .\Set-M365StudentPasswords.ps1 -CsvPath ".\class-list.xlsx"

The bulk file just needs a column of email addresses - either "Email" or "Pupil Email Address"
(case-insensitive) is accepted, so you can point it straight at a class export.

Either way, it asks:
    A = auto-generate an easy password (a different one per account, for bulk)
    M = you type the password yourself (one shared password for everyone, for bulk)

Results (email + password used) are printed to the console. For bulk, they're also saved next to
your input file as "<file> - passwords.csv" so they're easy to hand out.
#>

param(
    [string]$Email,

    [string]$CsvPath
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

function Import-DataFile {
    param([string]$Path)

    if ($Path -match '\.xlsx$') {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Host "Installing ImportExcel module (needed to read .xlsx files)..."
            Install-Module ImportExcel -Scope CurrentUser -Force
        }
        Import-Module ImportExcel
        return Import-Excel -Path $Path
    }

    return Import-Csv -Path $Path
}

function Get-EmailColumnName {
    param($Row)
    foreach ($candidate in @('Email', 'Pupil Email Address')) {
        if ($Row.PSObject.Properties.Name -contains $candidate) { return $candidate }
    }
    throw "Couldn't find an 'Email' or 'Pupil Email Address' column in the file."
}

function Get-CleanEmail {
    param([string]$Text)
    if ($Text -match '[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}') {
        return $matches[0]
    }
    return $Text.Trim()
}

function Get-UsernameFromEmail {
    param([string]$EmailAddress)
    return $EmailAddress.Split('@')[0]
}

function New-EasyPassword {
    $words = @(
        "Tiger", "Panda", "Lion", "Bunny", "Star", "Moon", "Apple", "Ocean",
        "Happy", "Sunny", "Rainbow", "Cookie", "Puppy", "Rocket", "Dolphin", "Cloud"
    )
    $word = Get-Random -InputObject $words
    $number = Get-Random -Minimum 10 -Maximum 99

    return "$word$number@"
}

function Set-OneStudentPassword {
    param(
        [string]$TargetEmail,
        [string]$FixedPassword,   # if set, used as-is (no retry - it's what the admin typed)
        [int]$MaxAttempts = 5
    )

    $succeeded = $false
    $lastError = $null
    $usedPassword = $null

    for ($attempt = 1; $attempt -le $MaxAttempts -and -not $succeeded; $attempt++) {
        $usedPassword = if ($FixedPassword) { $FixedPassword } else { New-EasyPassword }

        try {
            Update-MgUser -UserId $TargetEmail -PasswordProfile @{
                Password                      = $usedPassword
                ForceChangePasswordNextSignIn = $false
            } -ErrorAction Stop
            $succeeded = $true
        } catch {
            $lastError = $_.Exception.Message
            if ($FixedPassword -or $attempt -ge $MaxAttempts) { break }
            Write-Warning "Attempt $attempt failed for $TargetEmail (retrying with a new password): $lastError"
        }
    }

    if ($succeeded) {
        Write-Host "$TargetEmail  ->  $usedPassword" -ForegroundColor Green
        return [pscustomobject]@{ Email = $TargetEmail; Password = $usedPassword; Status = 'Reset' }
    } else {
        Write-Warning "Failed for $TargetEmail after $attempt attempt(s): $lastError"
        return [pscustomobject]@{ Email = $TargetEmail; Password = ''; Status = "Failed: $lastError" }
    }
}

function Open-OutlookDraft {
    param([string]$To, [string]$Subject, [string]$Body)
    $mailto = "mailto:{0}?subject={1}&body={2}" -f `
        [uri]::EscapeDataString($To), `
        [uri]::EscapeDataString($Subject), `
        [uri]::EscapeDataString($Body)
    Start-Process $mailto
}

if (-not $Email -and -not $CsvPath) {
    throw "Provide either -Email (single account) or -CsvPath (bulk)."
}

Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All"

if ($Email) {
    $Email = Get-CleanEmail -Text $Email

    $mode = $null
    while ($mode -notin @('A', 'M')) {
        $mode = (Read-Host "`nType A to auto-generate an easy password, or M to type your own").Trim().ToUpper()
    }

    $fixedPassword = $null
    if ($mode -eq 'M') {
        while (-not $fixedPassword) {
            $fixedPassword = (Read-Host "Enter the new password for $Email").Trim()
            if (-not $fixedPassword) { Write-Warning "Password can't be blank." }
        }
    }

    $result = Set-OneStudentPassword -TargetEmail $Email -FixedPassword $fixedPassword

    if ($result.Status -eq 'Reset') {
        $subject = "Your Sherborne Qatar Office 365 Account Password"
        $body = @"
Hi,

Your Office 365 account password has been reset.

Email: $($result.Email)
Password: $($result.Password)

Please keep this password confidential.
"@
        Write-Host "`n----- Copy/paste email -----"
        Write-Host "To: $($result.Email)"
        Write-Host "Subject: $subject"
        Write-Host ""
        Write-Host $body
        Write-Host "-----------------------------"

        Write-Host "`nOpening Outlook with this email ready to send..."
        Open-OutlookDraft -To $result.Email -Subject $subject -Body $body
    }
} else {
    if (-not (Test-Path $CsvPath)) {
        throw "File not found: $CsvPath"
    }

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "Installing ImportExcel module (needed to read/write .xlsx files)..."
        Install-Module ImportExcel -Scope CurrentUser -Force
    }
    Import-Module ImportExcel

    $rows = Import-DataFile -Path $CsvPath
    if ($rows.Count -eq 0) { throw "No rows found in $CsvPath." }

    $emailCol = Get-EmailColumnName -Row $rows[0]

    Write-Host "`nFound $($rows.Count) account(s) in the file."

    $mode = $null
    while ($mode -notin @('A', 'M')) {
        $mode = (Read-Host "`nType A to auto-generate an easy password for each student, or M to set one password for everyone").Trim().ToUpper()
    }

    $sharedPassword = $null
    if ($mode -eq 'M') {
        while (-not $sharedPassword) {
            $sharedPassword = (Read-Host "Enter the password to apply to all $($rows.Count) account(s)").Trim()
            if (-not $sharedPassword) { Write-Warning "Password can't be blank." }
        }
    }

    # --- Prepare the output copy (always .xlsx, so results/highlighting match New-M365Students.ps1) ---
    $outDir  = Split-Path -Parent (Resolve-Path $CsvPath)
    $outName = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
    $outputPath = Join-Path $outDir "$outName - processed.xlsx"

    if ($CsvPath -match '\.xlsx$') {
        Copy-Item -Path $CsvPath -Destination $outputPath -Force
    } else {
        Remove-Item -Path $outputPath -ErrorAction SilentlyContinue
        $rows | Export-Excel -Path $outputPath -WorksheetName "Sheet1"
    }
    Write-Host "Working on a copy: $outputPath" -ForegroundColor Cyan

    # --- Reset passwords, tracking outcome per Excel row number (header = row 1) ---
    $results = @()
    $rowOutcomes = @{}   # sheetRow -> @{ Status; Username; Password }

    for ($i = 0; $i -lt $rows.Count; $i++) {
        $row = $rows[$i]
        $sheetRow = $i + 2

        $rowEmail = $row.$emailCol
        if (-not $rowEmail) { continue }
        $rowEmail = $rowEmail.ToString().Trim()
        if (-not $rowEmail) { continue }

        $result = Set-OneStudentPassword -TargetEmail $rowEmail -FixedPassword $sharedPassword
        $results += $result

        $rowOutcomes[$sheetRow] = @{
            Status   = $result.Status
            Username = Get-UsernameFromEmail -EmailAddress $rowEmail
            Password = $result.Password
        }
    }

    # --- Write Username / New Password columns + highlighting back into the copy ---
    $pkg = Open-ExcelPackage -Path $outputPath
    $ws = $pkg.Workbook.Worksheets[1]

    $headerCols = @{}
    for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
        $headerCols[$ws.Cells[1, $c].Text] = $c
    }
    $emailColIndex = $headerCols[$emailCol]

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
    $userColIndex = Get-OrAddColumn -Worksheet $ws -HeaderLookup $headerCols -Title "Username" -NextFreeCol ([ref]$nextFreeCol)
    $pwColIndex   = Get-OrAddColumn -Worksheet $ws -HeaderLookup $headerCols -Title "New Password" -NextFreeCol ([ref]$nextFreeCol)

    # --- Legend explaining the highlight colors ---
    $legendCol = Get-OrAddColumn -Worksheet $ws -HeaderLookup $headerCols -Title "Legend" -NextFreeCol ([ref]$nextFreeCol)
    $ws.Cells[1, $legendCol].Style.Font.Bold = $true

    $legendRows = @(
        @{ Text = "Green  = password reset successfully"; Color = [System.Drawing.Color]::LightGreen }
        @{ Text = "Orange = failed to reset (see console output)"; Color = [System.Drawing.Color]::Orange }
        @{ Text = "No fill = skipped (blank email)"; Color = $null }
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
    $ws.Column($legendCol).Width = 45

    foreach ($sheetRow in $rowOutcomes.Keys) {
        $outcome = $rowOutcomes[$sheetRow]
        $cell = $ws.Cells[$sheetRow, $emailColIndex]

        $ws.Cells[$sheetRow, $userColIndex].Value = $outcome.Username

        if ($outcome.Status -eq 'Reset') {
            $cell.Style.Fill.PatternType = 'Solid'
            $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen)
            $ws.Cells[$sheetRow, $pwColIndex].Value = $outcome.Password
        } else {
            $cell.Style.Fill.PatternType = 'Solid'
            $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Orange)
        }
    }

    Close-ExcelPackage $pkg

    $succeededCount = ($results | Where-Object { $_.Status -eq 'Reset' }).Count
    Write-Host "`n===================================="
    Write-Host "Reset:   $succeededCount of $($rows.Count)" -ForegroundColor Green
    Write-Host "Results written to: $outputPath"
    Write-Host "===================================="
}
