<#
Bulk resets passwords for a list of student accounts - built for Prep and SEN ("special child")
classes where a whole group needs a fresh, easy password at once.
Run manually by a Global Admin / User Administrator. Requires the Microsoft.Graph module:
    Install-Module Microsoft.Graph -Scope CurrentUser
(.xlsx input also requires the ImportExcel module - installed automatically if missing)

The file just needs a column of email addresses - either "Email" or "Pupil Email Address"
(case-insensitive) is accepted, so you can point it straight at a class export.

When you run it, it asks:
    A = auto-generate an easy password for each student (a different one per account)
    M = you type ONE password that gets applied to every account in the file

Usage:
    .\Set-M365StudentPasswords.ps1 -CsvPath ".\class-list.csv"
    .\Set-M365StudentPasswords.ps1 -CsvPath ".\class-list.xlsx"

Results (email + password used) are printed to the console and also saved next to your input
file as "<file> - passwords.csv" so they're easy to hand out.
#>

param(
    [Parameter(Mandatory = $true)]
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

function New-EasyPassword {
    $words = @(
        "Tiger", "Panda", "Lion", "Bunny", "Star", "Moon", "Apple", "Ocean",
        "Happy", "Sunny", "Rainbow", "Cookie", "Puppy", "Rocket", "Dolphin", "Cloud"
    )
    $word = Get-Random -InputObject $words
    $number = Get-Random -Minimum 10 -Maximum 99

    return "$word$number@"
}

if (-not (Test-Path $CsvPath)) {
    throw "File not found: $CsvPath"
}

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

Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All"

$results = @()
foreach ($row in $rows) {
    $email = $row.$emailCol
    if (-not $email) { continue }
    $email = $email.ToString().Trim()
    if (-not $email) { continue }

    $password = if ($mode -eq 'M') { $sharedPassword } else { New-EasyPassword }

    try {
        Update-MgUser -UserId $email -PasswordProfile @{
            Password                      = $password
            ForceChangePasswordNextSignIn = $false
        }
        Write-Host "$email  ->  $password" -ForegroundColor Green
        $results += [pscustomobject]@{ Email = $email; Password = $password; Status = 'Reset' }
    } catch {
        Write-Warning "Failed for $email : $($_.Exception.Message)"
        $results += [pscustomobject]@{ Email = $email; Password = ''; Status = "Failed: $($_.Exception.Message)" }
    }
}

$outputPath = Join-Path (Split-Path -Parent (Resolve-Path $CsvPath)) "$([System.IO.Path]::GetFileNameWithoutExtension($CsvPath)) - passwords.csv"
$results | Export-Csv -Path $outputPath -NoTypeInformation

$succeeded = ($results | Where-Object { $_.Status -eq 'Reset' }).Count
Write-Host "`n===================================="
Write-Host "Reset:   $succeeded of $($rows.Count)" -ForegroundColor Green
Write-Host "Results written to: $outputPath"
Write-Host "===================================="
