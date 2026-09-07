<#
Creates one or more student accounts from pasted full names (one name per line).
Every name is checked against Microsoft 365 before an account is created. Exact
display-name matches are left alone, so this is safe for quick additions between MIS exports.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$NamesBase64,

    [Parameter(Mandatory = $true)]
    [string]$UsageLocation,

    [string]$Domain = 'sherborneqatar.org'
)

function Get-CleanLetters {
    param([string]$Text)
    return ($Text -replace '[^a-zA-Z]', '').ToLower()
}

function Get-AssignableSeatCount {
    param($SubscribedSku)
    return ([int]$SubscribedSku.PrepaidUnits.Enabled + [int]$SubscribedSku.PrepaidUnits.Warning - [int]$SubscribedSku.ConsumedUnits)
}

function Test-UpnTaken {
    param([string]$Upn, [System.Collections.Generic.HashSet[string]]$UsedThisRun)

    if ($UsedThisRun.Contains($Upn)) { return $true }
    try { return $null -ne (Get-MgUser -UserId $Upn -ErrorAction Stop) }
    catch { return $false }
}

function New-StudentUpn {
    param([string]$FullName, [string]$Number, [System.Collections.Generic.HashSet[string]]$UsedThisRun)

    $letters = Get-CleanLetters $FullName
    if ($letters.Length -eq 0) { throw "'$FullName' has no usable letters for an email address." }

    $startLength = [Math]::Min(4, $letters.Length)
    for ($length = $startLength; $length -le $letters.Length; $length++) {
        $candidate = "$($letters.Substring(0, $length))$Number@$Domain"
        if (-not (Test-UpnTaken -Upn $candidate -UsedThisRun $UsedThisRun)) { return $candidate }
    }
    throw "Could not find a free email address for '$FullName'."
}

function New-StudentPassword {
    param([string]$FullName)
    $parts = @($FullName -split '\s+' | Where-Object { $_ })
    $first = Get-CleanLetters $parts[0]
    $last = Get-CleanLetters $parts[$parts.Count - 1]
    return "$($first.Substring(0, 1))$($last.Substring(0, 1))student@123"
}

try {
    $namesText = [Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($NamesBase64))
} catch {
    throw 'The pasted names could not be read.'
}

$names = @($namesText -split "`r?`n" | ForEach-Object { ($_ -replace '^\s*[-*•]\s*', '').Trim() } | Where-Object { $_ })
if ($names.Count -eq 0) { throw 'Paste at least one full name.' }

$uniqueNames = New-Object 'System.Collections.Generic.List[string]'
$seenNames = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($name in $names) {
    $key = ($name -replace '\s+', ' ').Trim().ToLower()
    if ($seenNames.Add($key)) { $uniqueNames.Add(($name -replace '\s+', ' ').Trim()) }
    else { Write-Warning "Duplicate pasted name '$name' ignored." }
}

Write-Host "`nPasted students to check:" -ForegroundColor Cyan
for ($i = 0; $i -lt $uniqueNames.Count; $i++) { Write-Host "  [$($i + 1)] $($uniqueNames[$i])" }

Connect-MgGraph -Scopes 'User.ReadWrite.All', 'Directory.ReadWrite.All', 'Organization.Read.All'

$studentA5SkuPartNumbers = @('M365EDU_A5_STUUSEBNFT', 'M365EDU_A5_STUDENT', 'M365EDU_A5_NOPSTNCONF_STUUSEBNFT', 'M365EDU_A5_NOPSTNCONF_STUDENT', 'ENTERPRISEPREMIUM_STUUSEBNFT', 'ENTERPRISEPREMIUM_STUDENT', 'ENTERPRISEPREMIUM_NOPSTNCONF_STUUSEBNFT', 'ENTERPRISEPREMIUM_NOPSTNCONF_STUDENT')
$skus = @(Get-MgSubscribedSku -All | Where-Object { (Get-AssignableSeatCount $_) -gt 0 })
if (-not $skus) { throw 'No licenses with available seats were found in this tenant.' }

Write-Host "`nAvailable licenses in this tenant:"
for ($i = 0; $i -lt $skus.Count; $i++) {
    $label = if ($skus[$i].SkuPartNumber -eq 'M365EDU_A5_STUUSEBNFT') { 'Microsoft 365 A5 for Students (Student Use Benefit)' } else { $skus[$i].SkuPartNumber }
    Write-Host "  [$i] $label  (available: $(Get-AssignableSeatCount $skus[$i]))"
}
$recommended = @(0..($skus.Count - 1) | Where-Object { $skus[$_].SkuPartNumber -eq 'M365EDU_A5_STUUSEBNFT' })
if (-not $recommended) { $recommended = @(0..($skus.Count - 1) | Where-Object { $studentA5SkuPartNumbers -contains $skus[$_].SkuPartNumber }) }
if ($recommended) { Write-Host "`nRecommended: [$($recommended[0])] Microsoft 365 A5 for Students" -ForegroundColor Yellow }
$choice = Read-Host 'Enter the license number (press Enter for the recommended license)'
if ([string]::IsNullOrWhiteSpace($choice) -and $recommended) { $choice = $recommended[0] }
$sku = $skus[[int]$choice]
if (-not $sku) { throw 'Invalid license selection.' }
$available = Get-AssignableSeatCount $sku

$emailNumber = Read-Host 'What should be appended to new email addresses? (e.g. 26 or 26q)'
if ($emailNumber -notmatch '^[a-zA-Z0-9]+$') { throw 'The email suffix can only contain letters and digits.' }

$passwordMode = ''
while ($passwordMode -notin @('A', 'M')) { $passwordMode = (Read-Host 'Type A for automatic passwords, or M for one password for all new students').Trim().ToUpper() }
$sharedPassword = $null
if ($passwordMode -eq 'M') {
    while (-not $sharedPassword) {
        $sharedPassword = (Read-Host 'Enter the password to apply to every new student').Trim()
        if (-not $sharedPassword) { Write-Warning "Password can't be blank." }
    }
}

$confirm = (Read-Host "Create accounts for the pasted names that do not already exist? Type YES to continue").Trim().ToUpper()
if ($confirm -ne 'YES') { Write-Host 'No accounts were created.' -ForegroundColor Yellow; exit 0 }

$usedUpns = New-Object 'System.Collections.Generic.HashSet[string]'
$results = @()
foreach ($fullName in $uniqueNames) {
    $escapedName = $fullName.Replace("'", "''")
    $matches = @(Get-MgUser -Filter "displayName eq '$escapedName'" -Property DisplayName,UserPrincipalName -All -ErrorAction SilentlyContinue)
    if ($matches.Count -gt 0) {
        $results += [pscustomobject]@{ Name = $fullName; Email = ($matches[0].UserPrincipalName); Status = 'Already exists - skipped' }
        Write-Host "Existing account found: $fullName ($($matches[0].UserPrincipalName))" -ForegroundColor Green
        continue
    }

    $parts = @($fullName -split '\s+' | Where-Object { $_ })
    if ($parts.Count -lt 2) {
        $results += [pscustomobject]@{ Name = $fullName; Email = ''; Status = 'Not created - enter at least first and last name' }
        Write-Warning "Skipped '$fullName': enter at least a first and last name."
        continue
    }

    if ($available -le 0) {
        $results += [pscustomobject]@{ Name = $fullName; Email = ''; Status = 'Not created - no selected license seats remain' }
        Write-Warning "Skipped '$fullName': no '$($sku.SkuPartNumber)' license seats remain."
        continue
    }

    $upn = New-StudentUpn -FullName $fullName -Number $emailNumber -UsedThisRun $usedUpns
    $usedUpns.Add($upn) | Out-Null
    $password = if ($sharedPassword) { $sharedPassword } else { New-StudentPassword -FullName $fullName }
    try {
        $newUser = New-MgUser -DisplayName $fullName -GivenName $parts[0] -Surname (($parts | Select-Object -Skip 1) -join ' ') -UserPrincipalName $upn -MailNickname $upn.Split('@')[0] -JobTitle 'Student' -UsageLocation $UsageLocation -AccountEnabled:$true -PasswordProfile @{ Password = $password; ForceChangePasswordNextSignIn = $false } -ErrorAction Stop
    } catch {
        $results += [pscustomobject]@{ Name = $fullName; Email = ''; Status = "Failed to create: $($_.Exception.Message)" }
        Write-Warning "Failed to create '$fullName': $($_.Exception.Message)"
        continue
    }
    try {
        Set-MgUserLicense -UserId $newUser.Id -AddLicenses @{ SkuId = $sku.SkuId } -RemoveLicenses @() -ErrorAction Stop
        $available--
        $results += [pscustomobject]@{ Name = $fullName; Email = $upn; Status = "Created and licensed (password: $password)" }
        Write-Host "Created and licensed: $fullName ($upn)" -ForegroundColor Yellow
    } catch {
        $results += [pscustomobject]@{ Name = $fullName; Email = $upn; Status = "Created but license failed: $($_.Exception.Message)" }
        Write-Warning "Created '$fullName' but could not assign the license: $($_.Exception.Message)"
    }
}

Write-Host "`nResults:" -ForegroundColor Cyan
$results | Format-Table -AutoSize
