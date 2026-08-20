<#
Sherborne Qatar Student Tools - GUI launcher for the student onboarding script in this folder.
Double-click "Launch Student Toolkit.bat" instead of running this directly.
#>

try {
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/xnostra/Sherborne-Student-Tools/master/README.md" -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop | Out-Null
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show(
        "This tool could not verify access to its source repository and cannot continue.`n(github.com/xnostra/Sherborne-Student-Tools)",
        "Sherborne Qatar Student Tools",
        "OK", "Error") | Out-Null
    exit 1
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# --- Colors ---
$colorPrimary   = [System.Drawing.Color]::FromArgb(24, 55, 94)     # deep navy
$colorAccent    = [System.Drawing.Color]::FromArgb(41, 98, 168)    # brighter blue
$colorAccentHov = [System.Drawing.Color]::FromArgb(30, 78, 140)
$colorBg        = [System.Drawing.Color]::FromArgb(245, 247, 250)
$colorText      = [System.Drawing.Color]::FromArgb(60, 60, 60)

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Sherborne Qatar Student Tools"
$form.Size = New-Object System.Drawing.Size(440, 470)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = $colorBg
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# --- Header banner ---
$header = New-Object System.Windows.Forms.Panel
$header.Size = New-Object System.Drawing.Size(440, 90)
$header.Location = New-Object System.Drawing.Point(0, 0)
$header.BackColor = $colorPrimary
$form.Controls.Add($header)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "SHERBORNE QATAR"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.AutoSize = $true
$titleLabel.Location = New-Object System.Drawing.Point(30, 16)
$header.Controls.Add($titleLabel)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "Student Tools"
$subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11)
$subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(190, 210, 235)
$subtitleLabel.AutoSize = $true
$subtitleLabel.Location = New-Object System.Drawing.Point(32, 52)
$header.Controls.Add($subtitleLabel)

function New-ActionButton {
    param([string]$Text, [int]$Y)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = $Text
    $btn.Size = New-Object System.Drawing.Size(360, 56)
    $btn.Location = New-Object System.Drawing.Point(30, $Y)
    $btn.Font = New-Object System.Drawing.Font("Segoe UI", 10.5)
    $btn.ForeColor = [System.Drawing.Color]::White
    $btn.BackColor = $colorAccent
    $btn.FlatStyle = "Flat"
    $btn.FlatAppearance.BorderSize = 0
    $btn.FlatAppearance.MouseOverBackColor = $colorAccentHov
    $btn.TextAlign = "MiddleCenter"
    $btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $form.Controls.Add($btn)
    return $btn
}

# --- Button 1: New students from MIS export ---
$btnNewStudents = New-ActionButton -Text "BULK - Add New Students (XLSX)" -Y 120
$btnNewStudents.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.InitialDirectory = $scriptDir
    $ofd.Filter = "Excel files (*.xlsx)|*.xlsx"
    $ofd.Title = "Select the student MIS export"
    if ($ofd.ShowDialog() -ne "OK") { return }

    $location = [Microsoft.VisualBasic.Interaction]::InputBox("Usage location (e.g. QA):", "Usage Location", "QA")
    if (-not $location) { return }

    $cmd = "& '$scriptDir\New-M365Students.ps1' -XlsxPath '$($ofd.FileName)' -UsageLocation '$location'"
    Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", $cmd
})

# --- Button 2: Bulk reset passwords (Prep / SEN students) ---
$btnBulkResetPw = New-ActionButton -Text "BULK - Reset Passwords (Prep / SEN)" -Y 190
$btnBulkResetPw.Add_Click({
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.InitialDirectory = $scriptDir
    $ofd.Filter = "CSV / Excel files (*.csv;*.xlsx)|*.csv;*.xlsx"
    $ofd.Title = "Select a file with an Email (or Pupil Email Address) column"
    if ($ofd.ShowDialog() -ne "OK") { return }

    $cmd = "& '$scriptDir\Set-M365StudentPasswords.ps1' -CsvPath '$($ofd.FileName)'"
    Start-Process powershell.exe -ArgumentList "-NoExit", "-Command", $cmd
})

# --- Footer ---
$note = New-Object System.Windows.Forms.Label
$note.Text = "Each action opens a console window - sign in there when prompted."
$note.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$note.AutoSize = $true
$note.ForeColor = [System.Drawing.Color]::Gray
$note.Location = New-Object System.Drawing.Point(30, 370)
$form.Controls.Add($note)

$version = New-Object System.Windows.Forms.Label
$version.Text = "Sherborne Qatar Student Tools"
$version.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$version.AutoSize = $true
$version.ForeColor = [System.Drawing.Color]::LightGray
$version.Location = New-Object System.Drawing.Point(30, 395)
$form.Controls.Add($version)

[void]$form.ShowDialog()
