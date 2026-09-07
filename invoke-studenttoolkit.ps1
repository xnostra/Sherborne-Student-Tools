<#
One-liner entry point for the Student Toolkit.
Downloads the toolkit scripts to the Desktop and launches the GUI.

    irm https://raw.githubusercontent.com/xnostra/Sherborne-Student-Tools/master/invoke-studenttoolkit.ps1 | iex
#>

$repoRaw = "https://raw.githubusercontent.com/xnostra/Sherborne-Student-Tools/master"
$targetDir = Join-Path $env:USERPROFILE "Desktop\StudentToolkit"
$cacheBuster = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()

New-Item -ItemType Directory -Path $targetDir -Force | Out-Null

$files = @("New-M365Students.ps1", "Add-M365StudentsByName.ps1", "Set-M365StudentPasswords.ps1", "StudentToolkit.ps1")
foreach ($file in $files) {
    Invoke-WebRequest -Uri "$repoRaw/$file?v=$cacheBuster" -OutFile (Join-Path $targetDir $file) -UseBasicParsing -Headers @{ 'Cache-Control' = 'no-cache' }
}

powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $targetDir "StudentToolkit.ps1")
