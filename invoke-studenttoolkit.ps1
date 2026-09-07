<#
One-liner entry point for the Student Toolkit.
Downloads the toolkit scripts to the Desktop and launches the GUI.

    irm https://raw.githubusercontent.com/xnostra/Sherborne-Student-Tools/master/invoke-studenttoolkit.ps1 | iex
#>

$repoApi = "https://api.github.com/repos/xnostra/Sherborne-Student-Tools"
$targetDir = Join-Path $env:USERPROFILE "Desktop\StudentToolkit"

try {
    $latestCommit = Invoke-RestMethod -Uri "$repoApi/commits/master" -Headers @{ 'User-Agent' = 'Sherborne-Student-Toolkit' } -ErrorAction Stop
    $commitSha = $latestCommit.sha
    if ($commitSha -notmatch '^[0-9a-f]{40}$') { throw 'GitHub did not return a valid commit identifier.' }
} catch {
    throw "Could not determine the latest toolkit version. $($_.Exception.Message)"
}

$repoRaw = "https://raw.githubusercontent.com/xnostra/Sherborne-Student-Tools/$commitSha"

New-Item -ItemType Directory -Path $targetDir -Force | Out-Null

$files = @("New-M365Students.ps1", "Add-M365StudentsByName.ps1", "Set-M365StudentPasswords.ps1", "StudentToolkit.ps1")
$downloads = New-Object 'System.Collections.Generic.List[object]'
try {
    foreach ($file in $files) {
        $destination = Join-Path $targetDir $file
        $temporaryDownload = "$destination.$commitSha.download"
        Invoke-WebRequest -Uri "$repoRaw/$file" -OutFile $temporaryDownload -UseBasicParsing -ErrorAction Stop
        $downloads.Add([pscustomobject]@{ Temporary = $temporaryDownload; Destination = $destination })
    }
    foreach ($download in $downloads) {
        Move-Item -LiteralPath $download.Temporary -Destination $download.Destination -Force
    }
} catch {
    foreach ($download in $downloads) {
        if (Test-Path -LiteralPath $download.Temporary) {
            Remove-Item -LiteralPath $download.Temporary -Force
        }
    }
    if ($temporaryDownload -and (Test-Path -LiteralPath $temporaryDownload)) {
        Remove-Item -LiteralPath $temporaryDownload -Force
    }
    throw "Toolkit update failed. Existing Desktop files were not launched. $($_.Exception.Message)"
}

powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $targetDir "StudentToolkit.ps1")
