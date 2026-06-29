$ErrorActionPreference = "Stop"

# Always run from the project root (folder containing this script)
Set-Location -Path $PSScriptRoot

# Use the venv Python if present, otherwise the system Python
$venvPython = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (Test-Path $venvPython) {
    $python = $venvPython
} else {
    $python = "python"
}

# Ensure Outlook is running before we automate it
if (-not (Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)) {
    Start-Process "outlook.exe"
    Start-Sleep -Seconds 20  # give Outlook time to load the profile
}

& $python -m src.main
exit $LASTEXITCODE
