# `run.ps1` and Task Scheduler — Running & automating

## What `run.ps1` does
A small PowerShell launcher that runs the app reliably on the remote PC:
1. `cd`s to the project folder (so relative paths and the venv resolve).
2. Picks the virtual-environment Python if present, else system Python.
3. Makes sure **Outlook is running** (starts it and waits if not).
4. Runs `python -m src.main` and passes through its exit code.

## Why it exists
Task Scheduler needs a single, dependable command to call. Doing the "make sure
Outlook is up, use the right Python, run from the right folder" steps in one
script means the scheduled task is just "run this file" — and the same script
works when you double-run it by hand.

## Why it is built this way

### `Set-Location $PSScriptRoot`
Scheduled tasks often start in `C:\Windows\System32`. Anchoring to the script's
own folder guarantees the project is the working directory.

### Prefer the venv Python
```powershell
$venvPython = ".venv\Scripts\python.exe"
```
Using the project's virtual environment ensures the exact installed versions of
openpyxl and pywin32 are used, not whatever global Python happens to be on PATH.

### Ensure Outlook is open first
```powershell
if (-not (Get-Process -Name "OUTLOOK" ...)) { Start-Process "outlook.exe"; Start-Sleep 20 }
```
COM automation needs Outlook actually running with your profile loaded. Starting
it (and pausing for it to load) avoids "Outlook not running" failures on a fresh
session.

### Pass through the exit code
```powershell
exit $LASTEXITCODE
```
So Task Scheduler can see success/failure (`0`/`1`/`2` from `main.py`).

## Why automation uses RDP + Task Scheduler (not remote calls)
Outlook COM only works in an **interactive desktop session** on the same
machine. So:
- Run the task **as the logged-in user**, "only when logged on", interactive.
- Avoid "run whether logged on or not" and the `SYSTEM` account — Outlook COM
  fails there because there's no real desktop/Outlook session.
- Connect to the remote PC with **Remote Desktop** to set things up and to keep
  an interactive session for scheduled runs.

## Registering the scheduled task (run once on the remote PC)
```powershell
$action = New-ScheduledTaskAction -Execute "powershell.exe" `
  -Argument "-NoProfile -ExecutionPolicy Bypass -File `"C:\path\to\email-automation\run.ps1`""
$trigger = New-ScheduledTaskTrigger -Daily -At 8:00AM
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
  -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName "OutlookDraftAutomation" `
  -Action $action -Trigger $trigger -Principal $principal `
  -Description "Generates Outlook drafts from the Excel file (never sends)"
```

Run on demand:
```powershell
Start-ScheduledTask -TaskName "OutlookDraftAutomation"
```

## Manual run (recommended before scheduling)
```powershell
cd C:\path\to\email-automation
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m src.main
```
Check the Drafts folder and the newest file in `logs/`.
