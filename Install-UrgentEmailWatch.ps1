<#
.SYNOPSIS
    Sets up urgent-email-watch v2 on Windows. Idempotent -- safe to re-run.

.DESCRIPTION
    This script:
      1. Checks that Python 3 is installed (prints install guidance if not).
      2. Copies config.example.json -> config.json if config.json does not
         exist yet (an existing config.json is never overwritten).
      3. Registers a Windows Scheduled Task that runs watchdog.py at user
         logon so the watchdog runs continuously.

    Re-running the script re-registers the task with current settings and
    leaves an existing config.json untouched.

    IMPORTANT: after filling in config.json you must run
        python watchdog.py --setup
    once. That signs in every Microsoft account interactively (device-code
    flow) and saves its refresh token. Gmail accounts need no interactive
    setup -- they authenticate with an App Password from config.json.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File Install-UrgentEmailWatch.ps1
#>

[CmdletBinding()]
param(
    [string]$TaskName = "UrgentEmailWatch"
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=== urgent-email-watch v2 installer ===" -ForegroundColor Cyan
Write-Host "Install directory: $scriptDir"
Write-Host ""

# --- 1. Check for Python 3 ---------------------------------------------------

function Find-Python {
    foreach ($candidate in @('python', 'python3', 'py')) {
        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($null -eq $cmd) { continue }
        try {
            if ($candidate -eq 'py') {
                $ver = & $candidate -3 --version 2>&1
            } else {
                $ver = & $candidate --version 2>&1
            }
        } catch { continue }
        if ($ver -match 'Python 3\.') {
            return [pscustomobject]@{ Exe = $cmd.Source; Version = "$ver".Trim() }
        }
    }
    return $null
}

$python = Find-Python
if ($null -eq $python) {
    Write-Host "Python 3 was not found on this system." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Install it with winget:"
    Write-Host "    winget install Python.Python.3.12" -ForegroundColor White
    Write-Host ""
    Write-Host "Or download the installer from:"
    Write-Host "    https://www.python.org/downloads/windows/" -ForegroundColor White
    Write-Host ""
    Write-Host "During install, tick 'Add python.exe to PATH'."
    Write-Host "Then re-run this installer."
    exit 1
}
Write-Host "Found Python: $($python.Version) at $($python.Exe)" -ForegroundColor Green
Write-Host "(No pip packages are needed -- watchdog.py is standard library only.)"

# --- 2. Copy the config template if missing ---------------------------------

function Copy-IfMissing {
    param([string]$Source, [string]$Dest)
    if (Test-Path -LiteralPath $Dest) {
        Write-Host "  exists, leaving as-is: $(Split-Path -Leaf $Dest)"
        return
    }
    if (-not (Test-Path -LiteralPath $Source)) {
        Write-Host "  WARNING: template not found: $Source" -ForegroundColor Yellow
        return
    }
    Copy-Item -LiteralPath $Source -Destination $Dest
    Write-Host "  created: $(Split-Path -Leaf $Dest)  (edit it before running)" -ForegroundColor Green
}

Write-Host ""
Write-Host "Setting up the config file..."
Copy-IfMissing -Source (Join-Path $scriptDir 'config.example.json') `
               -Dest   (Join-Path $scriptDir 'config.json')

# --- 3. Register the Scheduled Task -----------------------------------------

$watchdogPath = Join-Path $scriptDir 'watchdog.py'
if (-not (Test-Path -LiteralPath $watchdogPath)) {
    Write-Host "ERROR: watchdog.py not found at $watchdogPath" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Registering scheduled task '$TaskName'..."

# Remove any existing task with the same name so the script is idempotent.
$existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -ne $existing) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "  removed existing task"
}

# pythonw.exe runs without a console window if available; fall back to the
# resolved python executable otherwise.
$pythonw = Join-Path (Split-Path -Parent $python.Exe) 'pythonw.exe'
$runner = if (Test-Path -LiteralPath $pythonw) { $pythonw } else { $python.Exe }

$action = New-ScheduledTaskAction `
    -Execute $runner `
    -Argument "`"$watchdogPath`"" `
    -WorkingDirectory $scriptDir

$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME

# Keep it running: restart on failure, no execution time limit.
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -RestartCount 999 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -ExecutionTimeLimit ([TimeSpan]::Zero)

$principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME `
    -LogonType Interactive `
    -RunLevel Limited

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Description "Real-time urgent-email watchdog (urgent-email-watch v2)." | Out-Null

Write-Host "  task '$TaskName' registered (runs at logon)" -ForegroundColor Green

# --- Done --------------------------------------------------------------------

Write-Host ""
Write-Host "=== Setup complete ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Edit config.json -- notifications, the microsoft app client_id,"
Write-Host "     and the accounts list (one entry per mailbox, each with its own"
Write-Host "     provider, connection details, channels, and triggers)."
Write-Host ""
Write-Host "  2. RUN SETUP ONCE -- this is required:" -ForegroundColor Yellow
Write-Host "       $($python.Exe) `"$watchdogPath`" --setup" -ForegroundColor White
Write-Host "     This signs in every Microsoft account interactively (device-code"
Write-Host "     flow) and saves its refresh token. You will be shown a URL and a"
Write-Host "     short code to enter, once per Microsoft account. Gmail accounts"
Write-Host "     need no interactive setup -- they use an App Password from"
Write-Host "     config.json. The watchdog cannot poll a Microsoft account until"
Write-Host "     --setup has been run for it." -ForegroundColor Yellow
Write-Host ""
Write-Host "  3. Test it:"
Write-Host "       $($python.Exe) `"$watchdogPath`" --test" -ForegroundColor White
Write-Host ""
Write-Host "  4. Start it now (without waiting for the next logon):"
Write-Host "       Start-ScheduledTask -TaskName $TaskName" -ForegroundColor White
Write-Host ""
Write-Host "The watchdog will start automatically every time you log on."
