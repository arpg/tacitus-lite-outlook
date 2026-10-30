<#
.SYNOPSIS
  Register (or remove) the Windows scheduled task that copies new mail every 30 minutes.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File .\setup-task.ps1
  powershell -ExecutionPolicy Bypass -File .\setup-task.ps1 -NoAddresses
  powershell -ExecutionPolicy Bypass -File .\setup-task.ps1 -Remove
#>
param(
    [switch]$Remove,
    [switch]$NoAddresses,
    [int]$Minutes = 30,
    [string]$TaskName = "Outlook mail dump"
)

$ErrorActionPreference = "Stop"
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

if ($Remove) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed task '$TaskName'."
    exit 0
}

$py = (Get-Command py -ErrorAction SilentlyContinue)
if (-not $py) { $py = Get-Command python -ErrorAction SilentlyContinue }
if (-not $py) { throw "Python is not on PATH. Install it from python.org and tick 'Add python.exe to PATH'." }

$args = "tacitus_lite.py dump"
if ($NoAddresses) { $args += " --no-addresses" }

# pythonw has no console window; the task runs while you are logged in,
# since Outlook's automation interface needs your session.
$exe = $py.Source -replace "py\.exe$", "pythonw.exe" -replace "python\.exe$", "pythonw.exe"
if (-not (Test-Path $exe)) { $exe = $py.Source }

$action  = New-ScheduledTaskAction -Execute $exe -Argument $args -WorkingDirectory $here
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) `
           -RepetitionInterval (New-TimeSpan -Minutes $Minutes) -RepetitionDuration (New-TimeSpan -Days 3650)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
            -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null
Write-Host "Registered task '$TaskName': every $Minutes minutes, '$exe $args' in $here."
Write-Host "Check it in Task Scheduler, or run: py outlook_dump.py status"
