<#
.SYNOPSIS
    Registers a Windows scheduled task that starts the local add-in dev server at logon.

.DESCRIPTION
    The root manifest.json points at https://localhost:3000, so the sideloaded add-in only works
    while that dev server runs. This script creates (or removes) a per-user scheduled task that
    launches scripts\start-addin-server.ps1 hidden at every logon, restarts it if it crashes, and
    never times out. Administrator rights are not required.

.PARAMETER Port
    TCP port the dev server binds to. Must match the port used in manifest.json (3000 by default).

.PARAMETER DelaySeconds
    Delay after logon before the server starts, giving Windows time to finish booting.

.PARAMETER StartNow
    Start the task immediately after registering it.

.PARAMETER Uninstall
    Remove the scheduled task.

.PARAMETER Status
    Show the current task and port state.

.PARAMETER TaskName
    Name of the scheduled task.

.EXAMPLE
    npm run startup:install

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts\install-startup-task.ps1 -StartNow

.EXAMPLE
    npm run startup:uninstall
#>
[CmdletBinding(DefaultParameterSetName = "Install")]
param(
    [Parameter(ParameterSetName = "Install")]
    [ValidateRange(1, 65535)]
    [int]$Port = 3000,

    [Parameter(ParameterSetName = "Install")]
    [ValidateRange(0, 3600)]
    [int]$DelaySeconds = 30,

    [Parameter(ParameterSetName = "Install")]
    [switch]$StartNow,

    [Parameter(ParameterSetName = "Uninstall", Mandatory = $true)]
    [switch]$Uninstall,

    [Parameter(ParameterSetName = "Status", Mandatory = $true)]
    [switch]$Status,

    [string]$TaskName = "OutlookAutoReplyAutomater-DevServer"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$startScript = Join-Path $PSScriptRoot "start-addin-server.ps1"
$userId = "$env:USERDOMAIN\$env:USERNAME"

if (-not (Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue)) {
    throw "The ScheduledTasks PowerShell module is unavailable. This script requires Windows 8/Server 2012 or later."
}

function Get-ExistingTask {
    Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
}

function Test-PortListening {
    param([int]$PortNumber)

    try {
        $listening = Get-NetTCPConnection -LocalPort $PortNumber -State Listen -ErrorAction Stop
        return $null -ne $listening
    } catch {
        return $false
    }
}

if ($Status) {
    $task = Get-ExistingTask

    if (-not $task) {
        Write-Host "Scheduled task '$TaskName' is not registered."
    } else {
        $info = Get-ScheduledTaskInfo -TaskName $TaskName
        Write-Host "Scheduled task : $TaskName"
        Write-Host "State          : $($task.State)"
        Write-Host "Last run       : $($info.LastRunTime)"
        Write-Host "Last result    : $($info.LastTaskResult)"
        Write-Host "Next run       : $($info.NextRunTime)"
    }

    $portState = if (Test-PortListening -PortNumber $Port) { "listening" } else { "not listening" }
    Write-Host "Port $Port      : $portState"
    Write-Host "Log file       : $(Join-Path $env:LOCALAPPDATA 'outlook-auto-reply-automater\logs\dev-server.log')"
    return
}

if ($Uninstall) {
    if (-not (Get-ExistingTask)) {
        Write-Host "Scheduled task '$TaskName' is not registered. Nothing to remove."
        return
    }

    Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Removed scheduled task '$TaskName'."
    Write-Host "The dev server will no longer start automatically at logon."
    return
}

if (-not (Test-Path -LiteralPath $startScript)) {
    throw "Startup script not found at $startScript."
}

$powerShellPath = (Get-Command powershell.exe).Source
$arguments = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}" -Port {1}' -f $startScript, $Port

$action = New-ScheduledTaskAction -Execute $powerShellPath -Argument $arguments -WorkingDirectory $repoRoot
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $userId

if ($DelaySeconds -gt 0) {
    try {
        $trigger.Delay = "PT{0}S" -f $DelaySeconds
    } catch {
        Write-Warning "Could not apply the $DelaySeconds second logon delay; the task will start immediately at logon."
    }
}

$principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -ExecutionTimeLimit ([TimeSpan]::Zero)

$description = "Runs the Outlook Auto-Reply Automater local dev server on https://localhost:$Port so the sideloaded add-in keeps working."

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Settings $settings `
    -Description $description `
    -Force | Out-Null

Write-Host "Registered scheduled task '$TaskName'."
Write-Host "  Runs as      : $userId (at logon, delayed by $DelaySeconds s)"
Write-Host "  Serves       : https://localhost:$Port/src/taskpane/index.html"
Write-Host "  Script       : $startScript"
Write-Host "  Log file     : $(Join-Path $env:LOCALAPPDATA 'outlook-auto-reply-automater\logs\dev-server.log')"

if ($StartNow) {
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "Started the task now."
} else {
    Write-Host "Run 'npm run startup:start' to start it without logging off."
}
