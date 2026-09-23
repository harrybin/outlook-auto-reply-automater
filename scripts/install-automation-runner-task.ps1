[CmdletBinding(DefaultParameterSetName = "Install")]
param(
    [Parameter(ParameterSetName = "Install")]
    [ValidateRange(30, 86400)]
    [int]$IntervalSeconds = 60,
    [Parameter(ParameterSetName = "Install")]
    [switch]$StartNow,
    [Parameter(ParameterSetName = "Uninstall", Mandatory = $true)]
    [switch]$Uninstall,
    [Parameter(ParameterSetName = "Status", Mandatory = $true)]
    [switch]$Status,
    [string]$TaskName = "OutlookAutoReplyAutomater-Runner"
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$runner = Join-Path $PSScriptRoot "automation-runner.mjs"
$userId = "$env:USERDOMAIN\$env:USERNAME"

if (-not (Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue)) {
    throw "The ScheduledTasks PowerShell module is unavailable."
}

if ($Status) {
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($task) {
        $info = Get-ScheduledTaskInfo -TaskName $TaskName
        Write-Host "Scheduled task : $TaskName"
        Write-Host "State          : $($task.State)"
        Write-Host "Last run       : $($info.LastRunTime)"
        Write-Host "Last result    : $($info.LastTaskResult)"
    } else {
        Write-Host "Scheduled task '$TaskName' is not registered."
    }
    return
}

if ($Uninstall) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task '$TaskName'."
    return
}

if (-not (Test-Path -LiteralPath $runner)) {
    throw "Runner not found at $runner."
}

$node = (Get-Command node.exe -ErrorAction Stop).Source
$arguments = '"{0}" --interval {1}' -f $runner, $IntervalSeconds
$action = New-ScheduledTaskAction -Execute $node -Argument $arguments -WorkingDirectory $repoRoot
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $userId
$principal = New-ScheduledTaskPrincipal -UserId $userId -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit ([TimeSpan]::Zero)

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Runs Outlook Auto-Reply Automater continuously in the signed-in Windows session." -Force | Out-Null
Write-Host "Registered scheduled task '$TaskName' at user logon."
if ($StartNow) {
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "Started the task. Complete the device-code sign-in shown by the runner."
}
