<#
.SYNOPSIS
    Starts the local HTTPS dev server that hosts the Outlook Auto-Reply Automater task pane.

.DESCRIPTION
    Intended to be launched by the Windows scheduled task created by install-startup-task.ps1,
    but it can also be run manually. The script keeps running as long as the server is alive so
    the scheduled task reflects the real server state.

.PARAMETER Port
    TCP port the dev server must bind to. Must match the port used in manifest.json (3000 by default).

.PARAMETER SkipInstall
    Skip the automatic dependency installation when node_modules is missing.

.PARAMETER NoLogFile
    Write output only to the console instead of the log file.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts\start-addin-server.ps1
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 65535)]
    [int]$Port = 3000,

    [switch]$SkipInstall,

    [switch]$NoLogFile
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$logDirectory = Join-Path $env:LOCALAPPDATA "outlook-auto-reply-automater\logs"
$script:LogFile = Join-Path $logDirectory "dev-server.log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Write-Host $line

    if (-not $NoLogFile) {
        try {
            Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8
        } catch {
            # Never let logging failures stop the server.
        }
    }
}

function Initialize-Log {
    if ($NoLogFile) {
        return
    }

    if (-not (Test-Path -LiteralPath $logDirectory)) {
        New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
    }

    $maxSizeBytes = 5MB
    if (Test-Path -LiteralPath $script:LogFile) {
        $logItem = Get-Item -LiteralPath $script:LogFile
        if ($logItem.Length -gt $maxSizeBytes) {
            Move-Item -LiteralPath $script:LogFile -Destination "$($script:LogFile).1" -Force
        }
    }
}

function Test-PortListening {
    param([int]$PortNumber)

    try {
        $listening = Get-NetTCPConnection -LocalPort $PortNumber -State Listen -ErrorAction Stop
        return $null -ne $listening
    } catch [System.Management.Automation.CommandNotFoundException] {
        # Get-NetTCPConnection is unavailable; fall back to a connect probe below.
    } catch {
        return $false
    }

    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $client.Connect("127.0.0.1", $PortNumber)
        $connected = $client.Connected
        $client.Close()
        return $connected
    } catch {
        return $false
    }
}

Initialize-Log

Write-Log "Starting Outlook Auto-Reply Automater dev server."
Write-Log "Repository: $repoRoot"
Write-Log "Target URL: https://localhost:$Port/src/taskpane/index.html"

if (Test-PortListening -PortNumber $Port) {
    Write-Log "Port $Port is already in use. Assuming the dev server is already running; nothing to do." "WARN"
    exit 0
}

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Write-Log "Node.js was not found on PATH. Install Node.js 18 or later and re-run the startup task." "ERROR"
    exit 1
}

Set-Location -LiteralPath $repoRoot

$nodeModulesPath = Join-Path $repoRoot "node_modules"
if (-not (Test-Path -LiteralPath $nodeModulesPath)) {
    if ($SkipInstall) {
        Write-Log "node_modules is missing and -SkipInstall was set. Run 'npm ci --legacy-peer-deps' manually." "ERROR"
        exit 1
    }

    Write-Log "node_modules is missing. Installing dependencies with 'npm ci --legacy-peer-deps'."
    & $env:ComSpec /d /s /c "npm ci --legacy-peer-deps" 2>&1 | ForEach-Object { Write-Log $_ }

    if ($LASTEXITCODE -ne 0) {
        Write-Log "Dependency installation failed with exit code $LASTEXITCODE." "ERROR"
        exit $LASTEXITCODE
    }
}

& $env:ComSpec /d /s /c "npx --no-install office-addin-dev-certs verify" *>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Log "The local HTTPS development certificate is not trusted. Outlook will refuse to load the task pane." "WARN"
    Write-Log "Run 'npx office-addin-dev-certs install' once, then restart the task." "WARN"
} else {
    Write-Log "Local HTTPS development certificate is trusted."
}

Write-Log "Launching Vite on port $Port (strict port binding)."

# Vite writes progress information to stderr, so merged output must not be treated as a failure.
$ErrorActionPreference = "Continue"
& $env:ComSpec /d /s /c "npm run dev -- --port $Port --strictPort" 2>&1 | ForEach-Object { Write-Log $_ }
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
    Write-Log "Dev server exited with code $exitCode." "ERROR"
    exit $exitCode
}

Write-Log "Dev server stopped."
exit 0
