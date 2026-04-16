param(
    [switch]$SkipBrowser,
    [switch]$SkipHealthCheck
)

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = 'python'
$proxyUrl = 'http://127.0.0.1:12803'
$healthUrl = "$proxyUrl/health"

Write-Host '[1/4] Package directory:' $scriptDir

if (-not $SkipBrowser) {
    Write-Host '[2/4] Starting Edge with CDP...'
    & powershell -ExecutionPolicy Bypass -File (Join-Path $scriptDir 'start_edge_cdp_and_check.ps1')
}
else {
    Write-Host '[2/4] Skipping browser startup.'
    Write-Host '      Expect an existing Edge CDP session on :9222 with Excel Online and the Copilot taskpane already open.'
}

Write-Host '[3/4] Checking Python dependencies...'
& $pythonExe -c "import fastapi, uvicorn, websocket" | Out-Null

if (-not $SkipHealthCheck) {
    Write-Host '[4/4] Current health probe before proxy start (may fail if proxy not running yet)...'
    try {
        Invoke-WebRequest -Uri $healthUrl -UseBasicParsing -TimeoutSec 5 | Out-Null
        Write-Host 'Proxy already responds on' $healthUrl
        return
    }
    catch {
        Write-Host 'No running proxy detected yet. Starting a fresh foreground proxy...'
    }
}

Write-Host ''
Write-Host 'Starting browser_attached_proxy.py on http://127.0.0.1:12803'
Write-Host 'Press Ctrl+C to stop.'
Write-Host ''

Set-Location $scriptDir
& $pythonExe (Join-Path $scriptDir 'browser_attached_proxy.py')
