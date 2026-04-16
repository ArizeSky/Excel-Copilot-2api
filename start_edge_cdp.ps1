$edgePaths = @(
  "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
  "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
)

$edge = $edgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) {
  Write-Error "Edge not found. Please edit this script and set the correct path."
  exit 1
}

$userDataDir = Join-Path $env:TEMP "edge-cdp"
$debugUrl = "http://127.0.0.1:9222/json/version"
$jsonUrl = "http://127.0.0.1:9222/json"

Write-Host "Using Edge: $edge"
Write-Host "Debug port: 9222"
Write-Host "User data dir: $userDataDir"
Write-Host "Allow origins: *"

$ready = $false
try {
  $resp = Invoke-WebRequest -UseBasicParsing -Uri $debugUrl -TimeoutSec 2
  if ($resp.StatusCode -eq 200) {
    $ready = $true
    Write-Host "Reusing existing remote debugging endpoint: $debugUrl"
  }
} catch {
}

if (-not $ready) {
  Start-Process -FilePath $edge -ArgumentList @(
    "--remote-debugging-port=9222",
    "--remote-allow-origins=*",
    "--user-data-dir=$userDataDir"
  )

  Write-Host "Waiting for remote debugging endpoint..."
  for ($i = 0; $i -lt 20; $i++) {
    Start-Sleep -Milliseconds 500
    try {
      $resp = Invoke-WebRequest -UseBasicParsing -Uri $debugUrl -TimeoutSec 2
      if ($resp.StatusCode -eq 200) {
        $ready = $true
        break
      }
    } catch {
    }
  }
}

if (-not $ready) {
  Write-Error "Remote debugging endpoint did not come up: $debugUrl"
  exit 1
}

Write-Host "Remote debugging is ready: $debugUrl"
Start-Process $jsonUrl
Write-Host "Now open your Excel Online file in this Edge window, expand the Copilot taskpane, then run:"
Write-Host "  powershell -ExecutionPolicy Bypass -File .\start_personal_browser_proxy.ps1 -SkipBrowser"
