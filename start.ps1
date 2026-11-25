# Probridge Land Bot - start the laptop server.
# Keep this window open while you use the bot (Excel or phone).
#   powershell -ExecutionPolicy Bypass -File .\start.ps1

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

if (-not (Test-Path "$root\.venv\Scripts\python.exe")) {
    Write-Host "Run setup.ps1 first." -ForegroundColor Yellow
    exit 1
}

# Make sure Ollama (the AI) is running
try {
    Invoke-RestMethod "http://localhost:11434/api/tags" -TimeoutSec 3 | Out-Null
} catch {
    Write-Host "Starting Ollama..."
    Start-Process "ollama" -ArgumentList "serve" -WindowStyle Hidden
    Start-Sleep -Seconds 3
}

Write-Host "Probridge Land Bot is starting. Close this window to stop it." -ForegroundColor Green
& "$root\.venv\Scripts\python.exe" "$root\server\server.py"
