# Probridge Land Bot - one-time setup.
# Run in PowerShell from this folder:   powershell -ExecutionPolicy Bypass -File .\setup.ps1

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root
$config = Get-Content "$root\config.json" -Raw | ConvertFrom-Json

function Step($n, $text) { Write-Host "`n[$n] $text" -ForegroundColor Cyan }
function Ok($text) { Write-Host "    OK  $text" -ForegroundColor Green }
function Warn($text) { Write-Host "    !!  $text" -ForegroundColor Yellow }

# 1. Python packages ---------------------------------------------------------
Step 1 "Python packages (FastAPI + Whisper)"
if (-not (Test-Path "$root\.venv\Scripts\python.exe")) {
    py -3.13 -m venv .venv
    if (-not $?) { python -m venv .venv }
}
& "$root\.venv\Scripts\python.exe" -m pip install --quiet --disable-pip-version-check -r server\requirements.txt
Ok "Python packages installed"

# 2. AI model -----------------------------------------------------------------
Step 2 "AI model ($($config.ollama_model), about 4.9 GB download, one time)"
if (-not (Get-Command ollama -ErrorAction SilentlyContinue)) {
    Warn "Ollama not found. Install it from https://ollama.com, then run setup.ps1 again."
} else {
    ollama pull $config.ollama_model
    Ok "AI model ready"
}

# 3. Voice model -------------------------------------------------------------
Step 3 "Voice model (Whisper '$($config.whisper_model)', about 500 MB, one time)"
& "$root\.venv\Scripts\python.exe" -c "from faster_whisper import WhisperModel; WhisperModel('$($config.whisper_model)', device='cpu', compute_type='int8'); print('    OK  Voice model ready')"

# 4. HTTPS certificate for the Excel add-in ----------------------------------
Step 4 "HTTPS certificate for localhost (Excel add-ins need https)"
$cert = Join-Path $HOME ".office-addin-dev-certs\localhost.crt"
if (Test-Path $cert) {
    Ok "Certificate already installed"
} else {
    Write-Host "    Windows will ask you to trust a 'Developer CA for Microsoft Office Add-ins'. Click Yes."
    npx --yes office-addin-dev-certs install
    Ok "Certificate installed"
}

# 5. Register the add-in in Excel desktop -----------------------------------
Step 5 "Add 'Land Bot' button to Excel"
$manifest = Join-Path $root "addin\manifest.xml"
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
$id = ([xml](Get-Content $manifest)).OfficeApp.Id
Set-ItemProperty -Path $key -Name $id -Value $manifest
Ok "Registered. Close and reopen Excel to see Home > Land Bot."

Write-Host "`nSetup finished. Now run:  .\start.ps1" -ForegroundColor Green
