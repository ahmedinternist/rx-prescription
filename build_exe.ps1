$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$versionMatch = Select-String -LiteralPath (Join-Path $projectRoot 'config.py') `
  -Pattern '^APP_VERSION\s*=\s*"([^"]+)"\s*$' | Select-Object -First 1
if (-not $versionMatch) {
  throw 'Could not read APP_VERSION from config.py.'
}
$version = $versionMatch.Matches[0].Groups[1].Value
$appName = "RxPrescription-v$version"
$specRoot = Join-Path $projectRoot 'build\spec'
New-Item -ItemType Directory -Force -Path $specRoot | Out-Null

$pythonCandidates = @()
if ($env:RX_PYTHON -and (Test-Path -LiteralPath $env:RX_PYTHON)) {
  $pythonCandidates += $env:RX_PYTHON
}
if (Test-Path -LiteralPath 'E:\PDF\python.exe') {
  $pythonCandidates += 'E:\PDF\python.exe'
}
$systemPython = Get-Command python -ErrorAction SilentlyContinue
if ($systemPython) {
  $pythonCandidates += $systemPython.Source
}

$python = $null
foreach ($candidate in ($pythonCandidates | Select-Object -Unique)) {
  & $candidate -c "import tkinter as tk; root=tk.Tk(); root.withdraw(); root.destroy()" 2>$null
  if ($LASTEXITCODE -eq 0) {
    $python = $candidate
    break
  }
}
if (-not $python) {
  throw 'No Python installation with a working Tcl/Tk runtime was found. Set RX_PYTHON to a suitable python.exe.'
}

Write-Host "Using $python"
Push-Location $projectRoot
try {
  $requiredImports = @(
    'PyInstaller', 'customtkinter', 'qrcode', 'reportlab', 'PIL', 'docx',
    'arabic_reshaper', 'bidi', 'openpyxl', 'cryptography',
    'xlrd', 'xlwt'
  )
  $importList = ($requiredImports | ForEach-Object { "'$_'" }) -join ', '
  $importCheck = "import importlib; [importlib.import_module(name) for name in [$importList]]"
  & $python -c $importCheck
  if ($LASTEXITCODE -ne 0) {
    Write-Host 'Installing missing build requirements...'
    & $python -m pip install --disable-pip-version-check -r requirements.txt pyinstaller
    if ($LASTEXITCODE -ne 0) { throw 'Could not install the build requirements.' }
  } else {
    Write-Host 'Build requirements are already available.'
  }
  & $python -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name $appName --workpath "build" --specpath $specRoot --distpath "dist" `
    --add-data "$projectRoot\data;data" main.py
  if ($LASTEXITCODE -ne 0) { throw 'PyInstaller did not complete successfully.' }
} finally {
  Pop-Location
}
Write-Host "Built dist\\$appName.exe"
