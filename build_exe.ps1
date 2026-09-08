$ErrorActionPreference = 'Stop'

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
& $python -m pip install -r requirements.txt pyinstaller
if ($LASTEXITCODE -ne 0) { throw 'Could not install the build requirements.' }
& $python -m PyInstaller --noconfirm --clean --onefile --windowed `
  --name RxPrescription --add-data "data;data" --add-data "viewer.html;." main.py
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller did not complete successfully.' }
Write-Host "Built dist\\RxPrescription.exe"
