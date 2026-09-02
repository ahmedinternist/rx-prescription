$ErrorActionPreference = 'Stop'

$python = Get-Command python -ErrorAction Stop
& $python.Source -c "import tkinter as tk; root=tk.Tk(); root.withdraw(); root.destroy()"
& $python.Source -m pip install -r requirements.txt pyinstaller
& $python.Source -m PyInstaller --noconfirm --clean --onefile --windowed `
  --name RxPrescription --add-data "data;data" --add-data "viewer.html;." main.py
Write-Host "Built dist\\RxPrescription.exe"
