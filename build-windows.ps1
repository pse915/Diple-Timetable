$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
python -m pip install --upgrade pip
python -m pip install pyinstaller
python -m unittest discover -s tests -v
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "Diple-Timetable" --add-data "templates;templates" main.py
Write-Host "EXE generated at: $PSScriptRoot\dist\Diple-Timetable.exe"
