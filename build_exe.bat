@echo off
cd /d "%~dp0"
py -3 -m venv .venv-build
.venv-build\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv-build\Scripts\pyinstaller.exe --noconfirm --clean --onefile --windowed --name BackupUnicoPeople --paths src src\app.py
