@echo off
setlocal
cd /d "%~dp0"
py -3 -m venv .venv-build
.venv-build\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv-build\Scripts\pyinstaller.exe --noconfirm --clean --onefile --windowed --name BuscadorDeASO src\buscador_aso.py
echo Executavel gerado em dist\BuscadorDeASO.exe
