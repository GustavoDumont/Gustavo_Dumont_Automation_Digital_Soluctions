@echo off
cd /d "%~dp0"
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\pyinstaller.exe --noconfirm --clean --onefile --windowed --name PendenciasNACT egot_automacao.py
.venv\Scripts\pyinstaller.exe --noconfirm --clean --onefile --windowed --name OrganizadorNACT organizador_arquivos_sondagem.py
.venv\Scripts\pyinstaller.exe --noconfirm --clean --onefile --windowed --name MapeamentoFOPAG mapeamento_fopag.py
