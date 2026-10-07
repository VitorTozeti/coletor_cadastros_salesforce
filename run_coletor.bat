@echo off
rem Roda o coletor de cadastros Salesforce -> PowerApp_Clientes.xlsx. Agendar 1x/dia no schtasks.
cd /d "%~dp0"
for /f "tokens=1,* delims==" %%a in ('findstr /b "XLSX_PATH=" .env') do set "XLSX_PATH=%%b"
if not exist .venv ( python -m venv .venv )
call .venv\Scripts\activate.bat
pip install -q -r requirements.txt
echo [%date% %time%] inicio >> coletor.log
python sf_clientes_para_excel.py --xlsx "%XLSX_PATH%" >> coletor.log 2>&1
echo [%date% %time%] fim (codigo %errorlevel%) >> coletor.log
exit /b %errorlevel%
