@echo off
cd /d "%~dp0"
"C:\Users\RajKumar5\AppData\Local\Python\bin\python.exe" -m pip install -r requirements.txt -q
if "%~1"=="" (
  "C:\Users\RajKumar5\AppData\Local\Python\bin\python.exe" export_powerbi.py
) else (
  "C:\Users\RajKumar5\AppData\Local\Python\bin\python.exe" export_powerbi.py "%~1"
)
pause
