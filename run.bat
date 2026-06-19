@echo off
cd /d "%~dp0"
"C:\Users\RajKumar5\AppData\Local\Python\bin\python.exe" -m pip install -r requirements.txt -q
"C:\Users\RajKumar5\AppData\Local\Python\bin\python.exe" app.py
