@echo off
rem Hidden server start - run by run_hidden.vbs
cd /d "%~dp0"
netstat -ano | findstr /r /c:":8080 .*LISTENING" >nul && exit /b 0
set PY=
py -3 --version >nul 2>&1 && set PY=py -3
if not defined PY python --version >nul 2>&1 && set PY=python
if not defined PY for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%D\python.exe" set PY="%%D\python.exe"
if not defined PY exit /b 1
if not exist logs mkdir logs
set PYTHONIOENCODING=utf-8
%PY% -m pip install -r requirements.txt -q --disable-pip-version-check >> logs\server.log 2>&1
echo ===== %date% %time% server start ===== >> logs\server.log
%PY% app.py >> logs\server.log 2>&1
