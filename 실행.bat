@echo off
chcp 65001 > nul
cd /d "%~dp0"
title 창고 재고 서버 (이 창을 닫으면 페이지가 꺼집니다)

set PY=
py -3 --version >nul 2>&1 && set PY=py -3
if not defined PY python --version >nul 2>&1 && set PY=python
if not defined PY for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%D\python.exe" set PY="%%D\python.exe"
if not defined PY (
  echo.
  echo [오류] Python이 설치되어 있지 않습니다.
  echo  1. 열리는 페이지에서 Python을 내려받아 설치하세요.
  echo  2. 설치 첫 화면에서 "Add python.exe to PATH" 를 꼭 체크하세요.
  echo  3. 설치 후 이 파일을 다시 실행하세요.
  start "" https://www.python.org/downloads/
  pause
  exit /b 1
)

echo 필요한 프로그램 설치 확인 중...
%PY% -m pip install -r requirements.txt -q --disable-pip-version-check
if errorlevel 1 (
  echo [오류] 패키지 설치에 실패했습니다. 인터넷 연결을 확인해 주세요.
  pause
  exit /b 1
)

rem 서버가 뜬 뒤 3초 후 브라우저 열기
start "" cmd /c "timeout /t 3 /nobreak >nul & start http://localhost:8080"
echo.
echo 서버 시작: http://localhost:8080   (이 창은 닫지 마세요)
%PY% app.py
echo.
echo [서버가 종료되었습니다] 위 메시지를 캡처해서 보내주세요.
pause
