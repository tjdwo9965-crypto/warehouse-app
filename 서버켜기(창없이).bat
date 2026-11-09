@echo off
chcp 65001 > nul
cd /d "%~dp0"
netstat -ano | findstr /r /c:":8080 .*LISTENING" >nul && (
  echo 서버가 이미 켜져 있습니다.  http://localhost:8080
  timeout /t 5 >nul
  exit /b 0
)
echo 서버를 창 없이 켜는 중입니다...
wscript "%~dp0run_hidden.vbs"
set /a N=0
:wait
timeout /t 2 /nobreak >nul
netstat -ano | findstr /r /c:":8080 .*LISTENING" >nul && goto ok
set /a N+=1
if %N% lss 15 goto wait
echo.
echo [오류] 서버가 켜지지 않았습니다. logs\server.log 내용을 캡처해서 보내 주세요.
pause
exit /b 1
:ok
echo.
echo 서버가 켜졌습니다.  http://localhost:8080
echo 이 창은 닫아도 됩니다. 서버는 화면 없이 계속 돌아갑니다. (끌 때: 서버끄기.bat)
start "" http://localhost:8080
timeout /t 8 >nul
