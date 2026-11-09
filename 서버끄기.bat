@echo off
chcp 65001 > nul
cd /d "%~dp0"
netstat -ano | findstr /r /c:":8080 .*LISTENING" >nul || (
  echo 서버가 꺼져 있습니다.
  timeout /t 5 >nul
  exit /b 0
)
echo 서버를 끄는 중입니다...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-NetTCPConnection -LocalPort 8080 -State Listen -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique | ForEach-Object { $p = Get-CimInstance Win32_Process -Filter ('ProcessId=' + $_); if ($p) { $pp = Get-Process -Id $p.ParentProcessId -ErrorAction SilentlyContinue; Stop-Process -Id $_ -Force; if ($pp -and $pp.ProcessName -eq 'py') { Stop-Process -Id $pp.Id -Force -ErrorAction SilentlyContinue } } }"
timeout /t 2 /nobreak >nul
netstat -ano | findstr /r /c:":8080 .*LISTENING" >nul && (
  echo [오류] 서버를 끄지 못했습니다. PC를 다시 시작하거나 작업 관리자에서 python 을 종료해 주세요.
  pause
  exit /b 1
)
echo 서버를 껐습니다. 다시 켜려면 서버켜기(창없이).bat 또는 실행.bat 을 실행하세요.
timeout /t 5 >nul
