@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo 덮개를 닫으면 모니터는 꺼지고, 서버와 인터넷 연결은 계속 유지되도록 설정합니다.
echo.
rem 1) 화면이 꺼진 대기 상태에서도 네트워크 연결 유지 (전원 연결 시)
powercfg /setacvalueindex SCHEME_CURRENT SUB_NONE CONNECTIVITYINSTANDBY 1 >nul 2>&1
powercfg /setactive SCHEME_CURRENT
rem 2) 덮개 감시 프로그램 다시 시작 (새 버전 적용)
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'powershell.exe' -and $_.ProcessId -ne $PID -and $_.CommandLine -like '*-File*lid_watcher.ps1*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"
rem 3) 서버 다시 시작 (서버도 화면이 꺼져도 깨어 있도록 바뀐 버전)
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-NetTCPConnection -LocalPort 8080 -State Listen -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique | ForEach-Object { $p = Get-CimInstance Win32_Process -Filter ('ProcessId=' + $_); if ($p) { $pp = Get-Process -Id $p.ParentProcessId -ErrorAction SilentlyContinue; Stop-Process -Id $_ -Force; if ($pp -and $pp.ProcessName -eq 'py') { Stop-Process -Id $pp.Id -Force -ErrorAction SilentlyContinue } } }"
timeout /t 2 /nobreak >nul
wscript "%~dp0run_hidden.vbs"
timeout /t 8 /nobreak >nul
rem 4) 진단 기록 (Claude 확인용)
if not exist logs mkdir logs
(
  echo ===== %date% %time% =====
  powercfg /a
  powercfg /requests
  powercfg /query SCHEME_CURRENT SUB_BUTTONS LIDACTION
  netstat -ano | findstr ":8080"
) > logs\전원진단.txt 2>&1
netstat -ano | findstr /r /c:":8080 .*LISTENING" >nul && (
  echo 완료. 서버가 켜져 있습니다. 덮개를 닫고 휴대폰으로 접속되는지 확인해 주세요.
) || (
  echo [확인 필요] 서버가 아직 안 켜졌습니다. 잠시 뒤 서버켜기^(창없이^).bat 을 실행해 주세요.
)
pause
