@echo off
chcp 65001 > nul
del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\창고재고서버.lnk" >nul 2>&1
echo 자동 시작을 껐습니다. (지금 켜져 있는 서버는 그대로입니다. 끄려면 서버끄기.bat)
pause
