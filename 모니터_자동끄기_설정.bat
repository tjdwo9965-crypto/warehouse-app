@echo off
chcp 65001 > nul
echo 아무 조작이 없으면 5분 뒤 모니터(연장 모니터 포함)가 꺼지도록 설정합니다. (전원 연결 시)
echo 화면만 꺼지고 서버는 계속 돌아갑니다.
powercfg /setacvalueindex SCHEME_CURRENT SUB_VIDEO VIDEOIDLE 300
powercfg /setactive SCHEME_CURRENT
echo.
echo 완료. 바로 끄고 싶을 때는 화면끄기.bat 을 실행하세요.
pause
