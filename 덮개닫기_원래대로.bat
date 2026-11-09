@echo off
chcp 65001 > nul
rem LIDACTION 1 = 절전 (Windows 기본값)
powercfg /setacvalueindex SCHEME_CURRENT SUB_BUTTONS LIDACTION 1
powercfg /setactive SCHEME_CURRENT
echo 덮개를 닫으면 다시 절전으로 들어가도록 되돌렸습니다.
pause
