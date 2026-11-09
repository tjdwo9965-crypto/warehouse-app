@echo off
chcp 65001 > nul
echo 노트북 덮개를 닫아도 절전으로 들어가지 않게 설정합니다. (전원 어댑터 연결 시에만)
rem LIDACTION 0 = 아무 것도 안 함 / STANDBYIDLE 0 = 절전 안 함 (AC 전원 기준)
powercfg /setacvalueindex SCHEME_CURRENT SUB_BUTTONS LIDACTION 0
powercfg /setacvalueindex SCHEME_CURRENT SUB_SLEEP STANDBYIDLE 0
powercfg /setacvalueindex SCHEME_CURRENT SUB_SLEEP HIBERNATEIDLE 0
powercfg /setactive SCHEME_CURRENT
echo.
powercfg /query SCHEME_CURRENT SUB_BUTTONS LIDACTION | findstr /c:"0x00000000" >nul && (
  echo 완료: 전원 연결 상태에서는 덮개를 닫아도 서버가 계속 돌아갑니다.
) || (
  echo [확인 필요] 설정이 적용되지 않았을 수 있습니다. 이 화면을 캡처해서 보내 주세요.
)
echo 배터리만 쓸 때는 기존처럼 덮개를 닫으면 절전됩니다. (가방 속 과열 방지)
echo 원래대로 되돌리려면 덮개닫기_원래대로.bat 을 실행하세요.
pause
