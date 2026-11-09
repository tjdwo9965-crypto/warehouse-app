@echo off
chcp 65001 > nul
cd /d "%~dp0"
set "APPDIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$s = (New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Startup') + '\창고재고서버.lnk'); $s.TargetPath = 'wscript.exe'; $s.Arguments = '\"' + $env:APPDIR + 'run_hidden.vbs\"'; $s.WorkingDirectory = $env:APPDIR; $s.Description = '창고 재고 서버 (창 없이 자동 시작)'; $s.Save()"
if exist "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\창고재고서버.lnk" (
  echo 자동 시작을 켰습니다. 이제 PC를 켜고 로그인하면 서버가 창 없이 자동으로 켜집니다.
  echo 끄려면 자동시작_끄기.bat 을 실행하세요.
) else (
  echo [오류] 자동 시작 등록에 실패했습니다. 이 화면을 캡처해서 보내 주세요.
)
pause
