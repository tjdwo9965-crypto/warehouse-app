@echo off
chcp 65001 > nul
rem 모니터(노트북 화면 + 연장 모니터)만 바로 끕니다. 서버는 계속 돌아갑니다. 마우스/키보드를 움직이면 다시 켜집니다.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Sleep -Milliseconds 800; (Add-Type -MemberDefinition '[DllImport(\"user32.dll\")] public static extern int PostMessage(int hWnd, int msg, int wParam, int lParam);' -Name M -Namespace W -PassThru)::PostMessage(0xFFFF, 0x0112, 0xF170, 2) | Out-Null"
