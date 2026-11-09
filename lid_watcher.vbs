' Start lid_watcher.ps1 with no window (errors go to logs\lid_error.log)
Set sh = CreateObject("WScript.Shell")
d = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.Run "cmd /c powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & d & "\lid_watcher.ps1"" 2> """ & d & "\logs\lid_error.log""", 0, False
