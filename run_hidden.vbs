' Start the warehouse server with no window (runs server_bg.bat hidden) + lid watcher (monitors off when lid closes)
Set sh = CreateObject("WScript.Shell")
d = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.Run "cmd /c """ & d & "\server_bg.bat""", 0, False
sh.Run "wscript """ & d & "\lid_watcher.vbs""", 0, False
