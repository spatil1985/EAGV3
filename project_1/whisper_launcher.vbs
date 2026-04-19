' WinWhisper silent launcher
' Runs whisper_app.py without showing a console window
Set objShell = CreateObject("WScript.Shell")
objShell.Run """C:\Users\sudip\AppData\Local\Programs\Python\Python314\python.exe"" ""C:\work\code\MyWinWhisper\whisper_app.py""", 0, False
