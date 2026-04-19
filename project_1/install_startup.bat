@echo off
:: Installs WinWhisper to run automatically at Windows startup
:: Places a shortcut in the user Startup folder — no admin needed

set STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set SHORTCUT=%STARTUP%\WinWhisper.lnk
set LAUNCHER=C:\work\code\MyWinWhisper\whisper_launcher.vbs

echo Installing WinWhisper to startup...

:: Create shortcut using PowerShell
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('%SHORTCUT%'); $sc.TargetPath = 'wscript.exe'; $sc.Arguments = '\"%LAUNCHER%\"'; $sc.Description = 'WinWhisper voice-to-text'; $sc.Save()"

if exist "%SHORTCUT%" (
    echo.
    echo [OK] WinWhisper will start automatically when you log in.
    echo      Shortcut placed at:
    echo      %SHORTCUT%
    echo.
    echo      To remove: run uninstall_startup.bat
) else (
    echo [ERR] Failed to create shortcut.
)
pause
