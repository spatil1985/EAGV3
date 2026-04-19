@echo off
:: Removes WinWhisper from Windows startup

set SHORTCUT=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\WinWhisper.lnk

if exist "%SHORTCUT%" (
    del "%SHORTCUT%"
    echo [OK] WinWhisper removed from startup.
) else (
    echo [INFO] WinWhisper was not in startup (already removed).
)
pause
