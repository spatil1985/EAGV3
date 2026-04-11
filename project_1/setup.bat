@echo off
echo Installing WinWhisper dependencies...
py -m pip install openai keyboard pyperclip
if errorlevel 1 (
    echo.
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)
echo.
echo Setup complete! Run run.bat to start WinWhisper.
pause
