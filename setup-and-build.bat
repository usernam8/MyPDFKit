@echo off
REM Run from the project folder in a normal Command Prompt.
REM Prereqs: Node.js LTS and Python 3.11 or 3.12 (tick "Add to PATH" in the installer).

echo === Installing Python packages ===
python -m pip install -r requirements.txt || goto :fail

echo === Installing Node packages ===
call npm install || goto :fail

echo === Building installer (server.exe + Electron) ===
call npm run build || goto :fail

echo.
echo Done. Installer is in the "release" folder.
goto :eof

:fail
echo Build failed - see the error above.
exit /b 1
