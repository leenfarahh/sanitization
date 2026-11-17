@echo off
rem One-time setup per device: puts a "Sanitize" shortcut on the desktop.
rem Needs Python 3.9 or later from python.org. There are no packages to install.
cd /d "%~dp0"

set "PYW="
for /f "delims=" %%p in ('where pythonw 2^>nul') do if not defined PYW set "PYW=%%p"
if not defined PYW (
  echo Python is not installed. Install it from python.org, tick "Add Python to PATH", then run this again.
  pause
  exit /b 1
)
set "PY=%PYW:pythonw.exe=python.exe%"

"%PY%" -c "import sys, tkinter; sys.exit(sys.version_info < (3, 9))" 2>nul || (
  echo Sanitize needs Python 3.9 or later from python.org, with "tcl/tk and IDLE" ticked in the installer.
  echo Install or repair Python, then run this again.
  pause
  exit /b 1
)

echo Creating desktop shortcut...
"%PY%" sanitize_desktop.py --setup || (
  pause
  exit /b 1
)

echo.
echo Done. Double-click "Sanitize" on your desktop to start.
pause
