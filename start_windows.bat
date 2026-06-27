@echo off
setlocal
cd /d "%~dp0"

echo Starting Local Inventory Catalog...
echo.

if exist ".venv\Scripts\python.exe" (
  set "APP_PYTHON=.venv\Scripts\python.exe"
) else (
  where py >nul 2>nul
  if %errorlevel%==0 (
    py -3 -m venv .venv
  ) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
      python -m venv .venv
    ) else (
      echo Python is not installed or not added to PATH.
      echo Install Python from https://www.python.org/downloads/
      echo During install, tick "Add python.exe to PATH".
      pause
      exit /b 1
    )
  )
  set "APP_PYTHON=.venv\Scripts\python.exe"
)

"%APP_PYTHON%" -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Package installation failed. Check your internet connection and try again.
  pause
  exit /b 1
)

echo.
echo Open this link in your browser:
echo http://127.0.0.1:5000
echo.
set OPEN_BROWSER=1
"%APP_PYTHON%" app.py
pause
