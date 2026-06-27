@echo off
setlocal
cd /d "%~dp0"

echo Building a shareable Windows app folder...
echo.

if not exist ".venv\Scripts\python.exe" (
  where py >nul 2>nul
  if %errorlevel%==0 (
    py -3 -m venv .venv
  ) else (
    python -m venv .venv
  )
)

".venv\Scripts\python.exe" -m pip install -r requirements.txt
".venv\Scripts\python.exe" -m pip install pyinstaller

if exist "build" rmdir /s /q "build"
if exist "dist\InventoryCatalog" rmdir /s /q "dist\InventoryCatalog"

".venv\Scripts\python.exe" -m PyInstaller ^
  --name InventoryCatalog ^
  --onedir ^
  --clean ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --hidden-import openpyxl ^
  --hidden-import ddgs ^
  app.py

echo.
echo Done.
echo Share this folder with another Windows PC:
echo dist\InventoryCatalog
echo.
echo On the other PC, double-click InventoryCatalog.exe and open http://127.0.0.1:5000
pause
