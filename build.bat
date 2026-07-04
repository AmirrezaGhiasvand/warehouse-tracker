@echo off
REM Builds Warehouse Tracker into a single standalone .exe
REM Run this on Windows, inside the project folder, after installing Python.

echo Installing dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo FAILED to install dependencies. See the error above.
    pause
    exit /b 1
)

echo Building exe...
python -m PyInstaller --onefile --windowed --name WarehouseTracker app.py
if %errorlevel% neq 0 (
    echo.
    echo FAILED to build the exe. See the error above.
    pause
    exit /b 1
)

echo.
echo Done! Your exe is at: dist\WarehouseTracker.exe
echo You can copy that single file anywhere and run it - no Python needed.
pause
