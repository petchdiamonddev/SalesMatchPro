@echo off
chcp 65001 >nul

pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [!] Installing pyinstaller...
    pip install pyinstaller
)

echo [*] Building SalesMatchPro.exe...
pyinstaller --onefile --name SalesMatchPro --console --clean ^
    --exclude-module pandas ^
    --exclude-module numpy ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module tkinter ^
    process.py

if exist "dist\SalesMatchPro.exe" (
    copy /y "dist\SalesMatchPro.exe" "SalesMatchPro.exe" >nul
    echo.
    echo [*] Build finished: SalesMatchPro.exe
) else (
    echo.
    echo [X] Build failed.
)

pause
