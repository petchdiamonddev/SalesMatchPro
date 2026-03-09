@echo off
chcp 65001 >nul
echo ========================================
echo  SalesMatchPro - Build EXE
echo ========================================
echo.

REM ตรวจสอบว่ามี PyInstaller หรือยัง
pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [!] ไม่พบ PyInstaller กำลังติดตั้ง...
    echo [!] PyInstaller not found. Installing...
    pip install pyinstaller
    echo.
)

echo [*] กำลังสร้างไฟล์ .exe ...
echo [*] Building .exe file ...
echo.

pyinstaller ^
    --onefile ^
    --name SalesMatchPro ^
    --console ^
    --clean ^
    process.py

echo.
if exist "dist\SalesMatchPro.exe" (
    echo ========================================
    echo  ✅ สร้าง .exe สำเร็จ!
    echo  ✅ Build successful!
    echo ========================================
    echo.
    echo  ไฟล์อยู่ที่: dist\SalesMatchPro.exe
    echo  File location: dist\SalesMatchPro.exe
    echo.
    echo  วิธีใช้งาน:
    echo  1. คัดลอก SalesMatchPro.exe ไปยัง folder ที่ต้องการ
    echo  2. สร้าง folder input_sold/ และ input_all_data/
    echo  3. ใส่ไฟล์ Excel แล้วดับเบิลคลิก SalesMatchPro.exe
    echo.
) else (
    echo ========================================
    echo  ❌ สร้าง .exe ไม่สำเร็จ!
    echo  ❌ Build failed!
    echo ========================================
)

echo กด Enter เพื่อปิด...
pause >nul
