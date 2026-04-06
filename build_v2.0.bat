@echo off
echo ========================================
echo Kruger Lab System - Build Script v2.0
echo ========================================
echo.

echo [1/4] Checking Python environment...
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python not found!
    pause
    exit /b 1
)
echo.

echo [2/4] Installing/updating dependencies...
pip install tkcalendar --quiet
if %errorlevel% neq 0 (
    echo ERROR: Failed to install tkcalendar!
    pause
    exit /b 1
)
echo Dependencies installed successfully.
echo.

echo [3/4] Building executable with PyInstaller...
if not exist "dist" mkdir dist
pyinstaller --clean Kruger_Lab_System.spec
if %errorlevel% neq 0 (
    echo ERROR: PyInstaller build failed!
    pause
    exit /b 1
)
echo Executable built successfully.
echo.

echo [4/4] Preparing setup files...
if not exist "Output" mkdir Output
echo Setup files ready.
echo.

echo ========================================
echo BUILD COMPLETED SUCCESSFULLY!
echo ========================================
echo.
echo Next steps:
echo 1. Open KrugerSetupScript.iss with Inno Setup Compiler
echo 2. Click Build -> Compile
echo 3. Find the installer in Output\ folder
echo.
echo Press any key to exit...
pause >nul