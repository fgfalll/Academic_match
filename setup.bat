@echo off
chcp 65001 >nul

echo =============================================
echo   Academic Match Build Script
echo =============================================
echo.

echo [1/4] Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.8+
    pause
    exit /b 1
)
echo   Python found

echo.
echo [2/4] Checking virtual environment...
if not exist ".venv\Scripts\activate.bat" (
    echo   Creating .venv...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
) else (
    echo   .venv already exists
)

echo.
echo   Activating .venv...
call .venv\Scripts\activate.bat >nul 2>&1

echo.
echo [3/4] Installing dependencies...
pip install --upgrade pip >nul 2>&1
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)

echo.
echo [4/4] Building executable...
if exist "dist" (
    echo   Cleaning previous build...
    rmdir /s /q dist 2>nul
)
if exist "build" (
    rmdir /s /q build 2>nul
)
if exist "AcademicMatch.spec" (
    del /q AcademicMatch.spec 2>nul
)

pyinstaller --noconfirm --onefile --windowed --name "AcademicMatch" ^
    --add-data "ai_advisor.py;." ^
    --add-data "crypto_utils.py;." ^
    --hidden-import=crypto_utils ^
    --hidden-import=litellm ^
    --hidden-import=markdown ^
    --hidden-import=tkinterweb ^
    --hidden-import=ddgs ^
    academ_back.py

if errorlevel 1 (
    echo.
    echo ERROR: Build failed
    pause
    exit /b 1
)

echo.
echo =============================================
echo   BUILD COMPLETE!
echo =============================================
echo.
echo   Output: dist\AcademicMatch.exe
echo.
echo   Double-click to run or drag .acmp session file onto exe
echo.
pause
