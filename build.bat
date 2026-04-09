@echo off
echo Stopping running AcademicMatch.exe...
taskkill /F /IM AcademicMatch.exe 2>nul
timeout /t 2 /nobreak >nul
echo Building AcademicMatch.exe...
.\.venv\Scripts\pyinstaller.exe AcademicMatch.spec --clean
echo.
echo Done! Check dist\AcademicMatch.exe
pause
