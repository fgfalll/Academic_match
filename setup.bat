@echo off
chcp 65001 >nul

echo Перевірка віртуального середовища...
if not exist ".venv\Scripts\activate.bat" (
    echo Створення віртуального середовища .venv...
    python -m venv .venv
    if errorlevel 1 (
        echo Помилка при створенні віртуального середовища. Переконайтеся, що встановлено Python.
        pause
        exit /b 1
    )
)

echo Активація віртуального середовища...
call .venv\Scripts\activate.bat

echo.
echo Встановлення залежностей...
pip install -r requirements.txt

echo.
echo Збірка виконуваного файлу...
pyinstaller --noconfirm --onefile --windowed --name "AcademicMatch" academ_back.py

echo.
echo Збірка завершена! Файл знаходиться у папці dist.
pause