@echo off
REM =====================================
REM Сборка PlayerComparer.exe
REM =====================================

REM Активируем виртуальное окружение
call .venv\Scripts\activate

REM Удаляем старые сборки
echo Очистка старых сборок...
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul
del /q PlayerComparer.spec 2>nul

REM Проверяем наличие pyinstaller
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo Устанавливаю pyinstaller...
    pip install pyinstaller
)

REM Собираем exe
echo Собираю exe...
pyinstaller --onefile --windowed --name PlayerComparer main.py

echo.
echo =====================================
echo Сборка завершена!
echo Файл: dist\PlayerComparer.exe
echo =====================================
pause
