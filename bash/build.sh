$APP_NAME = "PlayerComparer"

# Проверяем виртуальное окружение
if (!(Test-Path ".venv")) {
    Write-Host "Создаю виртуальное окружение..."
    python -m venv .venv
}

# Активируем окружение
. .\.venv\Scripts\Activate.ps1

# Обновляем pip
pip install --upgrade pip

# Устанавливаем зависимости
if (Test-Path "requirements.txt") {
    pip install -r requirements.txt
}

# Ставим PyInstaller
pip install pyinstaller

# Собираем exe
pyinstaller --onefile --windowed --name $APP_NAME main.py

Write-Host "✅ Сборка завершена!"
Write-Host "Файл находится в dist\$APP_NAME.exe"