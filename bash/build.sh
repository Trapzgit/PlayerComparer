#!/bin/bash
set -e

# Название проекта
APP_NAME="PlayerComparer"

# Проверяем виртуальное окружение
if [ ! -d ".venv" ]; then
  echo "Создаю виртуальное окружение..."
  python -m venv .venv
fi

# Активируем окружение
source .venv/bin/activate

# Обновляем pip
pip install --upgrade pip

# Устанавливаем зависимости
if [ -f "requirements.txt" ]; then
  pip install -r requirements.txt
fi

# Ставим PyInstaller
pip install pyinstaller

# Собираем exe
pyinstaller --onefile --windowed --name "$APP_NAME" main.py

echo "✅ Сборка завершена!"
echo "Файл находится в dist/$APP_NAME.exe"
