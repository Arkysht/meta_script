FROM python:3.10-slim

WORKDIR /app

# Копируем всё из репозитория
COPY . /app

# Устанавливаем зависимости
RUN if [ -f requirements.txt ]; then pip install --no-cache-dir -r requirements.txt; fi

# Команда для запуска скрипта
CMD ["python", "meta_ads_to_sheets.py"]