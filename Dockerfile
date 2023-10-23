# Используйте официальный образ Python в качестве базового образа
FROM python:3.9

# Установка Nginx
RUN apt-get update && apt-get install -y nginx

# Копирование файлов конфигурации Nginx в контейнер
COPY nginx.conf /etc/nginx/nginx.conf
COPY default.conf /etc/nginx/conf.d/default.conf

# Открытие порта 80 для входящих соединений
EXPOSE 80

# Копирование файлов проекта и установка зависимостей Python
COPY excelanalog/requirements.txt /temp/requirements.txt
COPY excelanalog /excelanalog
WORKDIR /excelanalog
RUN pip install -r requirements.txt

# Запуск Nginx и Python-приложения
CMD service nginx start && python manage.py runserver 0.0.0.0:8000




