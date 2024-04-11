FROM python:3.8-slim-buster

# Establece el directorio de trabajo en el contenedor
WORKDIR /app

# Copia el archivo de requisitos primero, para la caché de la capa
COPY requirements.txt /app/requirements.txt

# Instala las dependencias
RUN pip install --no-cache-dir -r requirements.txt

# Copia el resto del código de la aplicación
COPY . /app

# Expone el puerto en el que la aplicación se ejecutará
EXPOSE 5000

# Ejecuta la aplicación
CMD ["python", "app.py"]
