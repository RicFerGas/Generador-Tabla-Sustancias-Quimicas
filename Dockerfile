# Usar una imagen base más ligera de Debian Slim
FROM debian:bullseye-slim

# Configurar las variables de entorno
ENV WINEDEBUG=-all
ENV DEBIAN_FRONTEND=noninteractive

# Instalar las dependencias necesarias sin wine32, solo wine64
RUN dpkg --add-architecture i386 && apt-get update && apt-get install -y \
    wine64 \
    xvfb \
    python3 \
    python3-pip \
    python3-setuptools \
    wget \
    unzip \
    curl

# Descargar e instalar Python para Windows
RUN wget https://www.python.org/ftp/python/3.9.6/python-3.9.6-amd64.exe -O /tmp/python.exe && \
    wine /tmp/python.exe /quiet InstallAllUsers=1 PrependPath=1

# Instalar PyInstaller en el entorno Wine (Python para Windows)
RUN wine python -m pip install pyinstaller

# Crear el directorio de la aplicación
WORKDIR /app

# Copiar los archivos de tu proyecto al contenedor
COPY . /app

# Ejecutar PyInstaller para empaquetar la aplicación en un ejecutable de Windows
ENTRYPOINT ["wine", "python", "-m", "pyinstaller", "--onefile", "main.py"]
