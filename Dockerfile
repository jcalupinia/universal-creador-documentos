# Imagen base ligera y estable para FastAPI
FROM python:3.11-slim

# Evita prompts interactivos durante la instalación
ENV DEBIAN_FRONTEND=noninteractive

# Instala librerías del sistema necesarias para PDF, SVG y fuentes
RUN apt-get update && apt-get install -y \
    build-essential \
    libffi-dev \
    libpango-1.0-0 \
    libcairo2 \
    libcairo2-dev \
    libpangoft2-1.0-0 \
    libjpeg62-turbo-dev \
    shared-mime-info \
    fonts-dejavu-core \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Define el directorio de trabajo
WORKDIR /app

# Copia el código al contenedor
COPY . /app

# Instala dependencias de Python
RUN pip install --no-cache-dir -r requirements.txt

# Expone el puerto (Render usará su propio valor, pero este sirve como fallback)
EXPOSE 10000

# Variable de entorno para endpoint de salud
ENV HEALTHCHECK_PATH=/healthz

# Inicia la aplicación FastAPI usando el puerto dinámico de Render
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-10000}"]
