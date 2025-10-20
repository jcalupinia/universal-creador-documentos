# Imagen base ligera y estable para FastAPI
FROM python:3.11-slim

# Evita prompts interactivos durante el build
ENV DEBIAN_FRONTEND=noninteractive

# Instala librerías del sistema necesarias para PDF, SVG y fuentes
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libffi-dev \
    libcairo2 \
    libcairo2-dev \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libjpeg-dev \
    libgdk-pixbuf2.0-0 \
    shared-mime-info \
    fonts-dejavu-core \
    libxml2 \
    libxslt1.1 \
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

# Inicia la aplicación FastAPI usando el puerto dinámico de Render
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-10000}"]
