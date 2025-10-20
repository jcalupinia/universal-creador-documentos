# Imagen base ligera y estable para FastAPI
FROM python:3.11-slim

# Evita prompts durante instalación
ENV DEBIAN_FRONTEND=noninteractive

# Instala dependencias del sistema necesarias para PDF, SVG, fuentes, etc.
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libffi-dev \
    libcairo2 \
    libcairo2-dev \
    libpango1.0-0 \
    libpangocairo-1.0-0 \
    libjpeg-dev \
    libgdk-pixbuf-2.0-0 \
    libxml2 \
    libxslt1.1 \
    fonts-dejavu-core \
    shared-mime-info \
    wget \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Define el directorio de trabajo
WORKDIR /app

# Copia el código fuente
COPY . /app

# Instala dependencias Python
RUN pip install --no-cache-dir -r requirements.txt

# Expone puerto (Render define su propio PORT)
EXPOSE 10000

# Ejecuta la app FastAPI con puerto dinámico
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-10000}"]
