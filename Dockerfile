# Imagen base recomendada para FastAPI en Render
FROM python:3.11-slim

# Instala dependencias del sistema necesarias para WeasyPrint / CairoSVG
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

# Configura el entorno
WORKDIR /app
COPY . /app

# Instala dependencias de Python (usa el nombre correcto del archivo)
RUN pip install --no-cache-dir -r requirements.txt

# Expone el puerto que Render usa
EXPOSE 10000

# Comando para iniciar la API
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "10000"]
