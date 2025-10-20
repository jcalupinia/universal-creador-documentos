# Imagen base recomendada para FastAPI en Render
FROM python:3.11-slim

# Instala dependencias del sistema (para WeasyPrint y CairoSVG)
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

# Instala dependencias de Python
RUN pip install --no-cache-dir -r requisitos.txt

# Expone el puerto usado por Render
EXPOSE 10000

# Comando para iniciar la API (ajusta “main:app” si tu archivo se llama distinto)
CMD ["uvicorn", "principal:app", "--host", "0.0.0.0", "--port", "10000"]
