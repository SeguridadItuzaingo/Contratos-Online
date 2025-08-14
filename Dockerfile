FROM python:3.11-slim

# LibreOffice para convertir DOCX->PDF en Linux
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer libreoffice-core fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Instalar dependencias primero (mejor cache)
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copiar el resto del proyecto
COPY . /app

ENV PYTHONUNBUFFERED=1 \
    PORT=10000 \
    FLASK_ENV=production

EXPOSE 10000

# Ejecutar la app con Gunicorn
CMD ["gunicorn", "-b", "0.0.0.0:10000", "app:app"]
