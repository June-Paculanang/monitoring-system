
FROM python:3.14-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    LIBREOFFICE_PATH=/usr/bin/soffice

WORKDIR /app

# Install LibreOffice for Excel-template-to-PDF conversion.
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        fonts-liberation \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies.
COPY requirements.txt /app/requirements.txt

RUN pip install --no-cache-dir -r /app/requirements.txt \
    && pip install --no-cache-dir "python-dotenv>=1.0.1"

# Copy the Django project.
COPY . /app

# Make the startup script executable.
RUN chmod +x /app/start.sh

# Render supplies the actual PORT environment variable.
EXPOSE 10000

CMD ["/app/start.sh"]