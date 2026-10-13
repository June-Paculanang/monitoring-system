# Use the same Python major/minor version currently used by the Render service.
FROM python:3.14-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    LIBREOFFICE_PATH=/usr/bin/soffice \
    PORT=8000

WORKDIR /app

# LibreOffice Calc is required to convert the uploaded Excel report template
# into a PDF on the server. The user's phone/laptop does not need LibreOffice.
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        fonts-liberation \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt /app/requirements.txt
# Install project requirements and explicitly include python-dotenv because
# settings.py imports dotenv.load_dotenv.
RUN pip install --no-cache-dir -r /app/requirements.txt "python-dotenv>=1.0.1"

COPY . /app
RUN chmod +x /app/start.sh

EXPOSE 8000
CMD ["/app/start.sh"]
