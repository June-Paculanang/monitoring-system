
#!/usr/bin/env bash

set -e

echo "Starting Django deployment..."

# Apply database migrations.
python manage.py migrate --noinput

# Collect static files.
python manage.py collectstatic --noinput

# Start Gunicorn on the port provided by Render.
exec gunicorn TeacherMonitoringSystem.wsgi:application \
    --bind "0.0.0.0:${PORT:-10000}" \
    --workers 1 \
    --timeout 120