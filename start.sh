#!/usr/bin/env bash
set -e

# Run database migrations when the deployed container starts, after Render
# provides the production environment variables such as DATABASE_URL.
python manage.py migrate --noinput

# Ensure static files are collected with the runtime environment available.
python manage.py collectstatic --noinput

# Render provides PORT. Fall back to 8000 for local Docker runs.
exec gunicorn FacultyMonitoringSystem.wsgi:application --bind "0.0.0.0:${PORT:-8000}"
