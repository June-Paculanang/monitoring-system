#!/usr/bin/env bash
# Exit on error
set -o errexit

# Upgrade pip
pip install --upgrade pip

# Install requirements
pip install -r requirements.txt

# Convert static files
python manage.py collectstatic --no-input

# Run database migrations
python manage.py migrate