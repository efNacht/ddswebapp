#!/bin/bash
set -e

# Install dependencies
pip install -r requirements.txt

# Run Flask app with gunicorn
gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120
