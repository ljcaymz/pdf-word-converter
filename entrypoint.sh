#!/bin/sh
exec gunicorn app:app --bind "0.0.0.0:${PORT:-7860}" --workers 2 --timeout 120
