# -*- coding: utf-8 -*-
"""Vercel serverless entry point.

Vercel's @vercel/python runtime detects the module-level ``app`` (an ASGI
application) and serves it. ``vercel.json`` routes every path here so the single
FastAPI app handles the UI, static assets, and the API.

Requires ``DATABASE_URL`` to be set in the Vercel project environment: the
serverless filesystem is read-only, so the local SQLite fallback cannot be used
there — the app must talk to Supabase/Postgres.
"""
import sys
from pathlib import Path

# make the project root importable regardless of the function's working dir
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.main import app  # noqa: E402,F401  (re-exported for the Vercel runtime)
