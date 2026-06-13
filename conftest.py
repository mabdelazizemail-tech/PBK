# -*- coding: utf-8 -*-
"""Pytest session bootstrap.

Guarantees the test suite always runs against an isolated SQLite database and
never a Postgres/Supabase instance configured via ``DATABASE_URL`` or a local
``.env``. This file is imported by pytest before any test module (and therefore
before ``app.db``), so setting ``PBK_DB`` here pins the SQLite backend: ``db.py``
only selects Postgres when ``DATABASE_URL`` is set AND ``PBK_DB`` is unset.

``setdefault`` keeps any explicit ``PBK_DB`` a test or CI already provided.
"""
import os
import tempfile

os.environ.setdefault("PBK_DB", os.path.join(tempfile.mkdtemp(), "pbk_test.db"))
