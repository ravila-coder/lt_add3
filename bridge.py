"""Runs the unchanged Flask app inside the browser (Pyodide).

The page talks to worker.js, worker.js calls `call()` here, and `call()` drives app.py
through Flask's test client. No HTTP server exists: the routes, the global state and
build_workbook() are exactly the ones `python app.py` serves on a desktop, they just live
in this browser tab. A loan tape therefore never leaves the user's machine.
"""
import io
import os

import app as A

A.app.config["PROPAGATE_EXCEPTIONS"] = True   # a crash reaches the page as text, not a bare 500
client = A.app.test_client()


def call(method, path, body=None, fname=None, data=None):
    """One request. Returns (status, content_type, payload bytes)."""
    if data is not None:
        raw = data.to_bytes() if hasattr(data, "to_bytes") else bytes(data)
        r = client.post(path, data={"file": (io.BytesIO(raw), fname or "upload.csv")},
                        content_type="multipart/form-data")
    elif method == "GET":
        r = client.get(path)
    else:
        r = client.post(path, data=body or "{}", content_type="application/json")
    out = r.get_data()
    ctype = r.headers.get("Content-Type", "")
    r.close()
    # the browser keeps the downloaded workbook as a blob; free the in-memory file
    if method == "GET" and path.startswith("/download/") and r.status_code == 200:
        p = A.FILES.pop(path.rsplit("/", 1)[1], None)
        if p and os.path.exists(p):
            os.remove(p)
    return r.status_code, ctype, out
