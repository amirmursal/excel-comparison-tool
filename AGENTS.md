# Excel Comparison Tool

Flask app in `excel_comparison.py` (single large file, many sidebar tools). Default port **5002**. Start with `./run-dev.sh` (reloader is on; saving `.py` files restarts the app).

Python 3.9 venv: `.venv`. Install with `pip install -r requirements.txt`.

## Before changing code

If the task is **Imagen Row Form** (allocation Link URLs, PDF upload, Process → PDF), read [`.cursor/HANDOFF.md`](.cursor/HANDOFF.md) first. That file is the handoff from the excel-allocation chat where this feature started.

## Conventions

- Keep new Imagen Row logic in `imagen_row_form.py` plus thin routes in `excel_comparison.py`. Do not dump helpers into the 20k-line file.
- `/` with no query params still redirects to `/comparison`. Do not break that.
- Query param names are a contract with **excel-allocation** (`IMAGEN_LINK_QUERY_FIELDS`). Do not rename them without changing allocation too.
- Do not commit or push unless the user asks.
- Tests: `python -m unittest tests.test_imagen_row_form -v` (from repo root, venv on).
