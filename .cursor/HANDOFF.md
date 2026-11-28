# Handoff: Imagen Row Form

Continue this work in this repo. The previous chat was in **excel-allocation**. Allocation already adds a clickable **Link** column; this app is the page those URLs open.

Do not recap this file to the user unless they ask. Pick up the feature.

## What is already done

### excel-allocation (sibling repo: `../excel-allocation`)

Imagen output (main file, per-agent emailed slices, agent downloads) appends a **Link** column at the end. It is a real Excel hyperlink.

Base URL: `COMPARISON_TOOL_BASE_URL` or `http://localhost:5002/`.

Example:

```
http://localhost:5002/?patient_name=Jane+Doe&patient_dob=01%2F15%2F2010&subscriber_name=John+Doe&subscriber_dob=03%2F02%2F1980&pms_subscriber_id=12345&pms_payer=Aetna
```

Stable query keys (do not rename unless both repos change):

| Excel column | Query key |
|---|---|
| Patient Name | `patient_name` |
| Patient DOB | `patient_dob` |
| Subscriber Name | `subscriber_name` |
| Subscriber DOB | `subscriber_dob` |
| (PMS) Subscriber ID | `pms_subscriber_id` |
| (PMS) Payer | `pms_payer` |

Allocation source: `app.py` (`IMAGEN_LINK_QUERY_FIELDS`, `add_imagen_link_column`, `apply_imagen_link_hyperlinks_openpyxl`). Tests: `tests/test_imagen_links.py`.

That Link work may still be uncommitted in allocation. Do not commit either repo unless asked.

### excel-comparison-tool (this repo) — implemented, uncommitted

Uncommitted / new files:

- `imagen_row_form.py` — field defs, URL parse, PDF text extract, fill-missing, output PDF, page HTML
- `excel_comparison.py` — routes + sidebar item **Imagen Row Form**
- `requirements.txt` — `pdfplumber==0.11.7`, `reportlab==4.2.5`
- `tests/test_imagen_row_form.py`
- `run-dev.sh`

Routes:

| Method | Path | Behavior |
|---|---|---|
| GET | `/` | If any of the six query keys are present → form page. Else redirect `/comparison`. |
| GET | `/imagen-row` | Same form (also accepts the query string). |
| POST | `/imagen-row/extract-pdf` | PDF + current field values. Returns JSON `{values, updated}`. **Only empty fields are filled.** |
| POST | `/imagen-row/process` | All six fields required. Downloads `imagen-row.pdf`. |

PDF extract looks for labeled lines, e.g. `Patient Name: Jane Doe`, `(PMS) Payer: Aetna`. Output PDF is a simple one-page dump of the six fields (reportlab). That output is a first cut — the user never fully specified what the downloaded PDF should contain.

Dev server: `./run-dev.sh` → http://localhost:5002  
Tests: `python -m unittest tests.test_imagen_row_form -v`

## User intent (original)

> When user clicked on this url it should open the page and filled the details from the page url params. that page will have upload a pdf input and process button it will pick the details from pdf and fill the missing details in the page form. once required details filled it will click on process button then output pdf file will get downloaded.

## Agreed product rules (from the allocation chat)

- Extra columns stay at the **end** (Link already follows this on the allocation side).
- Do **not** change Imagen matching rules unless asked.
- Dates: format on upload rather than reject (allocation convention; only relevant if this app starts ingesting those Excel files).
- Do not commit/push unless the user asks.

## Likely next work (not built)

The user has not specified these yet. Ask or infer from a sample PDF rather than inventing a full insurance form:

- Real uploaded-PDF layout (which extra fields exist beyond the six).
- What **Process** should actually produce (eligibility packet, letter, filled CMS form, etc.). Current output is just the six fields.
- Production `COMPARISON_TOOL_BASE_URL` on allocation (still localhost by default).

## How to run both apps

- Allocation: `../excel-allocation/run-dev.sh` (usually port 5003).
- This app: `./run-dev.sh` (port 5002). Reloader is on.

The in-editor Cursor browser often cannot hit localhost; use curl or the user’s real browser.
