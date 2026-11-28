"""Imagen allocation row form: URL params, PDF fill-in, and output PDF."""

import io
import os
import re
import tempfile
from datetime import datetime

from urllib.parse import unquote_plus


FORM_FIELDS = (
    ("patient_name", "Patient Name"),
    ("patient_dob", "Patient DOB"),
    ("subscriber_name", "Subscriber Name"),
    ("subscriber_dob", "Subscriber DOB"),
    ("pms_subscriber_id", "(PMS) Subscriber ID"),
    ("pms_payer", "(PMS) Payer"),
)

# Optional fields (form + PDF only; not allocation URL contract).
PERSONAL_DETAIL_FORM_FIELDS = (
    ("insurance_mailing_address", "Insurance Mailing Address"),
    ("payer_id", "Payer ID #"),
    ("group_name", "Group Name"),
    ("provider_network_status", "Provider Network Status"),
    ("coordination_of_benefits", "Coordination of Benefits"),
)
BENEFIT_DETAIL_FORM_FIELDS = (
    ("effective_date_of_plan", "Effective Date of Plan"),
    ("annual_max_remaining", "Annual Max Remaining"),
    ("individual_deductible_met", "Individual Deductible Met"),
    ("family_deductible_met", "Family Deductible Met"),
    (
        "preventative_subject_to_annual_max",
        "Preventative/Diagnostic Subject to Annual Max? Yes/No",
    ),
)
COVERAGE_DETAIL_FORM_FIELDS = (
    ("missing_tooth_clause", "Missing Tooth Clause? Y/N"),
)
FREQUENCY_DETAIL_FORM_FIELDS = (
    ("fmx_pano_shared_freq", "FMX/Pano Shared freq? Yes/No"),
    ("perm_molars_only", "Perm Molars only?"),
    (
        "downgrades_posterior_composites",
        "Downgrades on Posterior Composites? Yes/No",
    ),
    ("crowns_downgrade", "Crowns Downgrade:"),
    ("paid_at_prep_or_seat", "Paid at Prep or Seat Date?"),
    (
        "frequency_d4910_d1110",
        "Frequency of D4910 combined with D1110?",
    ),
    ("quads_allowed_one_day", "How many quads allowed in one day?"),
)
ORTHO_DETAIL_FORM_FIELDS = (
    ("ortho_ltm_remaining", "Ortho LTM Remaining"),
)
DETAIL_FORM_FIELDS = (
    ("exam", "Exam"),
    ("prophy", "Prophy"),
    ("bitewings", "Bitewings"),
    ("fluoride", "Fluoride"),
    ("pano_fmx", "Pano/ FMX"),
    ("verified_by", "Verified By"),
    ("date_verified", "Date Verified"),
)

OPTIONAL_FIELD_SECTIONS = (
    ("Personal Information", PERSONAL_DETAIL_FORM_FIELDS),
    ("Benefit Breakdown", BENEFIT_DETAIL_FORM_FIELDS),
    ("Coverage", COVERAGE_DETAIL_FORM_FIELDS),
    ("Frequencies", FREQUENCY_DETAIL_FORM_FIELDS),
    ("Orthodontics", ORTHO_DETAIL_FORM_FIELDS),
    ("History", DETAIL_FORM_FIELDS),
)

OPTIONAL_FORM_FIELDS = tuple(
    field
    for _title, fields in OPTIONAL_FIELD_SECTIONS
    for field in fields
)

FORM_FIELD_KEYS = tuple(key for key, _label in FORM_FIELDS)
OPTIONAL_FORM_FIELD_KEYS = tuple(key for key, _label in OPTIONAL_FORM_FIELDS)
ALL_FORM_FIELD_KEYS = FORM_FIELD_KEYS + OPTIONAL_FORM_FIELD_KEYS

_FIELD_PATTERNS = {
    "patient_name": [
        r"patient\s*name\s*[:\-]\s*(.+)",
    ],
    "patient_dob": [
        r"patient\s*(?:d\.?o\.?b\.?|date of birth)\s*[:\-]\s*([0-9]{1,2}[/\-][0-9]{1,2}[/\-][0-9]{2,4})",
    ],
    "subscriber_name": [
        r"subscriber\s*name\s*[:\-]\s*(.+)",
    ],
    "subscriber_dob": [
        r"subscriber\s*(?:d\.?o\.?b\.?|date of birth)\s*[:\-]\s*([0-9]{1,2}[/\-][0-9]{1,2}[/\-][0-9]{2,4})",
    ],
    "pms_subscriber_id": [
        r"\(pms\)\s*subscriber\s*id\s*[:\-]\s*([A-Za-z0-9\-]+)",
        r"(?:pms\s*)?subscriber\s*id\s*[:\-]\s*([A-Za-z0-9\-]+)",
        r"member\s*id\s*[:\-]\s*([A-Za-z0-9\-]+)",
    ],
    "pms_payer": [
        r"\(pms\)\s*payer\s*[:\-]\s*(.+)",
        r"(?:pms\s*)?payer\s*[:\-]\s*(.+)",
        r"insurance(?:\s*carrier|\s*name)?\s*[:\-]\s*(.+)",
    ],
}


def has_form_query(args):
    return any(str(args.get(key) or "").strip() for key in FORM_FIELD_KEYS)


def values_from_query(args):
    values = {}
    for key in ALL_FORM_FIELD_KEYS:
        raw = args.get(key)
        if raw is None:
            values[key] = ""
            continue
        text = unquote_plus(str(raw)).strip()
        values[key] = text
    return values


def values_from_form(form):
    return {
        key: (form.get(key) or "").strip()
        for key in ALL_FORM_FIELD_KEYS
    }


def _clean_extracted(value):
    text = str(value or "").strip()
    text = re.sub(r"\s+", " ", text)
    return text.strip(" |,-")


def extract_fields_from_text(text):
    found = {}
    blob = str(text or "")
    for key, patterns in _FIELD_PATTERNS.items():
        for pattern in patterns:
            match = re.search(pattern, blob, flags=re.IGNORECASE)
            if not match:
                continue
            cleaned = _clean_extracted(match.group(1))
            if cleaned:
                found[key] = cleaned
                break
    return found


def extract_text_from_pdf_bytes(pdf_bytes):
    import pdfplumber

    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf_bytes)
            tmp_path = tmp.name
        pages = []
        with pdfplumber.open(tmp_path) as pdf:
            for page in pdf.pages:
                pages.append(page.extract_text() or "")
        return "\n".join(pages)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)


def fill_missing_fields(current_values, extracted_values):
    filled = {}
    updates = {}
    for key, _label in FORM_FIELDS:
        current = str((current_values or {}).get(key) or "").strip()
        extracted = str((extracted_values or {}).get(key) or "").strip()
        if current:
            filled[key] = current
        elif extracted:
            filled[key] = extracted
            updates[key] = extracted
        else:
            filled[key] = ""
    # Preserve optional detail fields; PDF extract does not fill them.
    for key, _label in OPTIONAL_FORM_FIELDS:
        filled[key] = str((current_values or {}).get(key) or "").strip()
    return filled, updates


def missing_required_fields(values):
    return [
        label
        for key, label in FORM_FIELDS
        if not str((values or {}).get(key) or "").strip()
    ]


def _join_csv_parts(*parts):
    cleaned = [str(part or "").strip() for part in parts if str(part or "").strip()]
    return ", ".join(cleaned)


def build_output_pdf_bytes(values):
    """Build bordered Imagen form PDF (label + value / Percentage + Limitations)."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    values = values or {}
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
    )

    header_bg = colors.HexColor("#60caf3")
    border = colors.HexColor("#333333")
    last_col = 2  # label | Percentage | Limitations
    cell_style = ParagraphStyle(
        "ImagenCell",
        fontName="Helvetica",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#222222"),
    )
    header_style = ParagraphStyle(
        "ImagenHeader",
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=16,
        alignment=1,  # center
        textColor=colors.white,
    )
    subheader_style = ParagraphStyle(
        "ImagenSubHeader",
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        alignment=1,
        textColor=colors.white,
    )

    def cell(text):
        return Paragraph(str(text or "").replace("&", "&amp;").replace("<", "&lt;"), cell_style)

    def section_header(title):
        # 3 cells so Table stays consistent; SPAN covers all columns.
        return [Paragraph(title, header_style), "", ""]

    def label_value_row(label, value=""):
        # Value spans Percentage + Limitations columns for earlier sections.
        return [cell(label), cell(value), ""]

    patient_value = _join_csv_parts(
        values.get("patient_name"), values.get("patient_dob")
    )
    subscriber_value = _join_csv_parts(
        values.get("subscriber_name"), values.get("subscriber_dob")
    )
    subscriber_id = str(values.get("pms_subscriber_id") or "").strip()
    payer = str(values.get("pms_payer") or "").strip()

    def v(key):
        return str(values.get(key) or "").strip()

    personal_rows = [
        section_header("Personal Information"),
        label_value_row("Patient name, DOB", patient_value),
        label_value_row("Subscriber Name, DOB", subscriber_value),
        label_value_row("Subscriber ID #", subscriber_id),
        label_value_row("Insurance Company/Phone #", payer),
        label_value_row("Insurance Mailing Address", v("insurance_mailing_address")),
        label_value_row("Payer ID #", v("payer_id")),
        label_value_row("Group Name", v("group_name")),
        label_value_row("Group Number"),
        label_value_row("Provider Network Status", v("provider_network_status")),
        label_value_row("Coordination of Benefits", v("coordination_of_benefits")),
    ]
    benefit_rows = [
        section_header("Benefit Breakdown"),
        label_value_row("Effective Date of Plan", v("effective_date_of_plan")),
        label_value_row("Calendar or Fiscal Year"),
        label_value_row("Annual Max Amount"),
        label_value_row("Annual Max Remaining", v("annual_max_remaining")),
        label_value_row("Individual Deductible"),
        label_value_row("Individual Deductible Met", v("individual_deductible_met")),
        label_value_row("Family Deductible"),
        label_value_row("Family Deductible Met", v("family_deductible_met")),
        label_value_row("Deductible Applies to: Prev/Basic/Major/Ortho"),
        label_value_row(
            "Preventative/Diagnostic Subject to Annual Max? Yes/No",
            v("preventative_subject_to_annual_max"),
        ),
    ]
    coverage_rows = [
        section_header("Coverage Percentage & Limitations"),
        label_value_row("Preventative"),
        label_value_row("Basic"),
        label_value_row("Major"),
        label_value_row("Endo"),
        label_value_row("Perio"),
        label_value_row("Oral Surgery"),
        label_value_row(
            "Missing Tooth Clause? Y/N Notes", v("missing_tooth_clause")
        ),
        label_value_row("Waiting Periods? Y/N – Which procedures?"),
    ]
    frequencies_rows = [
        section_header("Frequencies, Coverage & Limitations"),
        [
            "",
            Paragraph("Percentage", subheader_style),
            Paragraph("Limitations", subheader_style),
        ],
        [cell("Periodic Exams (D0120)"), cell(""), cell("")],
        [cell("Limited Exams (D0140):"), cell(""), cell("")],
        [cell("FMX/Pano Frequency (D0210/D0330)"), cell(""), cell("")],
        # Merged Percentage + Limitations into one value cell
        [cell("FMX/Pano Shared freq? Yes/No"), cell(v("fmx_pano_shared_freq")), ""],
        [cell("D0220: PAs"), cell(""), cell("")],
        [cell("Bitewings (D0274)"), cell(""), cell("")],
        [cell("Prophy (D1120 or D1110)"), cell(""), cell("")],
        [cell("Fluoride (D1206)"), cell(""), cell("")],
        # Merged Percentage + Limitations into one value cell
        [cell("Age Limit on Fluoride?"), cell(""), ""],
        [cell("Sealants (D1351)"), cell(""), cell("")],
        # Merged Percentage + Limitations into one value cell
        [cell("Age Limit on Sealants"), cell(""), ""],
        [cell("Perm Molars only?"), cell(v("perm_molars_only")), ""],
        # Empty row spanning label + Percentage + Limitations
        ["", "", ""],
        [cell("D2391 Posterior Composite"), cell(""), cell("")],
        # Merged Percentage + Limitations into one value cell
        [
            cell("Downgrades on Posterior Composites? Yes/No"),
            cell(v("downgrades_posterior_composites")),
            "",
        ],
        # Empty row spanning label + Percentage + Limitations
        ["", "", ""],
        [cell("D2740 Posterior Crown:"), cell(""), cell("")],
        # Merged Percentage + Limitations into one value cell
        [cell("Crowns Downgrade:"), cell(v("crowns_downgrade")), ""],
        [cell("D2950 Build Up"), cell(""), cell("")],
        # Merged Percentage + Limitations into one value cell
        [cell("Paid at Prep or Seat Date?"), cell(v("paid_at_prep_or_seat")), ""],
        [cell("Endo D3310"), cell(""), cell("")],
        # Empty row spanning label + Percentage + Limitations
        ["", "", ""],
        [cell("Perio Maintenance (D4910)"), cell(""), cell("")],
        # Merged Percentage + Limitations into one value cell
        [
            cell("Frequency of D4910 combined with D1110?"),
            cell(v("frequency_d4910_d1110")),
            "",
        ],
        [cell("Scaling & Root Planning (D4341/D4342)"), cell(""), cell("")],
        # Merged Percentage + Limitations into one value cell
        [
            cell("How many quads allowed in one day?"),
            cell(v("quads_allowed_one_day")),
            "",
        ],
        [cell("D4381 Arestin"), cell(""), cell("")],
        [cell("D4355 FMD"), cell(""), cell("")],
        # Empty row spanning label + Percentage + Limitations
        ["", "", ""],
        [cell("D5110 Denture"), cell(""), cell("")],
        # Empty row spanning label + Percentage + Limitations
        ["", "", ""],
        [cell("D6010 Implant"), cell(""), cell("")],
        [cell("D6057 Abutment"), cell(""), cell("")],
        [cell("D6058 Implant Crown"), cell(""), cell("")],
        [cell("D7140 Simple Extraction"), cell(""), cell("")],
        [cell("D7210 Surgical Extraction"), cell(""), cell("")],
        [cell("D7953 Bone Graft"), cell(""), cell("")],
        [cell("D9110 Palliative"), cell(""), cell("")],
        [cell("D9223 Sedation"), cell(""), cell("")],
        [cell("D9230 Nitrous"), cell(""), cell("")],
        [cell("D9944 Occlusal Guard/Limitation"), cell(""), cell("")],
    ]
    orthodontics_rows = [
        section_header("Orthodontics"),
        [cell("Orthodontics"), cell(""), cell("")],
        [cell("Age Limits"), cell(""), cell("")],
        [cell("Ortho Lifetime Max"), cell(""), cell("")],
        [cell("Ortho LTM Remaining"), cell(v("ortho_ltm_remaining")), ""],
        [cell("Ortho Deductible"), cell(""), cell("")],
        # Common History header directly below Ortho Deductible
        [Paragraph("History", header_style), "", ""],
        [cell("Exam"), cell(v("exam")), ""],
        [cell("Prophy"), cell(v("prophy")), ""],
        [cell("Bitewings"), cell(v("bitewings")), ""],
        [cell("Fluoride"), cell(v("fluoride")), ""],
        [cell("Pano/ FMX"), cell(v("pano_fmx")), ""],
        # Empty bar: club columns with header background
        ["", "", ""],
        [cell("Verified By"), cell(v("verified_by")), ""],
        [cell("Date Verified"), cell(v("date_verified")), ""],
    ]
    table_data = (
        personal_rows
        + benefit_rows
        + coverage_rows
        + frequencies_rows
        + orthodontics_rows
    )

    frequencies_start = len(personal_rows) + len(benefit_rows) + len(coverage_rows)
    orthodontics_start = frequencies_start + len(frequencies_rows)
    section_header_rows = [
        0,
        len(personal_rows),
        len(personal_rows) + len(benefit_rows),
        frequencies_start,
        orthodontics_start,
    ]
    frequencies_subheader_row = frequencies_start + 1
    # Ortho Deductible is offset 5; History header is the next row.
    orthodontics_history_header_row = orthodontics_start + 6
    # Within frequencies_rows: merged value cells (Percentage + Limitations)
    frequencies_merged_value_rows = [
        frequencies_start + 5,  # FMX/Pano Shared freq? Yes/No
        frequencies_start + 10,  # Age Limit on Fluoride?
        frequencies_start + 12,  # Age Limit on Sealants
        frequencies_start + 13,  # Perm Molars only?
        frequencies_start + 16,  # Downgrades on Posterior Composites?
        frequencies_start + 19,  # Crowns Downgrade:
        frequencies_start + 21,  # Paid at Prep or Seat Date?
        frequencies_start + 25,  # Frequency of D4910 combined with D1110?
        frequencies_start + 27,  # How many quads allowed in one day?
    ]
    frequencies_full_span_rows = [
        frequencies_start + 14,  # empty spacer after Perm Molars
        frequencies_start + 17,  # empty spacer after Downgrades
        frequencies_start + 23,  # empty spacer after Endo D3310
        frequencies_start + 30,  # empty spacer after D4355 FMD
        frequencies_start + 32,  # empty spacer after D5110 Denture
    ]
    # Ortho / History detail rows: value spans Percentage + Limitations.
    history_value_span_rows = [
        orthodontics_start + 4,  # Ortho LTM Remaining
        orthodontics_start + 7,  # Exam
        orthodontics_start + 8,  # Prophy
        orthodontics_start + 9,  # Bitewings
        orthodontics_start + 10,  # Fluoride
        orthodontics_start + 11,  # Pano/ FMX
        orthodontics_start + 13,  # Verified By
        orthodontics_start + 14,  # Date Verified
    ]
    # Earlier sections + selected Frequencies rows: value spans Percentage + Limitations.
    value_span_rows = [
        idx
        for idx in range(frequencies_start)
        if idx not in section_header_rows
    ] + frequencies_merged_value_rows + history_value_span_rows
    # Empty colored bar after Pano/FMX (club all columns, header background).
    orthodontics_empty_bar_row = orthodontics_start + len(orthodontics_rows) - 3

    col_widths = [3.0 * inch, 2.0 * inch, 2.0 * inch]
    table = Table(table_data, colWidths=col_widths, hAlign="LEFT")
    style_commands = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("BOX", (0, 0), (-1, -1), 1, border),
        ("INNERGRID", (0, 0), (-1, -1), 1, border),
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
    ]
    for row_idx in section_header_rows:
        style_commands.extend(
            [
                ("SPAN", (0, row_idx), (last_col, row_idx)),
                ("BACKGROUND", (0, row_idx), (last_col, row_idx), header_bg),
                ("TEXTCOLOR", (0, row_idx), (last_col, row_idx), colors.white),
                ("ALIGN", (0, row_idx), (last_col, row_idx), "CENTER"),
            ]
        )
    for row_idx in value_span_rows:
        style_commands.append(("SPAN", (1, row_idx), (last_col, row_idx)))
    for row_idx in frequencies_full_span_rows:
        style_commands.append(("SPAN", (0, row_idx), (last_col, row_idx)))
    style_commands.extend(
        [
            (
                "SPAN",
                (0, orthodontics_history_header_row),
                (last_col, orthodontics_history_header_row),
            ),
            (
                "BACKGROUND",
                (0, orthodontics_history_header_row),
                (last_col, orthodontics_history_header_row),
                header_bg,
            ),
            (
                "TEXTCOLOR",
                (0, orthodontics_history_header_row),
                (last_col, orthodontics_history_header_row),
                colors.white,
            ),
            (
                "ALIGN",
                (0, orthodontics_history_header_row),
                (last_col, orthodontics_history_header_row),
                "CENTER",
            ),
            ("SPAN", (0, orthodontics_empty_bar_row), (last_col, orthodontics_empty_bar_row)),
            (
                "BACKGROUND",
                (0, orthodontics_empty_bar_row),
                (last_col, orthodontics_empty_bar_row),
                header_bg,
            ),
            (
                "BACKGROUND",
                (0, frequencies_subheader_row),
                (last_col, frequencies_subheader_row),
                header_bg,
            ),
            (
                "TEXTCOLOR",
                (0, frequencies_subheader_row),
                (last_col, frequencies_subheader_row),
                colors.white,
            ),
            (
                "ALIGN",
                (0, frequencies_subheader_row),
                (last_col, frequencies_subheader_row),
                "CENTER",
            ),
        ]
    )
    table.setStyle(TableStyle(style_commands))

    meta = ParagraphStyle(
        "ImagenMeta",
        fontName="Helvetica",
        fontSize=9,
        textColor=colors.HexColor("#666666"),
    )
    story = [
        Paragraph(
            datetime.now().strftime("Generated %m/%d/%Y %I:%M %p"),
            meta,
        ),
        Spacer(1, 12),
        table,
    ]
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


PAGE_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Imagen Row Form</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: #f5f7fb;
            color: #333;
        }
        .sidebar {
            position: fixed;
            left: 0;
            top: 0;
            bottom: 0;
            width: 280px;
            background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 24px 0;
            overflow-y: auto;
        }
        .sidebar-header { padding: 0 24px 24px; }
        .sidebar-header h1 { font-size: 1.4em; }
        .sidebar-header p { opacity: 0.85; font-size: 0.9em; margin-top: 4px; }
        .menu-group-header {
            padding: 18px 24px 8px;
            font-size: 0.75em;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            opacity: 0.8;
        }
        .menu-item {
            padding: 12px 24px;
            cursor: pointer;
            opacity: 0.9;
        }
        .menu-item:hover, .menu-item.active { background: rgba(255,255,255,0.15); opacity: 1; }
        .main-content { margin-left: 280px; padding: 32px 40px; }
        .section {
            background: white;
            border-radius: 12px;
            padding: 32px;
            box-shadow: 0 8px 24px rgba(102, 126, 234, 0.12);
            width: 100%;
            max-width: none;
        }
        h2 { margin-bottom: 8px; }
        .hint { color: #666; margin-bottom: 24px; }
        .form-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 14px 28px;
            align-items: start;
        }
        .form-section-title {
            grid-column: 1 / -1;
            margin: 18px 0 0;
            padding-top: 8px;
            border-top: 1px solid #e8eaf6;
            font-size: 1.05em;
            color: #3949ab;
        }
        .form-section-title:first-child {
            border-top: none;
            padding-top: 0;
            margin-top: 0;
        }
        .form-group { margin-bottom: 0; }
        label { display: block; font-weight: 600; margin-bottom: 6px; font-size: 0.92em; }
        input[type="text"], input[type="file"] {
            width: 100%;
            padding: 10px 12px;
            border: 2px solid #ddd;
            border-radius: 8px;
            font-size: 15px;
        }
        input[type="text"]:focus { outline: none; border-color: #667eea; }
        .form-footer {
            grid-column: 1 / -1;
            margin-top: 8px;
        }
        .actions { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 8px; }
        @media (max-width: 980px) {
            .form-grid { grid-template-columns: 1fr; }
        }
        button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 12px 25px;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-size: 16px;
            font-weight: 600;
        }
        button.secondary { background: #5c6bc0; }
        .status {
            margin: 16px 0;
            padding: 12px 16px;
            border-radius: 8px;
            display: none;
        }
        .status.info { display: block; background: #eef2ff; color: #3730a3; }
        .status.error { display: block; background: #fdecea; color: #b71c1c; }
        .status.success { display: block; background: #e8f5e9; color: #1b5e20; }
        .upload-box {
            border: 2px dashed #c5cae9;
            border-radius: 10px;
            padding: 16px;
            margin: 20px 0 12px;
            background: #f8f9ff;
        }
        .modal-overlay {
            display: none;
            position: fixed;
            inset: 0;
            background: rgba(15, 23, 42, 0.55);
            z-index: 1000;
            align-items: center;
            justify-content: center;
            padding: 24px;
        }
        .modal-overlay.open { display: flex; }
        .modal {
            background: #fff;
            border-radius: 12px;
            width: min(920px, 100%);
            height: min(90vh, 860px);
            display: flex;
            flex-direction: column;
            box-shadow: 0 20px 50px rgba(0,0,0,0.25);
            overflow: hidden;
        }
        .modal-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 14px 18px;
            border-bottom: 1px solid #e5e7eb;
            background: #f8fafc;
        }
        .modal-header h3 { font-size: 1.05em; }
        .modal-close {
            background: #64748b;
            padding: 8px 14px;
            font-size: 14px;
        }
        .modal-body { flex: 1; min-height: 0; background: #e2e8f0; }
        .modal-body iframe {
            width: 100%;
            height: 100%;
            border: 0;
            background: #fff;
        }
    </style>
</head>
<body>
    <div class="sidebar">
        <div class="sidebar-header">
            <h1>Excel Tools</h1>
            <p>Automation Suite</p>
        </div>
        <div class="menu-group-header">Imagen Allocation Automation</div>
        <div class="menu-item" onclick="window.location.href='/comparison'">Comparison Tool</div>
        <div class="menu-item active">Imagen Row Form</div>
    </div>
    <div class="main-content">
        <div class="section">
            <h2>Imagen Row Form</h2>
            <p class="hint">Values from the allocation Link fill in automatically. Upload a PDF to fill any remaining blanks, Preview the Personal Information table, then Process to download the same PDF.</p>
            <div id="status" class="status"></div>
            <form id="imagen-row-form">
                <div class="form-grid">
                    <h3 class="form-section-title">Required from Link</h3>
                    {% for key, label in fields %}
                    <div class="form-group">
                        <label for="{{ key }}">{{ label }}</label>
                        <input type="text" id="{{ key }}" name="{{ key }}" value="{{ values.get(key, '') }}">
                    </div>
                    {% endfor %}
                    {% for section_title, section_fields in optional_sections %}
                    <h3 class="form-section-title">{{ section_title }}</h3>
                    {% for key, label in section_fields %}
                    <div class="form-group">
                        <label for="{{ key }}">{{ label }}</label>
                        <input type="text" id="{{ key }}" name="{{ key }}" value="{{ values.get(key, '') }}">
                    </div>
                    {% endfor %}
                    {% endfor %}
                    <div class="form-footer">
                        <div class="upload-box">
                            <div class="form-group" style="margin-bottom: 0;">
                                <label for="pdf">Upload PDF</label>
                                <input type="file" id="pdf" name="pdf" accept="application/pdf,.pdf">
                            </div>
                        </div>
                        <div class="actions">
                            <button type="button" class="secondary" id="extract-btn">Fill missing from PDF</button>
                            <button type="button" class="secondary" id="preview-btn">Preview</button>
                            <button type="submit" id="process-btn">Process</button>
                        </div>
                    </div>
                </div>
            </form>
        </div>
    </div>
    <div class="modal-overlay" id="preview-modal" aria-hidden="true">
        <div class="modal" role="dialog" aria-modal="true" aria-labelledby="preview-title">
            <div class="modal-header">
                <h3 id="preview-title">PDF Preview</h3>
                <button type="button" class="modal-close" id="preview-close">Close</button>
            </div>
            <div class="modal-body">
                <iframe id="preview-frame" title="Imagen row PDF preview"></iframe>
            </div>
        </div>
    </div>
    <script>
        const form = document.getElementById('imagen-row-form');
        const statusEl = document.getElementById('status');
        const previewModal = document.getElementById('preview-modal');
        const previewFrame = document.getElementById('preview-frame');
        let previewObjectUrl = null;

        function showStatus(message, kind) {
            statusEl.className = 'status ' + kind;
            statusEl.textContent = message;
        }

        function formDataWithoutPdf() {
            const data = new FormData(form);
            data.delete('pdf');
            return data;
        }

        function closePreview() {
            previewModal.classList.remove('open');
            previewModal.setAttribute('aria-hidden', 'true');
            previewFrame.removeAttribute('src');
            if (previewObjectUrl) {
                URL.revokeObjectURL(previewObjectUrl);
                previewObjectUrl = null;
            }
        }

        async function requestPdf(url) {
            const res = await fetch(url, { method: 'POST', body: formDataWithoutPdf() });
            if (!res.ok) {
                let message = 'Request failed.';
                try {
                    const payload = await res.json();
                    message = payload.error || message;
                } catch (err) {}
                throw new Error(message);
            }
            return res.blob();
        }

        document.getElementById('extract-btn').addEventListener('click', async function () {
            const pdfInput = document.getElementById('pdf');
            if (!pdfInput.files.length) {
                showStatus('Choose a PDF first.', 'error');
                return;
            }
            const data = new FormData(form);
            data.set('pdf', pdfInput.files[0]);
            showStatus('Reading PDF…', 'info');
            try {
                const res = await fetch('/imagen-row/extract-pdf', { method: 'POST', body: data });
                const payload = await res.json();
                if (!res.ok) {
                    showStatus(payload.error || 'Could not read PDF.', 'error');
                    return;
                }
                const values = payload.values || {};
                Object.keys(values).forEach(function (key) {
                    const input = document.getElementById(key);
                    if (input) input.value = values[key] || '';
                });
                const updated = payload.updated || [];
                if (!updated.length) {
                    showStatus('No missing fields were found in the PDF.', 'info');
                } else {
                    showStatus('Filled from PDF: ' + updated.join(', '), 'success');
                }
            } catch (err) {
                showStatus('Could not read PDF.', 'error');
            }
        });

        document.getElementById('preview-btn').addEventListener('click', async function () {
            showStatus('Building PDF preview…', 'info');
            try {
                const blob = await requestPdf('/imagen-row/preview');
                if (previewObjectUrl) URL.revokeObjectURL(previewObjectUrl);
                previewObjectUrl = URL.createObjectURL(blob);
                previewFrame.src = previewObjectUrl;
                previewModal.classList.add('open');
                previewModal.setAttribute('aria-hidden', 'false');
                showStatus('Preview ready.', 'success');
            } catch (err) {
                showStatus(err.message || 'Preview failed.', 'error');
            }
        });

        document.getElementById('preview-close').addEventListener('click', closePreview);
        previewModal.addEventListener('click', function (event) {
            if (event.target === previewModal) closePreview();
        });
        document.addEventListener('keydown', function (event) {
            if (event.key === 'Escape' && previewModal.classList.contains('open')) {
                closePreview();
            }
        });

        form.addEventListener('submit', async function (event) {
            event.preventDefault();
            showStatus('Creating PDF…', 'info');
            try {
                const blob = await requestPdf('/imagen-row/process');
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = 'imagen-row.pdf';
                document.body.appendChild(link);
                link.click();
                link.remove();
                URL.revokeObjectURL(url);
                showStatus('PDF downloaded.', 'success');
            } catch (err) {
                showStatus(err.message || 'Process failed.', 'error');
            }
        });
    </script>
</body>
</html>
"""


def render_page(values, error=""):
    from flask import render_template_string

    return render_template_string(
        PAGE_TEMPLATE,
        fields=FORM_FIELDS,
        optional_sections=OPTIONAL_FIELD_SECTIONS,
        values=values or {},
        error=error or "",
    )
