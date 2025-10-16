import os
import sys
import pathlib
from typing import List, Dict, Any

from pyairtable import Api
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table as PdfTable,
    TableStyle,
    Image,
)
from reportlab.lib.styles import getSampleStyleSheet

# ---------------- Env / Config ----------------
AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ["AIRTABLE_BASE_ID"]
TRANSCRIPT_TABLE = os.environ.get("TRANSCRIPT_TABLE", "Students 1221")
SCHOOL_NAME = os.environ.get("SCHOOL_NAME", "Cornerstone Education Management Organization")

# Prefer explicitly provided logo path; otherwise try common filenames
LOGO_PATH = os.environ.get("LOGO_PATH") or (
    "logo.png" if pathlib.Path("logo.png").exists() else
    ("logo_cornerstone.png" if pathlib.Path("logo_cornerstone.png").exists() else None)
)

RECORD_ID = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else None)
if not RECORD_ID:
    print("[ERROR] Missing RECORD_ID")
    sys.exit(1)

# Field names in your table
FIELD_STUDENT_NAME = "Students name"
FIELD_STUDENT_ID   = "Student Canvas ID"
FIELD_SCHOOL_YEAR  = "School Year"
FIELD_GRADE_LEVEL  = "Grade"
FIELD_COURSE_NAME  = "Course Name Rollup"  # string or array depending on your rollup
FIELD_GRADE_LETTER = "Grade Letter"
FIELD_PERCENT_TOTAL= "%Total"

api = Api(AIRTABLE_API_KEY)
table = api.table(AIRTABLE_BASE_ID, TRANSCRIPT_TABLE)

# ---------------- Helpers ----------------
def _get(fields: Dict[str, Any], key: str, default: str = "") -> str:
    v = fields.get(key)
    if v is None:
        return default
    # convert list rollups to comma-separated for display
    if isinstance(v, list):
        return ", ".join(str(x) for x in v if x is not None)
    return str(v)

def escape_formula_value(val: str) -> str:
    # Use double quotes in Airtable formula to avoid issues with apostrophes
    return val.replace('"', '\\"')

def fetch_primary_record(record_id: str) -> Dict[str, Any]:
    print(f"[INFO] Using table: {TRANSCRIPT_TABLE}")
    print(f"[INFO] Fetching primary record: {record_id}")
    rec = table.get(record_id)
    if not rec or "fields" not in rec:
        raise SystemExit(f"[ERROR] Could not load record {record_id} from '{TRANSCRIPT_TABLE}'")
    return rec

def fetch_all_rows_for_student(student_name: str) -> List[Dict[str, Any]]:
    # Pull all records in the same table that belong to this student
    # NOTE: We use double quotes in the filter formula
    formula = f'{{{FIELD_STUDENT_NAME}}} = "{escape_formula_value(student_name)}"'
    print(f"[DEBUG] filterByFormula: {formula}")
    rows = table.all(formula=formula)
    print(f"[INFO] Matched rows for '{student_name}': {len(rows)}")
    return rows

def build_pdf(student_name: str, sample_fields: Dict[str, Any], rows: List[Dict[str, Any]], record_id: str) -> pathlib.Path:
    out_dir = pathlib.Path("output")
    out_dir.mkdir(parents=True, exist_ok=True)
    safe_name = student_name.replace(" ", "_").replace(",", "")
    pdf_path = out_dir / f"transcript_{safe_name}_{record_id}.pdf"

    styles = getSampleStyleSheet()
    story = []

    # Header
    if LOGO_PATH and pathlib.Path(LOGO_PATH).exists():
        try:
            story.append(Image(LOGO_PATH, width=60 * mm, height=20 * mm))
        except Exception as e:
            print(f"[WARN] Could not load logo '{LOGO_PATH}': {e}")
    story.append(Paragraph(f"<b>{SCHOOL_NAME}</b>", styles["Title"]))
    story.append(Paragraph("Official Transcript", styles["h2"]))

    school_year = _get(sample_fields, FIELD_SCHOOL_YEAR)
    if school_year:
        story.append(Paragraph(f"For School Year {school_year}", styles["Normal"]))
    story.append(Spacer(1, 6 * mm))

    # Student info box
    info_data = [
        ["Student Name", student_name],
        ["Student Canvas ID", _get(sample_fields, FIELD_STUDENT_ID)],
        ["Grade", _get(sample_fields, FIELD_GRADE_LEVEL)],
        ["School Year", school_year],
    ]
    info_table = PdfTable(info_data, colWidths=[50 * mm, 90 * mm])
    info_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 8 * mm))

    # Courses table
    table_data = [["Course Name", "Grade Letter", "% Total"]]
    if rows:
        # Many bases store one course per row for a student.
        # If your rollup returns an array/string, we still take row-level Grade Letter / %Total.
        for r in rows:
            f = r.get("fields", {})
            course_name = _get(f, FIELD_COURSE_NAME)
            grade_letter = _get(f, FIELD_GRADE_LETTER)
            pct_total = _get(f, FIELD_PERCENT_TOTAL)
            # If course_name is empty (rare), still show something
            table_data.append([course_name or "(course)", grade_letter, pct_total])
    else:
        table_data.append(["(no courses found)", "", ""])

    course_table = PdfTable(table_data, colWidths=[80 * mm, 40 * mm, 40 * mm])
    course_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
    ]))
    story.append(course_table)
    story.append(Spacer(1, 10 * mm))
    story.append(Paragraph(
        "This transcript is generated from the Cornerstone SIS and reflects the official record as of the run date.",
        styles["Italic"])
    )

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
    )
    doc.build(story)
    print(f"[OK] Transcript written → {pdf_path}")
    return pdf_path


# ---------------- Main ----------------
def main():
    # 1) Load the single record by ID (must belong to TRANSCRIPT_TABLE)
    rec = fetch_primary_record(RECORD_ID)
    fields = rec.get("fields", {})
    student_name = fields.get(FIELD_STUDENT_NAME)
    if not student_name:
        raise SystemExit(f"[ERROR] Field '{FIELD_STUDENT_NAME}' is empty on record {RECORD_ID}")

    print(f"[INFO] Student name: {student_name!r}")

    # 2) Pull ALL rows for that student (same table)
    rows = fetch_all_rows_for_student(student_name)
    if not rows:
        print("[ERROR] No rows matched this student. Check field names / values.")
        sys.exit(2)

    # 3) Generate PDF
    pdf_path = build_pdf(student_name, rows[0].get("fields", {}), rows, RECORD_ID)
    if not pdf_path or not pathlib.Path(pdf_path).exists():
        raise SystemExit("[ERROR] PDF not generated.")

if __name__ == "__main__":
    main()
