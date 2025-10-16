import os
import sys
import pathlib
from typing import Dict, List, Any
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

# ---------------- CONFIG ----------------
AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ["AIRTABLE_BASE_ID"]
TRANSCRIPT_TABLE = os.environ.get("TRANSCRIPT_TABLE", "Students 1221")
SCHOOL_NAME = os.environ.get("SCHOOL_NAME", "Cornerstone Education Management Organization")
LOGO_PATH = os.environ.get("LOGO_PATH") or "logo_cornerstone.png"

RECORD_ID = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else None)
if not RECORD_ID:
    print("[ERROR] Missing RECORD_ID")
    sys.exit(1)

FIELDS = {
    "student_name": "Students Name",
    "student_id": "Student Canvas ID",
    "grade": "Grade",
    "school_year": "School Year",
    "percent": "% Total",
    "letter": "Grade Letter",
    "course": "Course Name Rollup (from Southlands Courses Enrollment 3)"
}

api = Api(AIRTABLE_API_KEY)
table = api.table(AIRTABLE_BASE_ID, TRANSCRIPT_TABLE)

# ---------------- HELPERS ----------------
def safe_get(fields: Dict[str, Any], key: str) -> str:
    val = fields.get(key)
    if val is None:
        return ""
    if isinstance(val, list):
        return ", ".join(str(x) for x in val if x)
    return str(val)

def make_pdf(student_fields: Dict[str, Any], courses: List[Dict[str, Any]]):
    student_name = safe_get(student_fields, FIELDS["student_name"]).strip()
    school_year = safe_get(student_fields, FIELDS["school_year"])
    student_id = safe_get(student_fields, FIELDS["student_id"])
    grade = safe_get(student_fields, FIELDS["grade"])

    out_dir = pathlib.Path("output")
    out_dir.mkdir(parents=True, exist_ok=True)
    filename = f"transcript_{student_name.replace(' ', '_')}_{school_year}.pdf"
    pdf_path = out_dir / filename

    styles = getSampleStyleSheet()
    story = []

    # Header
    if LOGO_PATH and pathlib.Path(LOGO_PATH).exists():
        story.append(Image(LOGO_PATH, width=60 * mm, height=20 * mm))
    story.append(Paragraph(f"<b>{SCHOOL_NAME}</b>", styles["Title"]))
    story.append(Paragraph("Official Transcript", styles["h2"]))
    story.append(Paragraph(f"For School Year {school_year}", styles["Normal"]))
    story.append(Spacer(1, 8 * mm))

    # Student Info
    info = [
        ["Student Name", student_name],
        ["Student Canvas ID", student_id],
        ["Grade", grade],
        ["School Year", school_year],
    ]
    info_table = PdfTable(info, colWidths=[50 * mm, 90 * mm])
    info_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 10 * mm))

    # Course Table
    table_data = [["Course Name", "Grade Letter", "% Total"]]
    for row in courses:
        table_data.append([
            safe_get(row["fields"], FIELDS["course"]),
            safe_get(row["fields"], FIELDS["letter"]),
            safe_get(row["fields"], FIELDS["percent"]),
        ])

    course_table = PdfTable(table_data, colWidths=[90 * mm, 40 * mm, 40 * mm])
    course_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
    ]))
    story.append(course_table)
    story.append(Spacer(1, 10 * mm))

    story.append(Paragraph(
        "This transcript is generated from the Cornerstone SIS and reflects the official record as of the run date.",
        styles["Italic"]
    ))

    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4,
                            leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=18 * mm, bottomMargin=18 * mm)
    doc.build(story)
    print(f"[OK] Generated: {pdf_path}")


# ---------------- MAIN ----------------
def main():
    print(f"[INFO] Reading record {RECORD_ID} from {TRANSCRIPT_TABLE}")
    student_rec = table.get(RECORD_ID)
    fields = student_rec.get("fields", {})
    student_name = safe_get(fields, FIELDS["student_name"])
    if not student_name:
        raise SystemExit(f"[ERROR] Missing student name in record {RECORD_ID}")

    formula = f"{{{FIELDS['student_name']}}} = '{student_name}'"
    rows = table.all(formula=formula)
    print(f"[INFO] Found {len(rows)} records for {student_name}")

    make_pdf(fields, rows)


if __name__ == "__main__":
    main()
