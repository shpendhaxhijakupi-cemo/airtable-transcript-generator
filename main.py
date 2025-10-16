import os
import sys
import io
import pathlib
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

# ---------- Config ----------
AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ["AIRTABLE_BASE_ID"]
TRANSCRIPT_TABLE = os.environ.get("TRANSCRIPT_TABLE", "Students 1221")
SCHOOL_NAME = os.environ.get(
    "SCHOOL_NAME", "Cornerstone Education Management Organization"
)
LOGO_PATH = "logo.png"  # same folder as main.py
RECORD_ID = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else None)
if not RECORD_ID:
    print("Missing RECORD_ID")
    sys.exit(1)

# ---------- Airtable ----------
api = Api(AIRTABLE_API_KEY)
table = api.table(AIRTABLE_BASE_ID, TRANSCRIPT_TABLE)

def get_student_name(record):
    return record["fields"].get("Students name") or "Unknown"

def fetch_records_for_student(student_name):
    print(f"[INFO] Fetching records for {student_name}")
    formula = f"{{Students name}} = '{student_name}'"
    return table.all(formula=formula)

# ---------- PDF ----------
def build_transcript(student_name, rows):
    if not rows:
        print("[WARN] No rows for this student.")
        return None

    first = rows[0]["fields"]
    sid = first.get("Student Canvas ID", "")
    grade = first.get("Grade", "")
    year = first.get("School Year", "")
    out_dir = pathlib.Path("output")
    out_dir.mkdir(parents=True, exist_ok=True)
    safe_name = student_name.replace(" ", "_").replace(",", "")
    pdf_path = out_dir / f"transcript_{safe_name}_{RECORD_ID}.pdf"

    styles = getSampleStyleSheet()
    story = []

    # Header
    if pathlib.Path(LOGO_PATH).exists():
        story.append(Image(LOGO_PATH, width=60 * mm, height=20 * mm))
    story.append(Paragraph(f"<b>{SCHOOL_NAME}</b>", styles["Title"]))
    story.append(Paragraph("Official Transcript", styles["h2"]))
    story.append(Paragraph(f"For School Year {year}", styles["Normal"]))
    story.append(Spacer(1, 6 * mm))

    # Student Info box
    info_data = [
        ["Student Name", student_name],
        ["Student Canvas ID", sid],
        ["Grade", grade],
        ["School Year", year],
    ]
    info_table = PdfTable(info_data, colWidths=[50 * mm, 90 * mm])
    info_table.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
                ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    story.append(info_table)
    story.append(Spacer(1, 8 * mm))

    # Courses Table
    header = ["Course Name", "Grade Letter", "% Total"]
    table_data = [header]
    for r in rows:
        f = r.get("fields", {})
        cname = f.get("Course Name Rollup", "")
        gletter = f.get("Grade Letter", "")
        pct = f.get("%Total", "")
        table_data.append([cname, gletter, pct])

    course_table = PdfTable(table_data, colWidths=[80 * mm, 40 * mm, 40 * mm])
    course_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 1), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
            ]
        )
    )
    story.append(course_table)
    story.append(Spacer(1, 10 * mm))
    story.append(
        Paragraph(
            "This transcript is generated from the Cornerstone SIS and reflects the official record as of the run date.",
            styles["Italic"],
        )
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

def main():
    student_record = table.get(RECORD_ID)
    student_name = get_student_name(student_record)
    student_rows = fetch_records_for_student(student_name)
    build_transcript(student_name, student_rows)

if __name__ == "__main__":
    main()
