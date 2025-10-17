import os
import sys
import pathlib
from datetime import datetime
from typing import Dict, Any, List, Tuple

from pyairtable import Api
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table as PdfTable, TableStyle, Image, Flowable
)
from reportlab.lib.styles import getSampleStyleSheet

# ---------------- ENV CONFIG ----------------
AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ["AIRTABLE_BASE_ID"]
TRANSCRIPT_TABLE = os.environ.get("TRANSCRIPT_TABLE", "Students 1221")

SCHOOL_NAME = os.environ.get("SCHOOL_NAME", "Cornerstone Education Management Organization")
SCHOOL_HEADER_RIGHT_LINE1 = os.environ.get("SCHOOL_HEADER_RIGHT_LINE1", "Cornerstone Online")
SCHOOL_HEADER_RIGHT_LINE2 = os.environ.get("SCHOOL_HEADER_RIGHT_LINE2", "18550 Fajarado St.")
SCHOOL_HEADER_RIGHT_LINE3 = os.environ.get("SCHOOL_HEADER_RIGHT_LINE3", "Rowland Heights, CA 91748")

LOGO_PATH = os.environ.get("LOGO_PATH", "logo_cornerstone.png")  # your logo
SIGNATURE_PATH = os.environ.get("SIGNATURE_PATH", "signature_principal.png")  # your signature
PRINCIPAL_NAME = os.environ.get("PRINCIPAL_NAME", "Ursula Derios")
SIGN_DATE_FMT = os.environ.get("SIGN_DATE_FMT", "%B %d, %Y")

RECORD_ID = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else None)
if not RECORD_ID:
    sys.exit("[ERROR] Missing RECORD_ID")

F = {
    "student_name": "Students Name",
    "student_id": "Student Canvas ID",
    "grade": "Grade",
    "school_year": "School Year",
    "course_name_rollup": "Course Name Rollup (from Southlands Courses Enrollment 3)",
    "course_code_rollup": "Course Code Rollup (from Southlands Courses Enrollment 3)",
    "teacher": "Teacher",
    "letter": "Grade Letter",
    "percent": "% Total",
}

api = Api(AIRTABLE_API_KEY)
table = api.table(AIRTABLE_BASE_ID, TRANSCRIPT_TABLE)

# ---------------- HELPERS ----------------
def sget(fields: Dict[str, Any], key: str, default: str = "") -> str:
    v = fields.get(key)
    if v is None:
        return default
    if isinstance(v, list):
        return ", ".join(str(x) for x in v if x)
    return str(v)

def listify(v: Any) -> List[str]:
    if v is None:
        return []
    if isinstance(v, list):
        return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]

def escape_airtable_value(s: str) -> str:
    return (s or "").replace('"', '\\"')

def fetch_student_group(student_name: str) -> List[Dict[str, Any]]:
    field = F["student_name"]
    escaped = escape_airtable_value(student_name)
    formula = f'{{{field}}} = "{escaped}"'
    print(f"[DEBUG] filterByFormula: {formula}")
    return table.all(formula=formula)

def guess_semester(course_name: str, course_code: str):
    t = (course_name + " " + course_code).lower()
    return (" a" in t or "-a" in t), (" b" in t or "-b" in t)

class HLine(Flowable):
    def __init__(self, width=60*mm, thickness=0.6):
        super().__init__()
        self.width, self.thickness = width, thickness
    def draw(self):
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 0, self.width, 0)

# ---------------- PDF GENERATOR ----------------
def build_pdf(student_fields: Dict[str, Any], all_rows: List[Dict[str, Any]]):
    student_name = sget(student_fields, F["student_name"]).strip()
    student_id = sget(student_fields, F["student_id"])
    grade = sget(student_fields, F["grade"])
    year = sget(student_fields, F["school_year"])

    out_dir = pathlib.Path("output")
    out_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = out_dir / f"transcript_{student_name.replace(' ', '_')}.pdf"

    styles = getSampleStyleSheet()
    normal = styles["Normal"]
    h2 = styles["Heading2"]
    h2.alignment = 1  # center headings

    story = []

    # --- Header Logo ---
    if pathlib.Path(LOGO_PATH).exists():
        story.append(Image(LOGO_PATH, width=55*mm, height=20*mm))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(f"<b>{SCHOOL_NAME}</b>", normal))
    story.append(Paragraph(f"{SCHOOL_HEADER_RIGHT_LINE1}", normal))
    story.append(Paragraph(f"{SCHOOL_HEADER_RIGHT_LINE2}", normal))
    story.append(Paragraph(f"{SCHOOL_HEADER_RIGHT_LINE3}", normal))
    story.append(Spacer(1, 10*mm))

    # --- Student Info ---
    info = PdfTable([
        ["Student Info", ""],
        ["Name", student_name],
        ["Current Grade Level", grade],
        ["Student ID", student_id],
    ], colWidths=[50*mm, 80*mm])
    info.setStyle(TableStyle([
        ("SPAN", (0,0), (-1,0)),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BOX", (0,0), (-1,-1), 0.8, colors.black),
        ("INNERGRID", (0,0), (-1,-1), 0.4, colors.grey),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
    ]))
    story.append(info)
    story.append(Spacer(1, 12*mm))

    # --- Title ---
    story.append(Paragraph("<b>Official Transcript</b>", h2))
    story.append(Paragraph(f"For School Year {year}", normal))
    story.append(Spacer(1, 8*mm))

    # --- Courses ---
    table_data = [["Course Name", "Course Number", "Teacher", "S1", "S2"]]
    for r in all_rows:
        f = r.get("fields", {})
        names = listify(f.get(F["course_name_rollup"]))
        codes = listify(f.get(F["course_code_rollup"]))
        teacher = sget(f, F["teacher"])
        grade_value = sget(f, F["letter"]) or sget(f, F["percent"])
        for name, code in zip(names, codes or [""] * len(names)):
            a, b = guess_semester(name, code)
            s1, s2 = ("F", "") if a else ("", "F") if b else (grade_value, "")
            table_data.append([name, code, teacher, s1, s2])

    table = PdfTable(table_data, colWidths=[90*mm, 40*mm, 50*mm, 15*mm, 15*mm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.lightgrey]),
    ]))
    story.append(table)
    story.append(Spacer(1, 16*mm))

    # --- Signature block centered ---
    if pathlib.Path(SIGNATURE_PATH).exists():
        story.append(Image(SIGNATURE_PATH, width=50*mm, height=18*mm))
    story.append(HLine(width=60*mm))
    story.append(Paragraph(f"Principal - {PRINCIPAL_NAME}", normal))
    story.append(Paragraph(f"Date: {datetime.today().strftime(SIGN_DATE_FMT)}", normal))

    doc = SimpleDocTemplate(str(pdf_path), pagesize=landscape(A4),
                            leftMargin=20*mm, rightMargin=20*mm, topMargin=15*mm, bottomMargin=15*mm)
    doc.build(story)
    print(f"[OK] Generated → {pdf_path}")
    return pdf_path


def main():
    print(f"[INFO] Fetching record: {RECORD_ID}")
    rec = table.get(RECORD_ID)
    if not rec or "fields" not in rec:
        sys.exit("[ERROR] Could not fetch record or empty fields.")
    fields = rec["fields"]
    name = fields.get(F["student_name"])
    if isinstance(name, list):
        name = name[0]
    group = fetch_student_group(name)
    if not group:
        group = [rec]
    build_pdf(fields, group)

if __name__ == "__main__":
    main()
