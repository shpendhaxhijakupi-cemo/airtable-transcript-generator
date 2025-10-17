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

# ---------------- ENV / CONFIG ----------------
AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ["AIRTABLE_BASE_ID"]
TRANSCRIPT_TABLE = os.environ.get("TRANSCRIPT_TABLE", "Students 1221")

SCHOOL_NAME = os.environ.get("SCHOOL_NAME", "Cornerstone Education Management Organization")
SCHOOL_HEADER_RIGHT_LINE1 = os.environ.get("SCHOOL_HEADER_RIGHT_LINE1", "Cornerstone Online")
SCHOOL_HEADER_RIGHT_LINE2 = os.environ.get("SCHOOL_HEADER_RIGHT_LINE2", "")
SCHOOL_HEADER_RIGHT_LINE3 = os.environ.get("SCHOOL_HEADER_RIGHT_LINE3", "")
LOGO_PATH_ENV = os.environ.get("LOGO_PATH", "")  # e.g., logo_cornerstone.png

PRINCIPAL_NAME = os.environ.get("PRINCIPAL_NAME", "Principal Name")
SIGN_DATE_FMT = os.environ.get("SIGN_DATE_FMT", "%B %d, %Y")
SIGNATURE_PATH = os.environ.get("SIGNATURE_PATH", "")

RECORD_ID = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else None)
if not RECORD_ID:
    print("[ERROR] Missing RECORD_ID")
    sys.exit(1)

# Airtable field names
F = {
    "student_name": "Students Name",
    "student_id": "Student Canvas ID",
    "grade": "Grade",
    "school_year": "School Year",
    "course_name_rollup": "Course Name Rollup (from Southlands Courses Enrollment 3)",
    "course_code_rollup": "Course Code Rollup (from Southlands Courses Enrollment 3)",
    "teacher": "Teacher",     # optional if you add it later
    "s1": "S1",               # optional
    "s2": "S2",               # optional
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
        return ", ".join(str(x) for x in v if x is not None and str(x).strip())
    return str(v)

def listify(v: Any) -> List[str]:
    if v is None:
        return []
    if isinstance(v, list):
        return [str(x).strip() for x in v if x is not None and str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]

def escape_airtable_value(s: str) -> str:
    return (s or "").replace('"', '\\"')

def fetch_student_group(student_name: str) -> List[Dict[str, Any]]:
    field = F["student_name"]
    escaped = escape_airtable_value(student_name)
    formula = f'{{{field}}} = "{escaped}"'
    print(f"[DEBUG] filterByFormula: {formula}")
    return table.all(formula=formula)

def resolve_logo_path() -> str:
    # Priority: explicit env → common filenames
    candidates = [LOGO_PATH_ENV, "logo_cornerstone.png", "logo.png", "logo.jpg", "logo.jpeg"]
    for p in candidates:
        if p and pathlib.Path(p).exists():
            return p
    return ""  # not found

def guess_semester(course_name: str, course_code: str) -> Tuple[bool, bool]:
    """
    Returns (is_A, is_B) based on tokens in name/code:
    - English 7-A, ... A 2024/2025, ' A ' or '-A' => S1
    - English 7-B, ... B 2024/2025, ' B ' or '-B' => S2
    """
    text = f"{course_name} {course_code}".lower()
    # simple token checks
    has_a = ("-a" in text) or (" a " in text) or text.endswith(" a") or (" a/" in text)
    has_b = ("-b" in text) or (" b " in text) or text.endswith(" b") or (" b/" in text)
    return has_a and not has_b, has_b and not has_a

class HLine(Flowable):
    def __init__(self, width=60*mm, thickness=0.6):
        super().__init__()
        self.width, self.thickness = width, thickness
        self.height = 2
    def draw(self):
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 0, self.width, 0)

# ---------------- PDF ----------------
def build_pdf(student_fields: Dict[str, Any], all_rows: List[Dict[str, Any]]):
    student_name = sget(student_fields, F["student_name"]).strip()
    student_id   = sget(student_fields, F["student_id"])
    grade        = sget(student_fields, F["grade"])
    year         = sget(student_fields, F["school_year"])

    out_dir = pathlib.Path("output")
    out_dir.mkdir(parents=True, exist_ok=True)
    fname = f"transcript_{student_name.replace(' ', '_').replace(',', '')}_{year}.pdf"
    pdf_path = out_dir / fname

    styles = getSampleStyleSheet()
    normal, h2 = styles["Normal"], styles["Heading2"]
    h2.spaceBefore, h2.spaceAfter = 4, 4

    story: List[Any] = []

    # ---- Header row: Student box • Logo+name • Address
    info_data = [
        ["Student Info", ""],
        ["Name", student_name],
        ["Current Grade Level", grade],
        ["Student ID", student_id],
    ]
    info = PdfTable(info_data, colWidths=[45*mm, 70*mm])
    info.setStyle(TableStyle([
        ("SPAN", (0,0), (-1,0)),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BOX", (0,0), (-1,-1), 0.6, colors.black),
        ("INNERGRID", (0,0), (-1,-1), 0.3, colors.grey),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
    ]))

    center_cell: List[Any] = []
    logo_path = resolve_logo_path()
    if logo_path:
        try:
            center_cell.append(Image(logo_path, width=38*mm, height=14*mm))
        except Exception as e:
            print(f"[WARN] Could not load logo '{logo_path}': {e}")
    center_cell.append(Spacer(1, 2*mm))
    center_cell.append(Paragraph(f"<b>{SCHOOL_NAME}</b>", normal))

    right_block: List[Any] = [Paragraph(f"<b>{SCHOOL_HEADER_RIGHT_LINE1}</b>", normal)]
    if SCHOOL_HEADER_RIGHT_LINE2:
        right_block.append(Paragraph(SCHOOL_HEADER_RIGHT_LINE2, normal))
    if SCHOOL_HEADER_RIGHT_LINE3:
        right_block.append(Paragraph(SCHOOL_HEADER_RIGHT_LINE3, normal))

    header_table = PdfTable([[info, center_cell, right_block]], colWidths=[120*mm, 50*mm, 70*mm])
    header_table.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP")]))
    story.append(header_table)
    story.append(Spacer(1, 6*mm))

    story.append(Paragraph("<b>Report Card</b>", h2))
    story.append(Paragraph(f"For School Year {year}", normal))
    story.append(Spacer(1, 6*mm))

    # ---- Build course rows: expand rollups and map to S1/S2 using A/B
    table_data = [["Course Name", "Course Number", "Teacher", "S1", "S2"]]
    for r in all_rows:
        f = r.get("fields", {})
        names = listify(f.get(F["course_name_rollup"]))
        codes = listify(f.get(F["course_code_rollup"]))
        teacher = sget(f, F["teacher"], "")

        # grade value we can place into S1 or S2
        grade_value = sget(f, F["letter"]) or sget(f, F["percent"])
        if not names and not codes:
            # single-row fallback
            nm = sget(f, F["course_name_rollup"], "")
            cd = sget(f, F["course_code_rollup"], "")
            is_a, is_b = guess_semester(nm, cd)
            s1_val = grade_value if is_a or not (is_a or is_b) else ""
            s2_val = grade_value if is_b else ""
            table_data.append([nm, cd, teacher, s1_val, s2_val])
            continue

        # multiple items
        max_len = max(len(names), len(codes))
        for i in range(max_len):
            nm = names[i] if i < len(names) else ""
            cd = codes[i] if i < len(codes) else ""
            is_a, is_b = guess_semester(nm, cd)
            s1_val, s2_val = "", ""
            if is_a:
                s1_val = grade_value
            elif is_b:
                s2_val = grade_value
            else:
                # if we can't tell, put value in S1 as default
                s1_val = grade_value
            table_data.append([nm, cd, teacher, s1_val, s2_val])

    if len(table_data) == 1:
        table_data.append(["(no courses found)", "", "", "", ""])

    course_table = PdfTable(table_data, colWidths=[90*mm, 35*mm, 60*mm, 15*mm, 15*mm])
    course_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("BOTTOMPADDING", (0,0), (-1,0), 6),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.lightgrey]),
    ]))
    story.append(course_table)
    story.append(Spacer(1, 16*mm))

    if SIGNATURE_PATH and pathlib.Path(SIGNATURE_PATH).exists():
        story.append(Image(SIGNATURE_PATH, width=45*mm, height=15*mm))
    story.append(HLine(width=60*mm))
    story.append(Paragraph(f"Principal - {PRINCIPAL_NAME}", normal))
    story.append(Paragraph(f"Date: {datetime.today().strftime(SIGN_DATE_FMT)}", normal))

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=landscape(A4),
        leftMargin=15*mm, rightMargin=15*mm, topMargin=12*mm, bottomMargin=12*mm
    )
    doc.build(story)
    print(f"[OK] Generated → {pdf_path}")
    return pdf_path

# ---------------- MAIN ----------------
def main():
    print(f"[INFO] Fetching record: {RECORD_ID} from table '{TRANSCRIPT_TABLE}'")
    rec = table.get(RECORD_ID)
    if not rec or "fields" not in rec:
        raise SystemExit("[ERROR] Could not fetch the record or fields are empty.")
    fields = rec["fields"]

    raw_name = fields.get(F["student_name"])
    student_name = raw_name[0] if isinstance(raw_name, list) and raw_name else (str(raw_name or ""))
    if not student_name:
        raise SystemExit(f"[ERROR] Field '{F['student_name']}' is empty for record {RECORD_ID}")

    group = fetch_student_group(student_name)
    print(f"[INFO] Rows matched for '{student_name}': {len(group)}")
    if not group:
        group = [rec]  # still create a PDF

    build_pdf(fields, group)

if __name__ == "__main__":
    main()
