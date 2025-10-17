import os
import sys
import pathlib
from datetime import datetime
from typing import Dict, Any, List

from pyairtable import Api
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table as PdfTable, TableStyle, Image, Flowable
from reportlab.lib.styles import getSampleStyleSheet

# ---------------- ENV / CONFIG ----------------
AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ["AIRTABLE_BASE_ID"]
TRANSCRIPT_TABLE = os.environ.get("TRANSCRIPT_TABLE", "Students 1221")

# Branding & header
SCHOOL_NAME = os.environ.get("SCHOOL_NAME", "Cornerstone Education Management Organization")
SCHOOL_HEADER_RIGHT_LINE1 = os.environ.get("SCHOOL_HEADER_RIGHT_LINE1", "Cornerstone Online")
SCHOOL_HEADER_RIGHT_LINE2 = os.environ.get("SCHOOL_HEADER_RIGHT_LINE2", "Address line 1")
SCHOOL_HEADER_RIGHT_LINE3 = os.environ.get("SCHOOL_HEADER_RIGHT_LINE3", "City, State ZIP")

LOGO_PATH = os.environ.get("LOGO_PATH", "logo.png")   # put logo.png in repo root

# Signature (optional)
PRINCIPAL_NAME = os.environ.get("PRINCIPAL_NAME", "Principal Name")
SIGN_DATE_FMT = os.environ.get("SIGN_DATE_FMT", "%B %d, %Y")
SIGNATURE_PATH = os.environ.get("SIGNATURE_PATH", "")  # e.g., "signature.png" (optional)

# Record trigger
RECORD_ID = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else None)
if not RECORD_ID:
    print("[ERROR] Missing RECORD_ID")
    sys.exit(1)

# Airtable field names detected from your CSV
F = {
    "student_name": "Students Name",
    "student_id": "Student Canvas ID",
    "grade": "Grade",
    "school_year": "School Year",
    # Course details (rollups)
    "course_name_rollup": "Course Name Rollup (from Southlands Courses Enrollment 3)",
    "course_code_rollup": "Course Code Rollup (from Southlands Courses Enrollment 3)",
    # Optional (use if you later add them)
    "teacher": "Teacher",
    "s1": "S1",
    "s2": "S2",
    # Alternatives you already have
    "letter": "Grade Letter",
    "percent": "% Total",
}

api = Api(AIRTABLE_API_KEY)
table = api.table(AIRTABLE_BASE_ID, TRANSCRIPT_TABLE)

# ---------------- HELPERS ----------------
def sget(fields: Dict[str, Any], key: str, default: str = "") -> str:
    val = fields.get(key)
    if val is None:
        return default
    if isinstance(val, list):
        return ", ".join(str(x) for x in val if x)
    return str(val)

def listify(val: Any) -> List[str]:
    if val is None:
        return []
    if isinstance(val, list):
        return [str(x) for x in val if x is not None and str(x).strip()]
    # attempt to split comma-separated strings
    return [p.strip() for p in str(val).split(",") if p.strip()]

def fetch_student_group(student_name: str) -> List[Dict[str, Any]]:
    formula = f'{{{F["student_name"]}}} = "{student_name.replace(\'"\', "\\\"")}"'
    print(f"[DEBUG] filterByFormula: {formula}")
    return table.all(formula=formula)

# A tiny spacer line used under signature
class HLine(Flowable):
    def __init__(self, width=60*mm, thickness=0.6):
        Flowable.__init__(self)
        self.width = width
        self.thickness = thickness
        self.height = 2

    def draw(self):
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 0, self.width, 0)

# ---------------- PDF ----------------
def build_pdf(student_fields: Dict[str, Any], all_rows: List[Dict[str, Any]]):
    student_name = sget(student_fields, F["student_name"]).strip()
    student_id = sget(student_fields, F["student_id"])
    grade = sget(student_fields, F["grade"])
    year = sget(student_fields, F["school_year"])

    out_dir = pathlib.Path("output")
    out_dir.mkdir(parents=True, exist_ok=True)
    fname = f"transcript_{student_name.replace(' ', '_').replace(',', '')}_{year}.pdf"
    pdf_path = out_dir / fname

    styles = getSampleStyleSheet()
    title = styles["Title"]
    normal = styles["Normal"]
    h2 = styles["Heading2"]
    h2.spaceBefore, h2.spaceAfter = 4, 4

    story: List[Any] = []

    # --- Header row with 3 columns: Student box (left), Logo (center), School info (right) ---
    # 1) Student Info table
    info_data = [
        ["Student Info", ""],
        ["Name", student_name],
        ["Current Grade Level", grade],
        ["Student ID", student_id],
        # Address line is not stored; you can add one later if desired
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

    # 2) Center logo
    logo_cell: List[Flowable] = []
    if LOGO_PATH and pathlib.Path(LOGO_PATH).exists():
        try:
            logo_cell.append(Image(LOGO_PATH, width=38*mm, height=14*mm))
        except Exception as e:
            print(f"[WARN] Could not load logo: {e}")
    # Add school name below logo (centered)
    logo_cell.append(Spacer(1, 2*mm))
    logo_cell.append(Paragraph(f"<b>{SCHOOL_NAME}</b>", normal))

    # 3) Right school address
    right_block = []
    right_block.append(Paragraph(f"<b>{SCHOOL_HEADER_RIGHT_LINE1}</b>", normal))
    if SCHOOL_HEADER_RIGHT_LINE2:
        right_block.append(Paragraph(SCHOOL_HEADER_RIGHT_LINE2, normal))
    if SCHOOL_HEADER_RIGHT_LINE3:
        right_block.append(Paragraph(SCHOOL_HEADER_RIGHT_LINE3, normal))

    header_table = PdfTable(
        [
            [info, logo_cell, right_block]
        ],
        colWidths=[120*mm, 50*mm, 70*mm],
    )
    header_table.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 6*mm))

    # Centered "Report Card" / "For School Year ..."
    story.append(Paragraph("<b>Report Card</b>", h2))
    story.append(Paragraph(f"For School Year {year}", normal))
    story.append(Spacer(1, 6*mm))

    # --- Build course rows ---
    # We’ll expand rollups so each course has its own row.
    rows_out: List[List[str]] = []
    for r in all_rows:
        f = r.get("fields", {})
        # rollups can be comma-separated strings or arrays—handle both:
        names = listify(f.get(F["course_name_rollup"]))
        codes = listify(f.get(F["course_code_rollup"]))
        # optional fields (may be single values on the row)
        teacher = sget(f, F["teacher"], "")
        s1 = sget(f, F["s1"], "") or sget(f, "Grade Letter", "")  # fallback: use Grade Letter if you don't yet track S1/S2
        s2 = sget(f, F["s2"], "")
        # If there are multiple names in one row, pair best-effort with codes
        if names:
            for idx, nm in enumerate(names):
                code = codes[idx] if idx < len(codes) else ""
                rows_out.append([nm, code, teacher, s1, s2])
        else:
            # Single row fallback using whatever we have
            rows_out.append([
                sget(f, F["course_name_rollup"], ""),
                sget(f, F["course_code_rollup"], ""),
                teacher, s1, s2
            ])

    # Deduplicate repeated rows (common when multiple rows match)
    dedup = []
    seen = set()
    for row in rows_out:
        t = tuple(row)
        if t not in seen and any(x.strip() for x in row):
            seen.add(t)
            dedup.append(row)

    # --- Course table ---
    table_data = [["Course Name", "Course Number", "Teacher", "S1", "S2"]]
    table_data.extend(dedup if dedup else [["(no courses found)", "", "", "", ""]])

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

    # --- Signature block (optional) ---
    if SIGNATURE_PATH and pathlib.Path(SIGNATURE_PATH).exists():
        story.append(Image(SIGNATURE_PATH, width=45*mm, height=15*mm))
    story.append(HLine(width=60*mm))
    date_str = datetime.today().strftime(SIGN_DATE_FMT)
    story.append(Paragraph(f"Principal - {PRINCIPAL_NAME}", normal))
    story.append(Paragraph(f"Date: {date_str}", normal))

    # Build document (landscape A4 to match your sample proportions)
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

    # Handle linked/rollup student name types (array or string)
    raw_name = fields.get(F["student_name"])
    if isinstance(raw_name, list):
        student_name = raw_name[0] if raw_name else ""
    else:
        student_name = str(raw_name or "")

    if not student_name:
        raise SystemExit(f"[ERROR] Field '{F['student_name']}' is empty for record {RECORD_ID}")

    group = fetch_student_group(student_name)
    print(f"[INFO] Rows matched for '{student_name}': {len(group)}")
    if not group:
        # still generate using the single record for student info
        group = [rec]

    build_pdf(fields, group)

if __name__ == "__main__":
    main()
