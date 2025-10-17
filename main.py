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
    SimpleDocTemplate, Paragraph, Spacer, Table as PdfTable, TableStyle, Image, Flowable, KeepInFrame
)
from reportlab.lib.styles import getSampleStyleSheet

# ---------- CONFIG (env) ----------
AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ["AIRTABLE_BASE_ID"]
TRANSCRIPT_TABLE = os.environ.get("TRANSCRIPT_TABLE", "Students 1221")

SCHOOL_NAME  = os.environ.get("SCHOOL_NAME", "Cornerstone Education Management Organization")
ADDR_LINE_1  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE1", "Cornerstone Online")
ADDR_LINE_2  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE2", "18550 Fajarado St.")
ADDR_LINE_3  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE3", "Rowland Heights, CA 91748")

LOGO_PATH    = os.environ.get("LOGO_PATH", "logo_cornerstone.png")
SIGN_PATH    = os.environ.get("SIGNATURE_PATH", "signature_principal.png")
PRINCIPAL    = os.environ.get("PRINCIPAL_NAME", "Ursula Derios")
SIGN_DATEFMT = os.environ.get("SIGN_DATE_FMT", "%B %d, %Y")

RECORD_ID = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else None)
if not RECORD_ID:
    sys.exit("[ERROR] Missing RECORD_ID")

# Airtable fields (from your base/CSV)
F = {
    "student_name": "Students Name",
    "student_id"  : "Student Canvas ID",
    "grade"       : "Grade",
    "school_year" : "School Year",
    "course_name_rollup": "Course Name Rollup (from Southlands Courses Enrollment 3)",
    "course_code_rollup": "Course Code Rollup (from Southlands Courses Enrollment 3)",
    "teacher"     : "Teacher",       # optional
    "letter"      : "Grade Letter",
    "percent"     : "% Total",
}

api = Api(AIRTABLE_API_KEY)
table = api.table(AIRTABLE_BASE_ID, TRANSCRIPT_TABLE)

# ---------- helpers ----------
def sget(fields: Dict[str, Any], key: str, default: str = "") -> str:
    v = fields.get(key)
    if v is None:
        return default
    if isinstance(v, list):
        return ", ".join(str(x) for x in v if x and str(x).strip())
    return str(v)

def listify(v: Any) -> List[str]:
    if v is None:
        return []
    if isinstance(v, list):
        return [str(x).strip() for x in v if x and str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]

def escape_airtable_value(s: str) -> str:
    return (s or "").replace('"', '\\"')

def fetch_group(student_name: str) -> List[Dict[str, Any]]:
    field = F["student_name"]
    formula = f'{{{field}}} = "{escape_airtable_value(student_name)}"'
    print(f"[DEBUG] filterByFormula: {formula}")
    return table.all(formula=formula)

def detect_semester(name: str, code: str) -> Tuple[bool, bool]:
    t = f"{name} {code}".lower()
    is_a = ("-a" in t) or (" a " in t) or t.endswith(" a")
    is_b = ("-b" in t) or (" b " in t) or t.endswith(" b")
    return (is_a and not is_b, is_b and not is_a)

class CenterLine(Flowable):
    def __init__(self, width=70*mm, thickness=0.7):
        super().__init__()
        self.width, self.thickness = width, thickness
        self.height = 3
    def draw(self):
        self.canv.setLineWidth(self.thickness)
        self.canv.line(-self.width/2, 0, self.width/2, 0)

# ---------- PDF ----------
def build_pdf(student_fields: Dict[str, Any], rows: List[Dict[str, Any]]):
    student_name = sget(student_fields, F["student_name"]).strip()
    student_id   = sget(student_fields, F["student_id"])
    grade        = sget(student_fields, F["grade"])
    year         = sget(student_fields, F["school_year"])

    out = pathlib.Path("output"); out.mkdir(parents=True, exist_ok=True)
    pdf_path = out / f"transcript_{student_name.replace(' ', '_').replace(',', '')}_{year}.pdf"

    styles   = getSampleStyleSheet()
    normal   = styles["Normal"]
    h2       = styles["Heading2"]; h2.alignment = 1
    title    = styles["Title"];    title.alignment = 1

    story: List[Any] = []

    # ---- HEADER: Left info • Center logo/name • Right address
    # left: school block (not bold except first line)
    school_left = [
        Paragraph(f"<b>{SCHOOL_NAME}</b>", normal),
        Paragraph(ADDR_LINE_1, normal),
        Paragraph(ADDR_LINE_2, normal),
        Paragraph(ADDR_LINE_3, normal),
    ]

    # center: logo + school wordmark
    center_cells: List[Any] = []
    if pathlib.Path(LOGO_PATH).exists():
        center_cells.append(Image(LOGO_PATH, width=42*mm, height=16*mm))
    center_cells.append(Spacer(1, 2*mm))

    # right: empty (kept for symmetry or future use)
    right_cells: List[Any] = []

    header_tbl = PdfTable([[school_left, center_cells, right_cells]],
                          colWidths=[110*mm, 70*mm, 60*mm])
    header_tbl.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP")]))
    story.append(header_tbl)
    story.append(Spacer(1, 6*mm))

    # ---- Student Info box (LEFT) + logo already centered above → match Southlands order
    info = PdfTable([
        ["Student Info", ""],
        ["Name",               student_name],
        ["Current Grade Level", grade],
        ["Student ID",         student_id],
    ], colWidths=[55*mm, 85*mm])
    info.setStyle(TableStyle([
        ("SPAN", (0,0), (-1,0)),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9D9D9")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BOX", (0,0), (-1,-1), 0.9, colors.black),
        ("INNERGRID", (0,0), (-1,-1), 0.4, colors.grey),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
    ]))
    # put the info box at left using a 3-col table (info • spacer • spacer)
    info_row = PdfTable([[info, "", ""]], colWidths=[140*mm, 40*mm, 60*mm])
    info_row.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(info_row)
    story.append(Spacer(1, 10*mm))

    # ---- Titles centered
    story.append(Paragraph("Report Card", h2))
    story.append(Paragraph(f"For School Year {year}", normal))
    story.append(Spacer(1, 8*mm))

    # ---- Build Course rows
    table_data = [["Course Name", "Course Number", "Teacher", "S1", "S2"]]

    expanded = []
    for r in rows:
        f = r.get("fields", {})
        names   = listify(f.get(F["course_name_rollup"]))
        codes   = listify(f.get(F["course_code_rollup"]))
        teacher = sget(f, F["teacher"])
        grade_val = sget(f, F["letter"]) or sget(f, F["percent"])

        # pair up name/code lists (handle unequal lengths)
        maxlen = max(len(names), len(codes))
        for i in range(maxlen):
            nm = names[i] if i < len(names) else ""
            cd = codes[i] if i < len(codes) else ""
            a, b = detect_semester(nm, cd)
            s1, s2 = ("", "")
            if a:
                s1 = grade_val or "—"
            elif b:
                s2 = grade_val or "—"
            else:
                s1 = grade_val or "—"  # default S1
            expanded.append([nm, cd, teacher, s1, s2])

    # clean + sort: unique, then A before B where names match
    seen = set(); clean = []
    for row in expanded:
        t = tuple(row)
        if t not in seen and any(x.strip() for x in row):
            seen.add(t); clean.append(row)
    clean.sort(key=lambda x: (x[0].lower(), x[1].lower()))

    table_data.extend(clean if clean else [["(no courses found)", "", "", "", ""]])

    courses = PdfTable(table_data, colWidths=[105*mm, 40*mm, 60*mm, 15*mm, 15*mm])
    courses.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9D9D9")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.35, colors.grey),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.whitesmoke, colors.HexColor("#EFEFEF")]),
        ("FONTSIZE", (0,0), (-1,-1), 9.5),
        ("BOTTOMPADDING", (0,0), (-1,0), 6),
        ("TOPPADDING", (0,1), (-1,-1), 4),
        ("BOTTOMPADDING", (0,1), (-1,-1), 4),
    ]))
    # wrap the table to keep it neat if many rows
    story.append(KeepInFrame(0, 9999, [courses], mode="shrink"))
    story.append(Spacer(1, 20*mm))

    # ---- Centered signature block
    signature_stack: List[Any] = []
    if pathlib.Path(SIGN_PATH).exists():
        signature_stack.append(Image(SIGN_PATH, width=55*mm, height=18*mm))
        signature_stack.append(Spacer(1, 3*mm))
    signature_stack.append(CenterLine())
    signature_stack.append(Spacer(1, 2*mm))
    signature_stack.append(Paragraph(f"Principal - {PRINCIPAL}", normal))
    signature_stack.append(Paragraph(f"Date: {datetime.today().strftime(SIGN_DATEFMT)}", normal))

    # place the signature stack centered
    story.append(PdfTable([[signature_stack]], colWidths=[210*mm]))
    story.append(Spacer(1, 3*mm))

    # Optional footer (uncomment if you want it)
    # story.append(Paragraph("<i>Generated from the Cornerstone SIS</i>", normal))

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=landscape(A4),
        leftMargin=14*mm, rightMargin=14*mm, topMargin=12*mm, bottomMargin=12*mm,
    )
    doc.build(story)
    print(f"[OK] Generated → {pdf_path}")
    return pdf_path

# ---------- main ----------
def main():
    print(f"[INFO] Fetching record: {RECORD_ID} from '{TRANSCRIPT_TABLE}'")
    rec = table.get(RECORD_ID)
    if not rec or "fields" not in rec:
        sys.exit("[ERROR] Could not fetch record or empty fields.")
    fields = rec["fields"]

    raw_name = fields.get(F["student_name"])
    student_name = raw_name[0] if isinstance(raw_name, list) and raw_name else str(raw_name or "")
    if not student_name:
        sys.exit(f"[ERROR] Field '{F['student_name']}' is empty.")

    group = fetch_group(student_name)
    print(f"[INFO] Rows matched: {len(group)}")
    if not group:
        group = [rec]

    build_pdf(fields, group)

if __name__ == "__main__":
    main()
