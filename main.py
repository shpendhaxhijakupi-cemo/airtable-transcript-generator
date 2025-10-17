import os, sys, pathlib
from datetime import datetime
from typing import Dict, Any, List, Tuple
from pyairtable import Api
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table as PdfTable, TableStyle, Image, Flowable
from reportlab.lib.styles import getSampleStyleSheet

# -------- ENV --------
AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ["AIRTABLE_BASE_ID"]
TRANSCRIPT_TABLE = os.environ.get("TRANSCRIPT_TABLE", "Students 1221")

SCHOOL_NAME = os.environ.get("SCHOOL_NAME", "Cornerstone Education Management Organization")
ADDR_LINE_1  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE1", "Cornerstone Online")
ADDR_LINE_2  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE2", "18550 Fajarado St.")
ADDR_LINE_3  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE3", "Rowland Heights, CA 91748")

LOGO_PATH = os.environ.get("LOGO_PATH", "logo_cornerstone.png")
SIGN_PATH = os.environ.get("SIGNATURE_PATH", "signature_principal.png")
PRINCIPAL = os.environ.get("PRINCIPAL_NAME", "Ursula Derios")
SIGN_DATEFMT = os.environ.get("SIGN_DATE_FMT", "%B %d, %Y")

RECORD_ID = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else None)
if not RECORD_ID:
    sys.exit("[ERROR] Missing RECORD_ID")

# Airtable fields
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

# -------- helpers --------
def sget(fields: Dict[str, Any], key: str, default: str = "") -> str:
    v = fields.get(key)
    if v is None: return default
    if isinstance(v, list): return ", ".join(str(x) for x in v if str(x).strip())
    return str(v)

def listify(v: Any) -> List[str]:
    if v is None: return []
    if isinstance(v, list): return [str(x).strip() for x in v if str(x).strip()]
    return [p.strip() for p in str(v).split(",") if p.strip()]

def esc(s: str) -> str: return (s or "").replace('"', '\\"')

def fetch_group(student_name: str) -> List[Dict[str, Any]]:
    formula = f'{{{F["student_name"]}}} = "{esc(student_name)}"'
    return table.all(formula=formula)

def detect_semester(name: str, code: str) -> Tuple[bool, bool]:
    t = f"{name} {code}".lower()
    is_a = ("-a" in t) or (" a " in t) or t.endswith(" a")
    is_b = ("-b" in t) or (" b " in t) or t.endswith(" b")
    return (is_a and not is_b, is_b and not is_a)

class CenterLine(Flowable):
    def __init__(self, width=220, thickness=0.7):
        super().__init__()
        self.width, self.thickness = width, thickness
        self.height = 3
    def draw(self):
        self.canv.setLineWidth(self.thickness)
        self.canv.line(-self.width/2, 0, self.width/2, 0)

# -------- PDF --------
def build_pdf(student_fields: Dict[str, Any], rows: List[Dict[str, Any]]):
    student_name = sget(student_fields, F["student_name"]).strip()
    student_id   = sget(student_fields, F["student_id"])
    grade        = sget(student_fields, F["grade"])
    year         = sget(student_fields, F["school_year"])

    out = pathlib.Path("output"); out.mkdir(parents=True, exist_ok=True)
    pdf_path = out / f"transcript_{student_name.replace(' ', '_').replace(',', '')}_{year}.pdf"

    # Page geometry
    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,             # portrait to match template
        leftMargin=36, rightMargin=36,  # 0.5" margins
        topMargin=36, bottomMargin=40
    )
    W = doc.width  # printable width

    styles = getSampleStyleSheet()
    normal = styles["Normal"]
    h2 = styles["Heading2"]; h2.alignment = 1; h2.spaceBefore, h2.spaceAfter = 2, 10

    header_gray = colors.HexColor("#C8C8C8")
    row_alt1, row_alt2 = colors.whitesmoke, colors.HexColor("#EFEFEF")

    story: List[Any] = []

    # === Single header row: StudentInfo | Logo | Address ===
    # Student Info box (left)
    info_tbl = PdfTable([
        ["Student Info", ""],
        ["Name", student_name],
        ["Current Grade Level", grade],
        ["Student ID", student_id],
    ], colWidths=[0.28*W, 0.32*W])
    info_tbl.setStyle(TableStyle([
        ("SPAN", (0,0), (-1,0)),
        ("BACKGROUND", (0,0), (-1,0), header_gray),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("BOX", (0,0), (-1,-1), 0.9, colors.black),
        ("INNERGRID", (0,0), (-1,-1), 0.4, colors.grey),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
        ("TOPPADDING", (0,1), (-1,-1), 3.5),
        ("BOTTOMPADDING", (0,1), (-1,-1), 3.5),
    ]))

    # Center logo
    center = []
    if pathlib.Path(LOGO_PATH).exists():
        center.append(Image(LOGO_PATH, width=120, height=45))

    # Right address block
    right = [
        Paragraph(f"<b>{SCHOOL_NAME}</b>", normal),
        Paragraph(ADDR_LINE_1, normal),
        Paragraph(ADDR_LINE_2, normal),
        Paragraph(ADDR_LINE_3, normal),
    ]

    header = PdfTable([[info_tbl, center, right]],
                      colWidths=[0.60*W, 0.18*W, 0.22*W])
    header.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "TOP"),
                                ("ALIGN", (1,0), (1,0), "CENTER")]))
    story.append(header)
    story.append(Spacer(1, 18))

    # === Titles centered ===
    story.append(Paragraph("Report Card", h2))
    story.append(Paragraph(f"For School Year {year}", normal))
    story.append(Spacer(1, 10))

    # === Courses table ===
    table_data = [["Course Name", "Course Number", "Teacher", "S1", "S2"]]

    expanded: List[List[str]] = []
    for r in rows:
        f = r.get("fields", {})
        names   = listify(f.get(F["course_name_rollup"]))
        codes   = listify(f.get(F["course_code_rollup"]))
        teacher = sget(f, F["teacher"])
        grade_val = sget(f, F["letter"]) or sget(f, F["percent"])
        n = max(len(names), len(codes))
        for i in range(n):
            nm = names[i] if i < len(names) else ""
            cd = codes[i] if i < len(codes) else ""
            a, b = detect_semester(nm, cd)
            s1 = grade_val or "—" if (a or not (a or b)) else ""
            s2 = grade_val or "—" if b else ""
            expanded.append([nm, cd, teacher, s1, s2])

    # dedupe + tidy + sort
    seen, clean = set(), []
    for row in expanded:
        t = tuple(row)
        if t not in seen and any(x.strip() for x in row):
            seen.add(t); clean.append(row)
    clean.sort(key=lambda x: (x[0].lower(), x[1].lower()))
    table_data.extend(clean if clean else [["(no courses found)", "", "", "", ""]])

    # Column widths as fractions of W (sum = 1.0)
    cw = [0.58*W, 0.20*W, 0.14*W, 0.04*W, 0.04*W]
    courses = PdfTable(table_data, colWidths=cw)
    courses.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), header_gray),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.35, colors.grey),
        ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [row_alt1, row_alt2]),
        ("FONTSIZE", (0,0), (-1,-1), 9.6),
        ("TOPPADDING", (0,0), (-1,-1), 4.2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4.2),
    ]))
    story.append(courses)
    story.append(Spacer(1, 40))

    # === Centered signature block at bottom ===
    sig_cells = []
    if pathlib.Path(SIGN_PATH).exists():
        sig_cells.append(Image(SIGN_PATH, width=160, height=55))
        sig_cells.append(Spacer(1, 4))
    sig_cells.append(CenterLine(width=220))
    sig_cells.append(Spacer(1, 2))
    sig_cells.append(Paragraph(f"Principal - {PRINCIPAL}", normal))
    sig_cells.append(Paragraph(f"Date: {datetime.today().strftime(SIGN_DATEFMT)}", normal))

    sig_tbl = PdfTable([[sig_cells]], colWidths=[W],
                       style=TableStyle([("ALIGN", (0,0), (-1,-1), "CENTER")]))
    story.append(sig_tbl)

    doc.build(story)
    print(f"[OK] Generated → {pdf_path}")
    return pdf_path

# -------- main --------
def main():
    rec = table.get(RECORD_ID)
    if not rec or "fields" not in rec:
        sys.exit("[ERROR] Could not fetch record or empty fields.")
    fields = rec["fields"]

    raw_name = fields.get(F["student_name"])
    student_name = raw_name[0] if isinstance(raw_name, list) and raw_name else str(raw_name or "")
    if not student_name:
        sys.exit(f"[ERROR] Field '{F['student_name']}' is empty.")

    group = fetch_group(student_name)
    if not group:
        group = [rec]

    build_pdf(fields, group)

if __name__ == "__main__":
    main()
