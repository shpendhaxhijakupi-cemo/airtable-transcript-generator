import os, sys, pathlib
from datetime import datetime
from typing import Dict, Any, List, Tuple

from pyairtable import Api
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table as PdfTable, TableStyle, Image, Flowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

# ========= ENV =========
AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ["AIRTABLE_BASE_ID"]

# NEW: list of partner tables to try (comma-separated). Falls back to TRANSCRIPT_TABLE or "Students 1221".
TRANSCRIPT_TABLES_ENV = os.environ.get("TRANSCRIPT_TABLES", "")
TRANSCRIPT_TABLE_FALLBACK = os.environ.get("TRANSCRIPT_TABLE", "Students 1221")

# Allow multiple IDs via env RECORD_IDS (comma-separated). Fallback to single RECORD_ID.
RECORD_IDS_ENV = os.getenv("RECORD_IDS", "").strip()
RECORD_ID_SINGLE = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else "")

if not RECORD_IDS_ENV and not RECORD_ID_SINGLE:
    sys.exit("[ERROR] Provide at least one record id via RECORD_IDS (comma-separated) or RECORD_ID")

SCHOOL_NAME = os.environ.get("SCHOOL_NAME", "Southlands Schools Online")
ADDR_LINE_1  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE1", "18550 Fajarado St.")
ADDR_LINE_2  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE2", "Rowland Heights, CA 91748")
ADDR_LINE_3  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE3", "")

LOGO_PATH      = os.environ.get("LOGO_PATH", "logo_cornerstone.png")
SIGNATURE_PATH = os.environ.get("SIGNATURE_PATH", "signature_principal.png")
PRINCIPAL      = os.environ.get("PRINCIPAL_NAME", "Ursula Derios")
SIGN_DATEFMT   = os.environ.get("SIGN_DATE_FMT", "%B %d, %Y")

# Airtable fields (UPDATED)
F = {
    "student_name": "Student Name",
    "student_id": "Student Canvas ID",
    "grade_select": "Grade Select",          # header grade
    "school_year": "School Year",

    # direct fields for courses
    "course_name": "Course Name",
    "course_code": "Course Code",            # <-- Course Number now uses this
    "assigned_teachers": "Assigned Teachers",# <-- Teacher now uses this (list)

    # grades
    "letter": "Grade Letter",                # S1/S2 course grade
    "percent": "% Total",

    # rollups (fallback only if missing direct fields)
    "course_name_rollup": "Course Name Rollup (from Southlands Courses Enrollment 3)",
    "course_code_rollup": "Course Code Rollup (from Southlands Courses Enrollment 3)",
}

# ========= THEME =========
GRAY_HEADER = colors.HexColor("#F1F3F5")
ROW_ALT     = colors.HexColor("#FBFCFD")
BORDER_GRAY = colors.HexColor("#D0D7DE")
INK         = colors.HexColor("#0F172A")
ACCENT      = colors.HexColor("#0C4A6E")

# ========= TWEAK KNOBS =========
TOP_GUTTER_PTS   = 200
LOGO_MAX_W_PCT   = float(os.environ.get("LOGO_MAX_W_PCT", "0.30"))
LOGO_MAX_H_PT    = int(os.environ.get("LOGO_MAX_H_PT", "72"))
LOGO_BOTTOM_SPACE= int(os.environ.get("LOGO_BOTTOM_SPACE", "15"))

SIG_LEFTPAD      = 0
SIG_IMG_SHIFT    = int(os.environ.get("SIG_IMG_SHIFT", "-90"))
SIG_IMG_MAX_W    = int(os.environ.get("SIG_IMG_MAX_W", "160"))
SIG_IMG_MAX_H    = int(os.environ.get("SIG_IMG_MAX_H", "50"))

api = Api(AIRTABLE_API_KEY)

# ========= HELPERS =========
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

def _table_names_to_try() -> List[str]:
    names = [t.strip() for t in TRANSCRIPT_TABLES_ENV.split(",") if t.strip()]
    if not names:
        names = [t.strip() for t in [TRANSCRIPT_TABLE_FALLBACK] if t.strip()]
    if not names:
        names = ["Students 1221"]
    return names

def get_table_and_record(record_id: str):
    """Try each partner table until the record is found. Return (table_obj, record)."""
    last_err = None
    for tbl_name in _table_names_to_try():
        try:
            tbl = api.table(AIRTABLE_BASE_ID, tbl_name)
            rec = tbl.get(record_id)
            print(f"[INFO] Found record in table: {tbl_name}")
            return tbl, rec
        except Exception as e:
            last_err = e
            print(f"[DEBUG] Not in '{tbl_name}': {e}")
    raise SystemExit(f"[ERROR] Record {record_id} not found in any configured tables. Last error: {last_err}")

def fetch_group(student_name: str, tbl) -> List[Dict[str, Any]]:
    formula = f'{{{F["student_name"]}}} = "{esc(student_name)}"'
    return tbl.all(formula=formula)

def detect_semester(name: str, code: str) -> Tuple[bool, bool]:
    t = f"{name} {code}".lower()
    is_a = ("-a" in t) or (" a " in t) or t.endswith(" a")
    is_b = ("-b" in t) or (" b " in t) or t.endswith(" b")
    return (is_a and not is_b, is_b and not is_a)

class CenterLine(Flowable):
    def __init__(self, width=220, thickness=0.9):
        super().__init__()
        self.width, self.thickness = width, thickness
        self.height = 3
    def draw(self):
        self.canv.setLineWidth(self.thickness)
        self.canv.line(-self.width/2, 0, self.width/2, 0)

def draw_page_border(canv: canvas.Canvas, doc):
    canv.saveState()
    canv.setStrokeColor(BORDER_GRAY)
    canv.setLineWidth(0.6)
    m = 18
    w, h = landscape(A4)
    canv.rect(m, m, w-2*m, h-2*m)
    canv.restoreState()

def fit_image(path: str, max_w: float, max_h: float) -> Image:
    img = Image(path)
    iw, ih = img.imageWidth, img.imageHeight
    if iw == 0 or ih == 0:
        return Image(path, width=max_w, height=max_h)
    scale = min(max_w/iw, max_h/ih)
    img._restrictSize(int(iw*scale), int(ih*scale))
    return img

class ShiftedImage(Flowable):
    def __init__(self, path: str, max_w: float, max_h: float, dx: int = 0):
        super().__init__()
        from reportlab.lib.utils import ImageReader
        self.path = path
        self.dx = dx
        self.img = ImageReader(path)
        iw, ih = self.img.getSize()
        if iw == 0 or ih == 0:
            iw, ih = max_w, max_h
        scale = min(max_w/iw, max_h/ih)
        self.w = iw * scale
        self.h = ih * scale
        self.width = self.w
        self.height = self.h
    def wrap(self, availW, availH):
        return (self.w, self.h)
    def draw(self):
        self.canv.saveState()
        self.canv.translate(self.dx, 0)
        self.canv.drawImage(self.img, 0, 0, width=self.w, height=self.h, mask='auto')
        self.canv.restoreState()

# ========= PDF =========
def build_pdf(student_fields: Dict[str, Any], rows: List[Dict[str, Any]]):
    student_name = sget(student_fields, F["student_name"]).strip()
    student_id   = sget(student_fields, F["student_id"])
    grade        = sget(student_fields, F["grade_select"])  # header shows Grade Select
    year         = sget(student_fields, F["school_year"])

    out = pathlib.Path("output"); out.mkdir(parents=True, exist_ok=True)
    safe_name = student_name.replace(" ", "_").replace(",", "")
    pdf_path = out / f"transcript_{safe_name}_{year}.pdf"

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=landscape(A4),
        leftMargin=28, rightMargin=28,
        topMargin=24, bottomMargin=32
    )
    W = doc.width

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle("rc_tiny",  fontName="Helvetica",      fontSize=8.5,  textColor=INK, leading=10))
    styles.add(ParagraphStyle("rc_small", fontName="Helvetica",      fontSize=9.5,  textColor=INK, leading=11))
    styles.add(ParagraphStyle("rc_body",  fontName="Helvetica",      fontSize=10.5, textColor=INK, leading=13))
    styles.add(ParagraphStyle("rc_bold",  parent=styles["rc_body"],  fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle("rc_h1",    fontName="Helvetica-Bold", fontSize=16,   textColor=INK, alignment=TA_CENTER, leading=18))
    styles.add(ParagraphStyle("rc_h2",    fontName="Helvetica-Bold", fontSize=12,   textColor=INK, alignment=TA_CENTER, leading=14))

    story: List[Any] = []

    # ======= TOP STRIP =======
    left_data = [
        [Paragraph("<b>Student Info</b>", styles["rc_bold"]), ""],
        ["Name", Paragraph(student_name, styles["rc_body"])],
        ["Current Grade Level", Paragraph(str(grade or ""), styles["rc_body"])],
        ["Student ID", Paragraph(str(student_id or ""), styles["rc_body"])],
    ]
    left_tbl = PdfTable(left_data, colWidths=[W*0.12, W*0.28])
    left_tbl.setStyle(TableStyle([
        ("SPAN", (0,0), (1,0)),
        ("BACKGROUND", (0,0), (1,0), GRAY_HEADER),
        ("FONTNAME", (0,0), (1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9.5),
        ("TEXTCOLOR", (0,0), (-1,-1), INK),
        ("ALIGN", (0,0), (-1,0), "LEFT"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.6, BORDER_GRAY),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))

    right_data = [
        [Paragraph(f"<b>{SCHOOL_NAME}</b>", styles["rc_body"])],
        [Paragraph(ADDR_LINE_1, styles["rc_small"])],
        [Paragraph(ADDR_LINE_2, styles["rc_small"])],
        [Paragraph(ADDR_LINE_3, styles["rc_small"]) if ADDR_LINE_3 else Paragraph("", styles["rc_small"])],
    ]
    right_tbl = PdfTable(right_data, colWidths=[W*0.45])
    right_tbl.setStyle(TableStyle([
        ("ALIGN", (0,0), (-1,-1), "LEFT"),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 0),
    ]))

    header_row = PdfTable([[left_tbl, "", right_tbl]],
                          colWidths=[W*0.40, 200, W*0.60 - 200])
    header_row.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING",  (1,0), (1,0), 0),
        ("RIGHTPADDING", (1,0), (1,0), 0),
    ]))
    story.append(header_row)
    story.append(Spacer(1, 6))

    if pathlib.Path(LOGO_PATH).exists():
        max_w = W * float(os.environ.get("LOGO_MAX_W_PCT", "0.30"))
        logo = fit_image(LOGO_PATH, max_w=max_w, max_h=int(os.environ.get("LOGO_MAX_H_PT", "72")))
        logo_tbl = PdfTable([[logo]], colWidths=[W])
        logo_tbl.setStyle(TableStyle([("ALIGN", (0,0), (-1,-1), "CENTER")]))
        story.append(logo_tbl)
        story.append(Spacer(1, int(os.environ.get("LOGO_BOTTOM_SPACE", "15"))))

    story.append(Paragraph("Report Card", styles["rc_h1"]))
    story.append(Paragraph(f"For School Year {year}", styles["rc_h2"]))
    story.append(Spacer(1, 8))

    # ======= Courses table =======
    table_data = [["Course Name", "Course Number", "Teacher", "S1", "S2"]]
    expanded: List[List[str]] = []

    for r in rows:
        f = r.get("fields", {})

        # Course name + code with fallback to rollups if blank
        names = listify(f.get(F["course_name"])) or listify(f.get(F["course_name_rollup"]))
        codes = listify(f.get(F["course_code"])) or listify(f.get(F["course_code_rollup"]))

        # Teachers from Assigned Teachers (list); fallback to first unique
        teachers = listify(f.get(F["assigned_teachers"]))

        # Course grade (S1/S2) uses Grade Letter only
        grade_v = sget(f, F["letter"])

        # fallback teacher
        fallback_teacher = ""
        if teachers:
            uniq = list(dict.fromkeys([t for t in teachers if t.strip()]))
            fallback_teacher = (uniq[0] if uniq else "").split(",")[0].strip()

        n = max(len(names), len(codes), len(teachers) if teachers else 0)
        for i in range(n):
            nm = names[i] if i < len(names) else ""
            cd = codes[i] if i < len(codes) else ""
            tchr = (teachers[i] if (teachers and i < len(teachers)) else fallback_teacher).split(",")[0].strip()

            a, b = detect_semester(nm, cd)
            s1 = (grade_v or "—") if (a or not (a or b)) else ""
            s2 = (grade_v or "—") if b else ""
            expanded.append([nm, cd, tchr, s1, s2])

    # tidy/sort
    seen, clean = set(), []
    for row in expanded:
        t = tuple(row)
        if t not in seen and any(x.strip() for x in row):
            seen.add(t); clean.append(row)
    clean.sort(key=lambda x: (x[0].lower(), x[1].lower()))
    if not clean:
        clean = [["(no courses found)", "", "", "", ""]]
    table_data.extend(clean)

    Wtbl = doc.width
    cw = [0.46*Wtbl, 0.18*Wtbl, 0.22*Wtbl, 0.07*Wtbl, 0.07*Wtbl]
    courses = PdfTable(table_data, colWidths=cw, repeatRows=1)
    courses.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), ACCENT),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 10.5),
        ("FONTSIZE", (0,1), (-1,-1), 10),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("ALIGN", (3,1), (4,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.6, BORDER_GRAY),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, ROW_ALT]),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(courses)
    story.append(Spacer(1, 10))

    # ======= Signature block (right-aligned) =======
    sig_col_w = W * 0.38

    if pathlib.Path(SIGNATURE_PATH).exists():
        sig_img = ShiftedImage(SIGNATURE_PATH, max_w=SIG_IMG_MAX_W, max_h=SIG_IMG_MAX_H, dx=SIG_IMG_SHIFT)
        img_tbl = PdfTable([[sig_img]], colWidths=[sig_col_w])
        img_tbl.setStyle(TableStyle([
            ("ALIGN", (0,0), (-1,-1), "CENTER"),
            ("LEFTPADDING",  (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
            ("TOPPADDING",   (0,0), (-1,-1), 0),
            ("BOTTOMPADDING",(0,0), (-1,-1), 0),
        ]))
        img_row = [img_tbl]
    else:
        img_row = [Spacer(1, 0)]

    line_tbl = PdfTable([[CenterLine(width=220, thickness=0.9)]], colWidths=[sig_col_w])
    line_tbl.setStyle(TableStyle([("ALIGN", (0,0), (-1,-1), "CENTER")]))
    principal_tbl = PdfTable([[Paragraph(f"Principal - {PRINCIPAL}", getSampleStyleSheet()["Normal"])]], colWidths=[sig_col_w])
    principal_tbl.setStyle(TableStyle([("ALIGN", (0,0), (-1,-1), "CENTER")]))
    date_tbl = PdfTable([[Paragraph(f"Date: {datetime.today().strftime(SIGN_DATEFMT)}", getSampleStyleSheet()["Normal"])]], colWidths=[sig_col_w])
    date_tbl.setStyle(TableStyle([("ALIGN", (0,0), (-1,-1), "CENTER")]))

    sig_stack = [img_row, [Spacer(1, 3)], [line_tbl], [Spacer(1, 4)], [principal_tbl], [date_tbl]]
    sig = PdfTable(sig_stack, colWidths=[sig_col_w])
    sig.setStyle(TableStyle([
        ("LEFTPADDING",  (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING",   (0,0), (-1,-1), 0),
        ("BOTTOMPADDING",(0,0), (-1,-1), 0),
    ]))

    sig_row = PdfTable([["", sig]], colWidths=[W*0.62, sig_col_w])
    sig_row.setStyle(TableStyle([("VALIGN", (0,0), (-1,-1), "BOTTOM")]))
    story.append(sig_row)

    doc.build(story, onFirstPage=draw_page_border, onLaterPages=draw_page_border)
    print(f"[OK] Generated landscape transcript → {pdf_path}")
    return pdf_path

# ========= main =========
def main():
    # Build list of record IDs
    ids: List[str] = []
    if RECORD_IDS_ENV:
        ids.extend([x.strip() for x in RECORD_IDS_ENV.split(",") if x.strip()])
    elif RECORD_ID_SINGLE:
        ids.append(RECORD_ID_SINGLE.strip())

    pathlib.Path("output").mkdir(parents=True, exist_ok=True)

    for rid in ids:
        try:
            table_obj, rec = get_table_and_record(rid)
            if not rec or "fields" not in rec:
                print(f"[WARN] Empty fields for {rid}; skipping.")
                continue
            fields = rec["fields"]

            raw_name = fields.get(F["student_name"])
            student_name = raw_name[0] if isinstance(raw_name, list) and raw_name else str(raw_name or "")
            if not student_name:
                print(f"[WARN] Field '{F['student_name']}' empty for {rid}; skipping.")
                continue

            group = fetch_group(student_name, table_obj) or [rec]
            build_pdf(fields, group)
        except SystemExit as e:
            print(str(e))
        except Exception as e:
            print(f"[ERROR] Failed for {rid}: {e}")

if __name__ == "__main__":
    main()
