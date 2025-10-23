import os, sys, pathlib, re
from datetime import datetime
from typing import Dict, Any, List, Tuple, Set

import requests  # for attachment fallback
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

# Partner tables (comma-separated). Fallback to TRANSCRIPT_TABLE or "Students 1221".
TRANSCRIPT_TABLES_ENV = os.environ.get("TRANSCRIPT_TABLES", "")
TRANSCRIPT_TABLE_FALLBACK = os.environ.get("TRANSCRIPT_TABLE", "Students 1221")

# Accept RECORD_IDS (comma) or single RECORD_ID (or argv[1])
RECORD_IDS_ENV   = os.getenv("RECORD_IDS", "").strip()
RECORD_ID_SINGLE = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else "")

# 0 (default): only explicit record IDs. 1: merge rows by student name across all tables.
INCLUDE_MATCH_BY_NAME = os.getenv("INCLUDE_MATCH_BY_NAME", "0").strip() == "1"

if not RECORD_IDS_ENV and not RECORD_ID_SINGLE:
    sys.exit("[ERROR] Provide at least one record id via RECORD_IDS (comma-separated) or RECORD_ID")

SCHOOL_NAME = os.environ.get("SCHOOL_NAME", "Cornerstone Education Management Organization")
ADDR_LINE_1  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE1", "Cornerstone Online")
ADDR_LINE_2  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE2", "18550 Fajarado St.")
ADDR_LINE_3  = os.environ.get("SCHOOL_HEADER_RIGHT_LINE3", "Rowland Heights, CA 91748")

LOGO_PATH      = os.environ.get("LOGO_PATH", "logo_cornerstone.png")
SIGNATURE_PATH = os.environ.get("SIGNATURE_PATH", "signature_principal.png")
PRINCIPAL      = os.environ.get("PRINCIPAL_NAME", "Ursula Derios")
SIGN_DATEFMT   = os.environ.get("SIGN_DATE_FMT", "%B %d, %Y")

# ========= LOG TABLE =========
TRANSCRIPT_LOG_TABLE = os.environ.get("TRANSCRIPT_LOG_TABLE", "TRANSCRIPT_LOG_TABLE")
LOG_FIELD_STUDENT_NAME      = os.environ.get("LOG_FIELD_STUDENT_NAME",      "Student Name")
LOG_FIELD_STUDENT_ID        = os.environ.get("LOG_FIELD_STUDENT_ID",        "Student Canvas ID")
LOG_FIELD_SCHOOL_YEAR       = os.environ.get("LOG_FIELD_SCHOOL_YEAR",       "School Year")
LOG_FIELD_GRADE_LEVEL       = os.environ.get("LOG_FIELD_GRADE_LEVEL",       "Grade Level")
LOG_FIELD_RUN_AT            = os.environ.get("LOG_FIELD_RUN_AT",            "Transcript Run At (ISO)")
LOG_FIELD_COURSES_TEXT      = os.environ.get("LOG_FIELD_COURSES_TEXT",      "Course List (text)")
LOG_FIELD_SOURCE_IDS        = os.environ.get("LOG_FIELD_SOURCE_IDS",        "Source Record IDs (csv)")
LOG_FIELD_PDF_ATTACHMENT    = os.environ.get("LOG_FIELD_PDF_ATTACHMENT",    "Transcript PDF")  # Attachment field

# ========= AIRTABLE FIELDS =========
F = {
    "student_name": "Student Name",
    "student_id": "Student Canvas ID",
    "grade_select": "Grade Select",
    "school_year": "School Year",

    # course details
    "course_name": "Course Name",
    "course_code": "Course Code",              # -> Course Number
    "assigned_teachers": "Assigned Teachers",  # -> Teacher (first)
    "letter": "Grade Letter",                  # -> Grade (Letter)

    # numeric score for Grade %
    "current_score": "# Current Score",

    # rollups (fallbacks)
    "course_name_rollup": "Course Name Rollup (from Southlands Courses Enrollment 3)",
    "course_code_rollup": "Course Code Rollup (from Southlands Courses Enrollment 3)",
}

# ========= THEME =========
GRAY_HEADER = colors.HexColor("#F1F3F5")
ROW_ALT     = colors.HexColor("#FBFCFD")
BORDER_GRAY = colors.HexColor("#D0D7DE")
INK         = colors.HexColor("#0F172A")
ACCENT      = colors.HexColor("#0C4A6E")

# ========= TWEAKS =========
TOP_GUTTER_PTS   = 200
LOGO_MAX_W_PCT   = float(os.environ.get("LOGO_MAX_W_PCT", "0.30"))
LOGO_MAX_H_PT    = int(os.environ.get("LOGO_MAX_H_PT", "72"))
LOGO_BOTTOM_SPACE= int(os.environ.get("LOGO_BOTTOM_SPACE", "15"))

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

def esc(s: str) -> str:
    return (s or "").replace('"', '\\"')

def table_names() -> List[str]:
    names = [t.strip() for t in TRANSCRIPT_TABLES_ENV.split(",") if t.strip()]
    if not names and TRANSCRIPT_TABLE_FALLBACK.strip():
        names = [TRANSCRIPT_TABLE_FALLBACK.strip()]
    if not names:
        names = ["Students 1221"]
    return names

def try_get(table_name: str, record_id: str):
    t = api.table(AIRTABLE_BASE_ID, table_name)
    return t, t.get(record_id)

def get_rec_and_table(record_id: str):
    last_err = None
    for tname in table_names():
        try:
            t, r = try_get(tname, record_id)
            print(f"[INFO] Found record {record_id} in table: {tname}")
            return t, r
        except Exception as e:
            last_err = e
            print(f"[DEBUG] Not in '{tname}': {e}")
    raise SystemExit(f"[ERROR] Record {record_id} not found in any configured tables. Last error: {last_err}")

def fetch_rows_for_name_across_all_tables(student_name: str) -> List[Dict[str, Any]]:
    merged: List[Dict[str, Any]] = []
    for tname in table_names():
        try:
            tbl = api.table(AIRTABLE_BASE_ID, tname)
            formula = f'{{{F["student_name"]}}} = "{esc(student_name)}"'
            rows = tbl.all(formula=formula)
            if rows:
                print(f"[INFO] +{len(rows)} rows from '{tname}' for {student_name}")
                merged.extend(rows)
        except Exception as e:
            print(f"[WARN] Could not query '{tname}': {e}")
    return merged

def looks_honors(name: str) -> bool:
    t = (name or "").lower()
    return ("honors" in t) or ("ap " in t) or ("ap®" in t) or ("ap " in t) or (" ap" in t)

QP_BASE = {
    "A": 4.0, "A-": 3.7,
    "B+": 3.5, "B": 3.0, "B-": 2.7,
    "C+": 2.5, "C": 2.0, "C-": 1.7,
    "D+": 1.5, "D": 1.0, "D-": 0.7,
    "F": 0.0
}

def quality_points(letter: str, honors: bool) -> str:
    l = (letter or "").strip().upper()
    if l in QP_BASE:
        qp = QP_BASE[l] + (1.0 if honors and l != "F" else 0.0)
        return f"{qp:.1f}".rstrip("0").rstrip(".")
    return ""

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
        self.dx = dx
        self.img = ImageReader(path)
        iw, ih = self.img.getSize()
        scale = min(max_w/iw, max_h/ih) if iw and ih else 1.0
        self.w, self.h = (iw*scale if iw else max_w), (ih*scale if ih else max_h)
        self.width = self.w
        self.height = self.h
    def draw(self):
        self.canv.saveState()
        self.canv.translate(self.dx, 0)
        self.canv.drawImage(self.img, 0, 0, width=self.w, height=self.h, mask='auto')
        self.canv.restoreState()

def safe_filename(raw: str) -> str:
    s = re.sub(r"\s+", "_", raw or "")
    s = s.replace(",", "_")
    s = re.sub(r"[^A-Za-z0-9._-]+", "_", s)
    return s.strip("_") or "file"

# ---- numeric helpers for Grade % ----
def _coerce_number(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s == "":
        return ""
    try:
        return float(s)
    except Exception:
        return ""

def _fmt_percent(v):
    n = _coerce_number(v)
    if n == "":
        return ""
    if 0 <= n <= 1:
        n *= 100.0
    txt = f"{n:.2f}".rstrip("0").rstrip(".")
    return txt

# ---- build rows for a single record (strict pairing, single letter grade col) ----
def build_course_rows(fields: Dict[str, Any]) -> List[List[str]]:
    names    = listify(fields.get(F["course_name"])) or listify(fields.get(F["course_name_rollup"]))
    codes    = listify(fields.get(F["course_code"])) or listify(fields.get(F["course_code_rollup"]))
    teachers = listify(fields.get(F["assigned_teachers"]))
    letter   = sget(fields, F["letter"]).strip()
    percent  = _fmt_percent(fields.get(F["current_score"]))

    if not names and not codes:
        return []

    if not names:
        names = [""] * len(codes)
    if not codes:
        codes = [""] * len(names)

    if len(names) != len(codes):
        if len(names) == 1 and len(codes) > 1:
            names = [names[0]] * len(codes)
        elif len(codes) == 1 and len(names) > 1:
            codes = [codes[0]] * len(names)
        else:
            m = min(len(names), len(codes))
            names, codes = names[:m], codes[:m]

    first_teacher = teachers[0].split(",")[0].strip() if teachers else ""
    rows: List[List[str]] = []

    for i in range(len(names)):
        nm = names[i].strip()
        cd = codes[i].strip()
        tchr = (teachers[i] if i < len(teachers) else first_teacher)
        tchr = tchr.split(",")[0].strip()

        if not (nm or cd):
            continue

        honors = looks_honors(nm)
        credits = quality_points(letter, honors)

        rows.append([
            nm,                    # Course Name
            cd,                    # Course Number
            tchr,                  # Teacher
            letter or "—",         # Grade (Letter)
            percent,               # Grade %
            credits if credits != "" else ""  # Transferred Credits
        ])

    return rows

# ---- summarize courses for logging ----
def summarize_courses(rows: List[Dict[str, Any]]) -> str:
    lines: List[str] = []
    for r in rows:
        f = r.get("fields", {})
        local = build_course_rows(f)
        for nm, cd, *_ in local:
            if not (nm or cd):
                continue
            lines.append(f"{nm} — {cd}" if cd else nm)
    return "\n".join(lines) if lines else ""

# ---- attachment helper ----
def attach_pdf_to_log_record(log_table, rec_id: str, field_name: str, pdf_path: str):
    if not field_name:
        print("[INFO] LOG_FIELD_PDF_ATTACHMENT is empty → skipping file attach.")
        return
    if not os.path.isfile(pdf_path):
        print(f"[WARN] PDF not found at {pdf_path} → skipping attachment.")
        return

    filename = os.path.basename(pdf_path)

    try:
        with open(pdf_path, "rb") as fh:
            content = fh.read()
        log_table.upload_attachment(
            rec_id,
            field=field_name,
            filename=filename,
            content=content,
            content_type="application/pdf",
        )
        print(f"[OK] Attached PDF via pyairtable.upload_attachment → field '{field_name}'.")
        return
    except Exception as e:
        print(f"[WARN] pyairtable.upload_attachment failed ({e}). Trying Web API fallback…")

    try:
        with open(pdf_path, "rb") as fh:
            files = {"file": (filename, fh, "application/pdf")}
            r = requests.post(
                f"https://api.airtable.com/v0/bases/{AIRTABLE_BASE_ID}/attachments",
                headers={"Authorization": f"Bearer {AIRTABLE_API_KEY}"},
                files=files,
                timeout=60,
            )
        r.raise_for_status()
        upload_id = r.json().get("id")
        if not upload_id:
            raise RuntimeError(f"No 'id' returned from upload-attachment endpoint. Response: {r.text}")

        log_table.update(rec_id, {field_name: [{"id": upload_id}]})
        print("[OK] Attached PDF via Web API upload-attachment + record update.")
    except Exception as e:
        print(f"[WARN] PDF attachment upload failed even with fallback: {e}")

# ---- log row ----
def log_to_airtable(pdf_path: pathlib.Path, header_fields: Dict[str, Any], rows: List[Dict[str, Any]]):
    try:
        tlog = api.table(AIRTABLE_BASE_ID, TRANSCRIPT_LOG_TABLE)
    except Exception as e:
        print(f"[WARN] Could not open log table '{TRANSCRIPT_LOG_TABLE}': {e}")
        return

    student_name = sget(header_fields, F["student_name"]).strip()
    student_id   = sget(header_fields, F["student_id"])
    grade        = sget(header_fields, F["grade_select"])
    year         = sget(header_fields, F["school_year"])
    run_at_iso   = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    courses_text = summarize_courses(rows)
    source_ids   = ",".join({ r.get("id","") for r in rows if r.get("id") })

    payload = {}
    if LOG_FIELD_STUDENT_NAME:   payload[LOG_FIELD_STUDENT_NAME]   = student_name
    if LOG_FIELD_STUDENT_ID:     payload[LOG_FIELD_STUDENT_ID]     = student_id
    if LOG_FIELD_SCHOOL_YEAR:    payload[LOG_FIELD_SCHOOL_YEAR]    = year
    if LOG_FIELD_GRADE_LEVEL:    payload[LOG_FIELD_GRADE_LEVEL]    = grade
    if LOG_FIELD_RUN_AT:         payload[LOG_FIELD_RUN_AT]         = run_at_iso
    if LOG_FIELD_COURSES_TEXT:   payload[LOG_FIELD_COURSES_TEXT]   = courses_text
    if LOG_FIELD_SOURCE_IDS:     payload[LOG_FIELD_SOURCE_IDS]     = source_ids

    try:
        rec = tlog.create(payload)
        print(f"[OK] Logged transcript to '{TRANSCRIPT_LOG_TABLE}' ({student_name}, {year})")
    except Exception as e:
        print(f"[WARN] Failed to create log record: {e}")
        return

    try:
        attach_pdf_to_log_record(tlog, rec["id"], LOG_FIELD_PDF_ATTACHMENT, str(pdf_path))
    except Exception as e:
        print(f"[WARN] Attach step failed: {e}")

# ========= PDF =========
def build_pdf(student_fields: Dict[str, Any], rows: List[Dict[str, Any]]):
    student_name = sget(student_fields, F["student_name"]).strip()
    student_id   = sget(student_fields, F["student_id"])
    grade        = sget(student_fields, F["grade_select"])
    year         = sget(student_fields, F["school_year"])

    out = pathlib.Path("output"); out.mkdir(parents=True, exist_ok=True)
    pdf_path = out / f"transcript_{safe_filename(student_name)}_{safe_filename(year)}.pdf"

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

    # Header strip
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
        ("GRID", (0,0), (-1,-1), 0.6, BORDER_GRAY),
        ("FONTSIZE", (0,0), (-1,-1), 9.5),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
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
                          colWidths=[W*0.40, TOP_GUTTER_PTS, W*0.60 - TOP_GUTTER_PTS])
    story.append(header_row)
    story.append(Spacer(1, 6))

    if pathlib.Path(LOGO_PATH).exists():
        logo = fit_image(LOGO_PATH, max_w=W*LOGO_MAX_W_PCT, max_h=LOGO_MAX_H_PT)
        story.append(PdfTable([[logo]], colWidths=[W], style=TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER")])))
        story.append(Spacer(1, LOGO_BOTTOM_SPACE))

    story.append(Paragraph("Report Card", styles["rc_h1"]))
    story.append(Paragraph(f"For School Year {year}", styles["rc_h2"]))
    story.append(Spacer(1, 8))

    # Courses (single letter column + percent + credits)
    table_data = [["Course Name", "Course Number", "Teacher", "Grade (Letter)", "Grade %", "Transferred Credits"]]
    expanded: List[List[str]] = []
    for r in rows:
        expanded.extend(build_course_rows(r.get("fields", {})))

    seen: Set[Tuple[str, str, str, str, str, str]] = set()
    clean: List[List[str]] = []
    for row in expanded:
        t = tuple(row)
        if t not in seen and any(x.strip() for x in row if isinstance(x, str)):
            seen.add(t); clean.append(row)
    clean.sort(key=lambda x: (x[0].lower(), x[1].lower()))
    if not clean:
        clean = [["(no courses found)", "", "", "", "", ""]]
    table_data.extend(clean)

    # Tighter layout so % and Credits fit nicely
    cw = [0.34*W, 0.18*W, 0.18*W, 0.12*W, 0.09*W, 0.09*W]
    courses = PdfTable(table_data, colWidths=cw, repeatRows=1)
    courses.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), ACCENT),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 10.5),
        ("FONTSIZE", (0,1), (-1,-1), 10),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("ALIGN", (3,1), (-1,-1), "CENTER"),
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

    # Signature
    sig_col_w = W * 0.38
    if pathlib.Path(SIGNATURE_PATH).exists():
        sig_img = ShiftedImage(SIGNATURE_PATH, max_w=SIG_IMG_MAX_W, max_h=SIG_IMG_MAX_H, dx=SIG_IMG_SHIFT)
        img_tbl = PdfTable([[sig_img]], colWidths=[sig_col_w],
                           style=TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER")]))
        img_row = [img_tbl]
    else:
        img_row = [Spacer(1, 0)]

    line_tbl = PdfTable([[CenterLine(width=220, thickness=0.9)]], colWidths=[sig_col_w],
                        style=TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER")]))
    normal = getSampleStyleSheet()["Normal"]
    principal_tbl = PdfTable([[Paragraph(f"Principal - {PRINCIPAL}", normal)]],
                             colWidths=[sig_col_w],
                             style=TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER")]))
    date_tbl = PdfTable([[Paragraph(f"Date: {datetime.today().strftime(SIGN_DATEFMT)}", normal)]],
                        colWidths=[sig_col_w],
                        style=TableStyle([("ALIGN",(0,0),(-1,-1),"CENTER")]))

    sig_stack = [img_row, [Spacer(1, 3)], [line_tbl], [Spacer(1, 4)], [principal_tbl], [date_tbl]]
    sig = PdfTable(sig_stack, colWidths=[sig_col_w])
    sig_row = PdfTable([["", sig]], colWidths=[W*0.62, sig_col_w],
                       style=TableStyle([("VALIGN",(0,0),(-1,-1),"BOTTOM")]))
    story.append(sig_row)

    doc.build(story, onFirstPage=draw_page_border, onLaterPages=draw_page_border)
    print(f"[OK] Generated landscape transcript → {pdf_path}")
    return pdf_path

# ========= main =========
def main():
    ids: List[str] = []
    if RECORD_IDS_ENV:
        ids.extend([x.strip() for x in RECORD_IDS_ENV.split(",") if x.strip()])
    elif RECORD_ID_SINGLE:
        ids.append(RECORD_ID_SINGLE.strip())

    name_to_header: Dict[str, Dict[str, Any]] = {}
    name_to_rows: Dict[str, List[Dict[str, Any]]] = {}

    for rid in ids:
        try:
            _, rec = get_rec_and_table(rid)
            fields = rec.get("fields", {})
            raw = fields.get(F["student_name"])
            name = raw[0] if isinstance(raw, list) and raw else str(raw or "")
            if not name:
                print(f"[WARN] '{F['student_name']}' empty for {rid}; skipping.")
                continue

            if name not in name_to_header:
                name_to_header[name] = fields

            if INCLUDE_MATCH_BY_NAME:
                rows = fetch_rows_for_name_across_all_tables(name)
            else:
                rows = [rec]

            name_to_rows.setdefault(name, []).extend(rows)

        except SystemExit as e:
            print(str(e))
        except Exception as e:
            print(f"[ERROR] Resolving {rid}: {e}")

    if not name_to_header:
        sys.exit("[ERROR] No usable records resolved from the provided IDs.")

    for student_name, header_fields in name_to_header.items():
        rows = name_to_rows.get(student_name, [])
        if not rows:
            print(f"[WARN] No rows for {student_name}; using header only.")
            rows = [{"fields": header_fields}]
        pdf_file = build_pdf(header_fields, rows)
        log_to_airtable(pdf_file, header_fields, rows)

if __name__ == "__main__":
    main()
