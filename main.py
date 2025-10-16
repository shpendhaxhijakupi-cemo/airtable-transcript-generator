import os
import sys
import pathlib
from typing import List, Dict, Any, Tuple, Optional

from pyairtable import Table, Api
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table as PdfTable, TableStyle

# ---------------- Env ----------------
AIRTABLE_API_KEY = os.environ["AIRTABLE_API_KEY"]
AIRTABLE_BASE_ID = os.environ["AIRTABLE_BASE_ID"]
STUDENT_TABLE = os.environ.get("STUDENT_TABLE", "Students")
ENROLLMENTS_TABLE = os.environ.get("ENROLLMENTS_TABLE", "Enrollments")
SCHOOL_NAME = os.environ.get("SCHOOL_NAME", "School")

RECORD_ID = os.getenv("RECORD_ID") or (sys.argv[1] if len(sys.argv) > 1 else None)
if not RECORD_ID:
    print("Missing RECORD_ID")
    sys.exit(1)

api = Api(AIRTABLE_API_KEY)
students = api.table(AIRTABLE_BASE_ID, STUDENT_TABLE)
enrollments = api.table(AIRTABLE_BASE_ID, ENROLLMENTS_TABLE)

# --------------- Helpers ------------
def safe_get(d: Dict[str, Any], key: str, default=None):
    v = d.get(key, default)
    return v if v is not None else default

def fetch_student(record_id: str) -> Dict[str, Any]:
    return students.get(record_id)

def fetch_enrollments_for_student(student_record: Dict[str, Any]) -> List[Dict[str, Any]]:
    fields = student_record.get("fields", {})
    # Common pattern: student has a linked-record field named "Enrollments"
    linked_ids = fields.get("Enrollments", []) or fields.get("Enrollment", [])
    results = []
    if linked_ids:
        # Batch fetch linked enrollment rows
        for rec_id in linked_ids:
            try:
                results.append(enrollments.get(rec_id))
            except Exception as e:
                print(f"[WARN] Could not fetch enrollment {rec_id}: {e}")
        return results

    # Fallback pattern: student table contains arrays of Courses/Grades directly
    # Try to normalize them into "pseudo-enrollment" rows
    courses = fields.get("Courses", [])
    grades = fields.get("Grades", [])
    credits = fields.get("Credits", [])
    results = []
    n = max(len(courses), len(grades), len(credits))
    for i in range(n):
        results.append({
            "id": f"pseudo-{i}",
            "fields": {
                "Course": courses[i] if i < len(courses) else "",
                "Grade": grades[i] if i < len(grades) else "",
                "Credits": credits[i] if i < len(credits) else ""
            }
        })
    return results

def extract_student_name(fields: Dict[str, Any]) -> Tuple[str, str, str]:
    first = str(fields.get("Student First Name") or fields.get("First Name") or fields.get("First") or "").strip()
    last = str(fields.get("Student Last Name") or fields.get("Last Name") or fields.get("Last") or "").strip()
    full = str(fields.get("Name") or f"{first} {last}".strip()).strip()
    if not first or not last:
        # Try to split Name if needed
        parts = full.split()
        if len(parts) >= 2:
            first = first or parts[0]
            last = last or " ".join(parts[1:])
    return full or f"{first} {last}".strip(), first, last

def build_rows_from_enrollments(enrollment_records: List[Dict[str, Any]]) -> List[Tuple[str, str, str]]:
    rows = []
    for r in enrollment_records:
        f = r.get("fields", {})
        # Common field names in an Enrollments table:
        # Course (text or lookup), Final Grade / Grade, Term, Credits
        course = str(f.get("Course Name") or f.get("Course") or f.get("Course Code") or "").strip()
        grade = str(f.get("Final Grade") or f.get("Grade") or "").strip()
        credits = str(f.get("Credits") or f.get("Credit") or "").strip()
        rows.append((course, grade, credits))
    return rows

def parse_float_or_zero(val: Any) -> float:
    try:
        return float(val)
    except Exception:
        return 0.0

def compute_gpa(enrollment_records: List[Dict[str, Any]]) -> Optional[float]:
    # If a GPA exists in student record we’ll use it; otherwise try to compute simple average from a numeric "GPA Points" field.
    # This is a placeholder; we’ll improve later if needed.
    pts = []
    for r in enrollment_records:
        f = r.get("fields", {})
        p = f.get("GPA Points")
        if p is not None:
            pts.append(parse_float_or_zero(p))
    if pts:
        return round(sum(pts) / len(pts), 3)
    return None

# --------------- Main ---------------
def main():
    student = fetch_student(RECORD_ID)
    s_fields = student.get("fields", {})
    full_name, first_name, last_name = extract_student_name(s_fields)
    grade_level = s_fields.get("Grade Level") or s_fields.get("Grade") or ""
    student_id = s_fields.get("Student ID") or s_fields.get("SID") or ""

    enrollment_records = fetch_enrollments_for_student(student)
    course_rows = build_rows_from_enrollments(enrollment_records)
    gpa = s_fields.get("GPA") or compute_gpa(enrollment_records)

    # ---- produce PDF ----
    out_dir = pathlib.Path("output")
    out_dir.mkdir(parents=True, exist_ok=True)
    safe_last = last_name.replace(" ", "_") if last_name else "student"
    pdf_path = out_dir / f"transcript_{safe_last}_{student['id']}.pdf"

    styles = getSampleStyleSheet()
    story = []

    # Header
    story.append(Paragraph(f"<b>{SCHOOL_NAME}</b>", styles["Title"]))
    story.append(Paragraph("Official Transcript", styles["h2"]))
    story.append(Spacer(1, 6 * mm))

    # Student block
    student_lines = [
        f"<b>Name:</b> {full_name}",
        f"<b>Student ID:</b> {student_id}" if student_id else "",
        f"<b>Grade Level:</b> {grade_level}" if grade_level else "",
        f"<b>GPA:</b> {gpa}" if gpa is not None else ""
    ]
    student_lines = [x for x in student_lines if x]
    for line in student_lines:
        story.append(Paragraph(line, styles["Normal"]))
    story.append(Spacer(1, 6 * mm))

    # Courses table
    table_data = [["Course", "Grade", "Credits"]]
    if course_rows:
        table_data.extend(course_rows)
    else:
        table_data.append(["(no courses found)", "", ""])

    t = PdfTable(table_data, colWidths=[110*mm, 30*mm, 30*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("TEXTCOLOR", (0,0), (-1,0), colors.black),
        ("ALIGN", (1,1), (-1,-1), "CENTER"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("BOTTOMPADDING", (0,0), (-1,0), 8),
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
    ]))
    story.append(t)

    story.append(Spacer(1, 10 * mm))
    story.append(Paragraph(
        "This transcript is generated from the Student Information System and reflects the official record as of the run date.",
        styles["Italic"]
    ))

    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm)
    doc.build(story)

    print(f"[OK] PDF written → {pdf_path}")

if __name__ == "__main__":
    main()
