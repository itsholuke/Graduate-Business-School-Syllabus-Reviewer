import streamlit as st
import pandas as pd
import os
import re
import zipfile
from io import BytesIO, StringIO
from tempfile import NamedTemporaryFile

# =====================
# 1. Helper functions
# =====================

# PDF/DOCX/TXT extractors
from pypdf import PdfReader
from docx import Document as DocxDocument

DAY_NAMES = r"(?:Mon(?:day)?|Tue(?:sday)?|Wed(?:nesday)?|Thu(?:rsday)?|Fri(?:day)?|Sat(?:urday)?|Sun(?:day)?)"
TIME_RE = r"(?:\b\d{1,2}:\d{2}\s?(?:AM|PM|am|pm)\b|\b\d{1,2}\s?(?:AM|PM|am|pm)\b)"
EMAIL_CPP_RE = re.compile(r"\b[A-Za-z0-9._%+-]+@(?:cpp|csupomona)\.edu\b", re.I)
INPERSON_TOKENS = [
    r"\bF2F\b", r"face[- ]?to[- ]?face", r"\bin[- ]?person\b", r"on[- ]?campus", r"classroom", r"room\s*[A-Za-z0-9-]+",
]
SCHEDULE_SECTION_HINTS = re.compile(
    r"(weekly\s*schedule|course\s*schedule|tentative\s*schedule|schedule\s*of\s*topics|class\s*schedule|calendar)",
    re.I,
)
MODALITY_HINTS = re.compile(
    r"(modality|instruction\s*mode|mode\s*of\s*instruction|delivery|format|meeting\s*modality|class\s*modality)",
    re.I,
)

class TextBlob:
    def __init__(self, text: str, source_path: str):
        self.text = text or ""
        self.source = source_path
        self.lines = [ln.strip() for ln in self.text.splitlines() if ln.strip()]

    def contains(self, pattern: str) -> bool:
        return re.search(pattern, self.text, re.I) is not None

    def find_lines(self, pattern: str, window: int = 0):
        out = []
        rx = re.compile(pattern, re.I)
        for i, ln in enumerate(self.lines):
            if rx.search(ln):
                start = max(0, i - window)
                end = min(len(self.lines), i + window + 1)
                out.append(" \n".join(self.lines[start:end]))
        return out

def clean_text(s: str) -> str:
    if not s:
        return ""
    s = re.sub(r"\u00A0", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def extract_text_pdf(path: str) -> str:
    text_parts = []
    try:
        reader = PdfReader(path)
        for page in reader.pages:
            try:
                t = page.extract_text() or ""
                if t:
                    text_parts.append(t)
            except Exception:
                pass
    except Exception:
        pass
    text = "\n".join(text_parts)
    return clean_text(text)

def extract_text_docx(path: str) -> str:
    try:
        doc = DocxDocument(path)
    except Exception:
        return ""
    parts = []
    for p in doc.paragraphs:
        if p.text:
            parts.append(p.text)
    for tbl in getattr(doc, 'tables', []):
        for row in tbl.rows:
            parts.append(" \t ".join([cell.text for cell in row.cells]))
    return clean_text("\n".join(parts))

def extract_text_txt(path: str) -> str:
    try:
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            return clean_text(f.read())
    except Exception:
        return ""

def extract_text_generic(path: str) -> str:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return extract_text_pdf(path)
    elif ext == ".docx":
        return extract_text_docx(path)
    elif ext in (".txt", ".md"):
        return extract_text_txt(path)
    else:
        return ""

def detect_course_name_number(tb: TextBlob) -> str:
    top = tb.lines[:60]
    joined_top = " \n".join(top)
    m = re.search(r"\b([A-Z]{2,4}\s*\d{3,4}[A-Za-z]?)\b", joined_top)
    if m:
        idx = 0
        for i, ln in enumerate(top):
            if m.group(1) in ln:
                idx = i
                break
        title = top[idx]
        if len(title) < 20 and idx + 1 < len(top):
            title += " - " + top[idx + 1]
        return clean_text(title)
    return ""

def detect_faculty_name(tb: TextBlob) -> str:
    for ln in tb.lines[:120]:
        m = re.search(r"^(Instructor|Professor|Faculty|Lecturer)\s*[:\-]\s*(.+)$", ln, re.I)
        if m:
            name = re.sub(EMAIL_CPP_RE, "", m.group(2))
            name = re.sub(r"\(.+?\)", "", name)
            parts = [p for p in name.split(',') if '@' not in p]
            if parts:
                return clean_text(parts[0])
            return clean_text(name)
    return ""

def detect_has_cpp_email(tb: TextBlob) -> bool:
    return EMAIL_CPP_RE.search(tb.text) is not None

def detect_class_schedule(tb: TextBlob) -> bool:
    has_day = re.search(DAY_NAMES, tb.text, re.I) is not None
    has_time = re.search(TIME_RE, tb.text) is not None
    return bool(has_day and has_time)

def detect_class_location(tb: TextBlob) -> bool:
    patterns = [r"\bBldg\.?\s*\w+", r"\bBuilding\b\s*\w+", r"\bRoom\b\s*\w+", r"\b\w+-?\d{2,4}\b"]
    return any(re.search(p, tb.text, re.I) for p in patterns)

def detect_office_hours(tb: TextBlob) -> bool:
    return tb.contains(r"(office\s*hours|student\s*hours)")

def detect_office_location(tb: TextBlob) -> bool:
    return tb.contains(r"\boffice\s*(location|:)\b") or tb.contains(r"\bOffice\b\s*:\s*\S+")

def detect_learning_outcomes(tb: TextBlob) -> bool:
    return tb.contains(r"(learning\s*outcomes?|course\s*objectives?)")

def detect_final_grade_components(tb: TextBlob) -> bool:
    has_grade_word = tb.contains(r"(grading|grade\s*breakdown|assessment|evaluation)")
    has_percent = re.search(r"\d+\s*%", tb.text) is not None
    return bool(has_grade_word and has_percent)

def detect_weekly_schedule_section(tb: TextBlob):
    for i, ln in enumerate(tb.lines):
        if SCHEDULE_SECTION_HINTS.search(ln):
            block = tb.lines[i:i+220]
            return True, "\n".join(block)
    return False, ""

def detect_modality_phrase(tb: TextBlob) -> str:
    lines = tb.lines
    for i, ln in enumerate(lines):
        if MODALITY_HINTS.search(ln):
            snippet = ln
            if len(snippet) < 40 and i + 1 < len(lines):
                snippet = snippet + " | " + lines[i + 1]
            return clean_text(snippet[:140])
    explicit = re.search(r"\b(in[- ]?person|hybrid|hyflex|online\s*(?:sync|synchronous|async|asynchronous)?)\b.*", tb.text, re.I)
    if explicit:
        return clean_text(explicit.group(0)[:140])
    return ""

def detect_50pct_inperson(tb: TextBlob) -> bool:
    has_sched, sched_text = detect_weekly_schedule_section(tb)
    if not has_sched:
        return False
    lines = [ln.strip() for ln in sched_text.splitlines() if ln.strip()]
    session_lines = []
    for ln in lines:
        if re.search(r"\b(Week|Session)\s*\d+\b", ln, re.I):
            session_lines.append(ln)
    if not session_lines:
        for ln in lines:
            if re.search(r"^(?:\d{1,2}\.|\d{1,2}\))\s+", ln):
                session_lines.append(ln)
    total = len(session_lines)
    if total == 0:
        return False
    def has_any(patterns, s):
        return any(re.search(p, s, re.I) for p in patterns)
    inperson_count = 0
    for ln in session_lines:
        if has_any(INPERSON_TOKENS, ln):
            inperson_count += 1
    if inperson_count == 0:
        return False
    return (inperson_count / max(1, total)) >= 0.5

def load_template_columns(template_path: str) -> list:
    try:
        if template_path.endswith('.ods'):
            df = pd.read_excel(template_path, engine="odf")
        else:
            df = pd.read_excel(template_path)
        cols = list(df.columns)
        return cols
    except Exception as e:
        raise RuntimeError(f"Failed to read template '{template_path}': {e}")

def analyze_one_file(path: str, template_cols: list, assistant_name="Ismail") -> dict:
    text = extract_text_generic(path)
    tb = TextBlob(text, path)
    row = {c: "" for c in template_cols}
    if "Student Assistants' Name (who works on the sheet)" in row:
        row["Student Assistants' Name (who works on the sheet)"] = assistant_name
    if "Course Name & Number" in row:
        row["Course Name & Number"] = detect_course_name_number(tb)
    if "Faculty Name" in row:
        row["Faculty Name"] = detect_faculty_name(tb)
    if "Faculty CPP email included?" in row:
        row["Faculty CPP email included?"] = "Yes" if detect_has_cpp_email(tb) else "No"
    if "Class schedule (day and time)?" in row:
        row["Class schedule (day and time)?"] = "Yes" if detect_class_schedule(tb) else "No"
    if "Class location (building number & classroom number)" in row:
        row["Class location (building number & classroom number)"] = "Yes" if detect_class_location(tb) else "No"
    if "Offic hours?" in row:
        row["Offic hours?"] = "Yes" if detect_office_hours(tb) else "No"
    if "Office location?" in row:
        row["Office location?"] = "Yes" if detect_office_location(tb) else "No"
    if "Course Learning Outcomes/Objectives included?" in row:
        row["Course Learning Outcomes/Objectives included?"] = "Yes" if detect_learning_outcomes(tb) else "No"
    if "Course modality specified?" in row:
        modality_phrase = detect_modality_phrase(tb)
        row["Course modality specified?"] = modality_phrase if modality_phrase else ""
    if "Final Grade components explained" in row:
        row["Final Grade components explained"] = "Yes" if detect_final_grade_components(tb) else "No"
    has_sched, _ = detect_weekly_schedule_section(tb)
    if "Weekly Schedule included?" in row:
        row["Weekly Schedule included?"] = "Yes" if has_sched else "No"
    if "Min. 50% in person class dates?" in row:
        row["Min. 50% in person class dates?"] = "Yes" if detect_50pct_inperson(tb) else "No"
    notes = []
    if len(tb.text) < 200:
        notes.append("Very little extractable text (possibly a scanned PDF).")
    if (row.get("Offic hours?", "No") == "Yes") and (row.get("Office location?", "No") == "No"):
        notes.append("Office hours found but no office location.")
    if row.get("Course modality specified?", "").strip() == "":
        notes.append("Modality not explicitly labeled.")
    if "Notes" in row:
        row["Notes"] = " | ".join(notes) if notes else ""
    return row

def gather_syllabus_paths(uploaded_names: list) -> list:
    paths = []
    for name in uploaded_names:
        if name.lower().endswith('.zip'):
            out_dir = os.path.join("unzipped", os.path.splitext(os.path.basename(name))[0])
            os.makedirs(out_dir, exist_ok=True)
            with zipfile.ZipFile(name) as z:
                z.extractall(out_dir)
            for root, _, files in os.walk(out_dir):
                for f in files:
                    if f.lower().endswith((".pdf", ".docx", ".txt")):
                        paths.append(os.path.join(root, f))
        else:
            if name.lower().endswith((".pdf", ".docx", ".txt")):
                paths.append(name)
    return sorted(paths)

def save_uploaded_files(uploaded_files):
    saved = []
    up_dir = "uploaded_files"
    os.makedirs(up_dir, exist_ok=True)
    for f in uploaded_files:
        out_path = os.path.join(up_dir, f.name)
        with open(out_path, "wb") as out_f:
            out_f.write(f.getbuffer())
        saved.append(out_path)
    return saved

# =====================
# 2. Streamlit UI
# =====================

st.set_page_config(page_title="Graduate Business Syllabus Reviewer", layout="centered")
st.title("Graduate Business Syllabus Reviewer")
st.markdown("""
Upload your **Excel/ODS template** and then upload one or more **syllabus files** (.pdf, .docx, .txt, or .zip).
The app will extract the required data and provide a single Excel file with one row per syllabus.
""")

# Step 1: Upload Template
template_file = st.file_uploader("Step 1: Upload your Excel/ODS template", type=['ods', 'xlsx'])
if not template_file:
    st.warning("Please upload your Excel/ODS template to continue.")
    st.stop()
else:
    with NamedTemporaryFile(delete=False, suffix="."+template_file.name.split(".")[-1]) as tmp:
        tmp.write(template_file.read())
        tmp_path = tmp.name
    try:
        template_cols = load_template_columns(tmp_path)
        st.success(f"Loaded template with {len(template_cols)} columns.")
    except Exception as e:
        st.error(f"Could not load template: {e}")
        st.stop()

# Step 2: Upload syllabi files
uploaded_files = st.file_uploader("Step 2: Upload syllabus files (.pdf, .docx, .txt, or .zip)", 
                                  type=['pdf', 'docx', 'txt', 'zip'], accept_multiple_files=True)

if uploaded_files:
    with st.spinner("Processing uploaded files..."):
        saved_paths = save_uploaded_files(uploaded_files)
        syllabus_paths = gather_syllabus_paths(saved_paths)
        st.info(f"Found {len(syllabus_paths)} syllabus files to process.")

        rows = []
        for path in syllabus_paths:
            st.write(f"Analyzing: `{os.path.basename(path)}`")
            try:
                row = analyze_one_file(path, template_cols)
                rows.append(row)
            except Exception as e:
                blank = {c: "" for c in template_cols}
                blank["Notes"] = f"Error: {e}"
                rows.append(blank)
        df_out = pd.DataFrame(rows, columns=template_cols)
        st.success(f"Done! Processed {len(df_out)} syllabi.")

        towrite = BytesIO()
        df_out.to_excel(towrite, index=False)
        towrite.seek(0)
        st.download_button(
            "Download Excel Output",
            data=towrite,
            file_name="syllabus_review_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.write("Preview:")
        st.dataframe(df_out)
else:
    st.info("Upload one or more syllabus files to begin.")

