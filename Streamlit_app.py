import streamlit as st
import pandas as pd
import os
import re
import zipfile
from io import BytesIO
from tempfile import NamedTemporaryFile
from pypdf import PdfReader
from docx import Document as DocxDocument

# -------------------------
# CONFIGURATION
# -------------------------
ASSISTANT_NAME = "Ismail"
TOTAL_SESSIONS = 15  # Each course is 15 sessions/weeks
MIN_INPERSON_SESSIONS = 8  # 50% or more

# Map each column to possible labels/keywords (add more as needed)
COLUMN_PATTERNS = {
    "Course Name & Number": [r"GBA\s*\d{4}[A-Za-z]?[:\-. ]+.+", r"Course Name", r"Course Title"],
    "Faculty Name": [r"Instructor", r"Professor", r"Faculty", r"Dr\."],
    "Faculty CPP email included?": [r"[a-zA-Z0-9._%+-]+@cpp\.edu"],
    "Class schedule (day and time)?": [r"Class schedule", r"Meeting Day", r"Meeting Time", r"Class Meetings", r"Schedule"],
    "Class location (building number & classroom number)": [r"Location", r"Room", r"Building", r"Bldg"],
    "Offic hours?": [r"Office Hours", r"student hours"],
    "Office location?": [r"Office Location", r"Office:", r"Office Bldg"],
    "Course Learning Outcomes/Objectives included?": [r"Learning Objectives", r"Learning Outcomes", r"Objectives"],
    "Course modality specified?": [r"modality", r"instruction mode", r"hybrid", r"in-person", r"asynchronous", r"synchronous", r"online", r"format"],
    "Final Grade components explained": [r"Grading", r"Grade", r"weight", r"points"],
    "Weekly Schedule included?": [r"Week\s*\d+", r"Session\s*\d+", r"Module\s*\d+", r"Schedule", r"Calendar"],
    "Min. 50% in person class dates?": [r"In[- ]?person", r"F2F", r"Face[- ]?to[- ]?Face", r"Session"],
}

# -------------------------
# FILE READERS
# -------------------------
def extract_text_pdf(path):
    text_parts = []
    try:
        reader = PdfReader(path)
        for page in reader.pages:
            t = page.extract_text() or ""
            if t:
                text_parts.append(t)
    except Exception:
        pass
    return "\n".join(text_parts)

def extract_text_docx(path):
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
    return "\n".join(parts)

def extract_text_txt(path):
    try:
        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    except Exception:
        return ""

def extract_text_generic(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return extract_text_pdf(path)
    elif ext == ".docx":
        return extract_text_docx(path)
    elif ext in (".txt", ".md"):
        return extract_text_txt(path)
    else:
        return ""

def gather_syllabus_paths(uploaded_names):
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

def load_template_columns(template_path):
    if template_path.endswith('.ods'):
        df = pd.read_excel(template_path, engine="odf")
    else:
        df = pd.read_excel(template_path)
    return list(df.columns)

# -------------------------
# EXTRACTION LOGIC
# -------------------------
def extract_value_for_column(column, lines):
    patterns = COLUMN_PATTERNS.get(column, [])
    if column == "Course Name & Number":
        # Prefer direct course code + title extraction
        for ln in lines:
            m = re.match(r"(GBA\s*\d{4}[A-Za-z]?)[\s:.\-]+(.+)", ln)
            if m:
                return f"{m.group(1).strip()}: {m.group(2).strip()}"
        # Fallback: search for label
        for pat in patterns:
            for i, ln in enumerate(lines):
                if re.search(pat, ln, re.I):
                    if ":" in ln:
                        return ln.split(":",1)[-1].strip()
                    elif i+1 < len(lines):
                        next_ln = lines[i+1].strip()
                        if len(next_ln) > 5:
                            return next_ln
    elif column == "Faculty Name":
        for ln in lines:
            if any(re.search(pat, ln, re.I) for pat in patterns):
                # Remove titles/credentials/emails/phones
                name = ln.split(":",1)[-1].strip()
                name = re.sub(r",?\s*(Ph\.?D\.?|MBA|CPA|CGMA|Esq\.?|Ed\.?D\.?|MSc|MSBA)", "", name)
                name = re.sub(r"[a-zA-Z0-9._%+-]+@cpp\.edu", "", name)
                name = re.sub(r"\(.+?\)", "", name)
                name = name.split('(')[0].strip()
                if name: return name
                # Fallback: next line
                idx = lines.index(ln)
                if idx+1 < len(lines):
                    next_ln = lines[idx+1].strip()
                    if len(next_ln.split()) <= 4 and "@" not in next_ln:
                        return next_ln
    elif column == "Faculty CPP email included?":
        for ln in lines:
            if re.search(r"[a-zA-Z0-9._%+-]+@cpp\.edu", ln):
                return "Yes"
        return "No"
    elif column == "Course modality specified?":
        for ln in lines:
            if any(re.search(pat, ln, re.I) for pat in patterns):
                # Only report the exact phrase
                match = re.search(r"(Hybrid Asynchronous|Hybrid Synchronous|Hybrid|In[- ]?person|Asynchronous|Synchronous|Online)", ln, re.I)
                if match:
                    return match.group(1).strip()
                return ln.strip()
        return ""
    elif column == "Min. 50% in person class dates?":
        # Use dedicated function (see below)
        return ""
    else:
        for pat in patterns:
            for i, ln in enumerate(lines):
                if re.search(pat, ln, re.I):
                    if ":" in ln:
                        value = ln.split(":",1)[-1].strip()
                        if value:
                            return "Yes" if column.endswith("?") else value
                    elif i+1 < len(lines):
                        next_ln = lines[i+1].strip()
                        if next_ln and not any(re.search(p, next_ln, re.I) for p in patterns):
                            return "Yes" if column.endswith("?") else next_ln
        # Default
        return ""
    return ""

def check_weekly_schedule(lines):
    for pat in COLUMN_PATTERNS["Weekly Schedule included?"]:
        for ln in lines:
            if re.search(pat, ln, re.I):
                return "Yes"
    return "No"

def check_50pct_inperson(lines):
    inperson_count = 0
    total_sessions = 0
    schedule_lines = []
    for ln in lines:
        if re.search(r"Week\s*\d+|Session\s*\d+", ln, re.I):
            schedule_lines.append(ln)
    # Scan for explicit "In-person" or "F2F"
    for ln in schedule_lines:
        total_sessions += 1
        if re.search(r"In[- ]?person|F2F|Face[- ]?to[- ]?Face", ln, re.I):
            inperson_count += 1
    # If schedule is not explicit, look for possible mention of all sessions elsewhere
    if total_sessions == 0:
        for ln in lines:
            if re.search(r"In[- ]?person|F2F|Face[- ]?to[- ]?Face", ln, re.I):
                inperson_count += 1
                total_sessions += 1
    if total_sessions == 0:
        return "No", "schedule/class dates not available"
    if inperson_count >= MIN_INPERSON_SESSIONS:
        return "Yes", ""
    else:
        # Only put Notes if schedule is missing/ambiguous (as per your option B)
        if inperson_count == 0:
            return "No", "no in-person sessions"
        else:
            return "No", ""
    # If ambiguous, default to "No"

def analyze_one_file_prompt(path, template_cols, assistant_name=ASSISTANT_NAME):
    text = extract_text_generic(path)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    row = {c: "" for c in template_cols}
    row["Student Assistants' Name (who works on the sheet)"] = assistant_name
    notes = []
    for col in template_cols:
        if col == "Student Assistants' Name (who works on the sheet)":
            continue
        if col == "Weekly Schedule included?":
            val = check_weekly_schedule(lines)
            row[col] = val
            if val == "No":
                notes.append("Weekly schedule not explicit")
        elif col == "Min. 50% in person class dates?":
            val, note = check_50pct_inperson(lines)
            row[col] = val
            if note:
                notes.append(note)
        else:
            val = extract_value_for_column(col, lines)
            if col.endswith("?"):
                row[col] = "Yes" if val else "No"
            else:
                row[col] = val
    row["Notes"] = " | ".join(notes) if notes else ""
    return row

# -------------------------
# STREAMLIT APP UI
# -------------------------
st.set_page_config(page_title="Graduate Business School Syllabus Reviewer", layout="centered")
st.title("Graduate Business School Syllabus Reviewer")
st.markdown("""
Upload your **Excel/ODS template** and then one or more **syllabus files** (.pdf, .docx, .txt, .zip).  
The app will extract the required data and provide a single Excel file with one row per syllabus.
""")

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
                row = analyze_one_file_prompt(path, template_cols)
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
