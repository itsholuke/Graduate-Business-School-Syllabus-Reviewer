import streamlit as st
import pandas as pd
import os
import re
import zipfile
from io import BytesIO
from tempfile import NamedTemporaryFile
from pypdf import PdfReader
from docx import Document as DocxDocument

st.set_page_config(page_title="Graduate Business School Syllabus Reviewer", layout="centered")
st.title("Graduate Business School Syllabus Reviewer")
st.markdown("""
Upload your **Excel/ODS template** and then one or more **syllabus files** (.pdf, .docx, .txt, .zip).  
The app will extract the required data and provide a single Excel file with one row per syllabus.
""")

# ========== Field Extraction Logic (Updated for Your Files) ==========

def extract_course_name_number(lines):
    for i, ln in enumerate(lines[:40]):
        m = re.match(r"(GBA\s*\d{4}[A-Za-z]?)[\s:.\-]+(.+)", ln)
        if m:
            return f"{m.group(1).strip()}: {m.group(2).strip()}"
        if ln.strip().startswith("GBA ") and len(ln.strip()) < 40:
            if i+1 < len(lines):
                return f"{ln.strip()}: {lines[i+1].strip()}"
    for i, ln in enumerate(lines[:20]):
        if ln.strip().isupper() and "GBA" in ln and len(ln.strip()) > 10:
            return ln.strip()
    return ""

def extract_faculty_name(lines):
    for ln in lines[:60]:
        if 'Instructor:' in ln or 'Professor:' in ln or 'Faculty:' in ln:
            name = ln.split(':',1)[-1].strip()
            name = re.sub(r",?\s*(Ph\.?D\.?|MBA|CPA|CGMA|Esq\.?|Ed\.?D\.?|MSc|MSBA|MS)", "", name)
            return name.split('(')[0].strip()
        if ln.strip().startswith("Dr. "):
            return ln.strip().split(",")[0].replace("Dr. ", "").strip()
    return ""

def extract_email(lines):
    for ln in lines[:40]:
        m = re.search(r"([a-zA-Z0-9._%+-]+@cpp\.edu)", ln)
        if m:
            return "Yes"
    return "No"

def extract_schedule(lines):
    for ln in lines[:80]:
        if "Class schedule" in ln.lower() or "Meeting Time" in ln or "Class Schedule" in ln or "Meeting Days" in ln:
            return "Yes"
        if re.search(r"\b(Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*\.?\s+\d{1,2}:", ln):
            return "Yes"
    return "No"

def extract_class_location(lines):
    for ln in lines[:80]:
        if "Location:" in ln or "Room" in ln or "Building" in ln:
            return "Yes"
    return "No"

def extract_office_hours(lines):
    for ln in lines[:80]:
        if "Office Hours" in ln or "office hours" in ln:
            return "Yes"
    return "No"

def extract_office_location(lines):
    for ln in lines[:80]:
        if ("Office Location:" in ln or "Office:" in ln) and "Office Hours" not in ln:
            return "Yes"
    return "No"

def extract_learning_outcomes(lines):
    for ln in lines[:120]:
        if "Learning Objectives" in ln or "Learning Outcomes" in ln or "Objectives" in ln:
            return "Yes"
    return "No"

def extract_modality(lines):
    for ln in lines[:80]:
        if re.search(r"hybrid|asynchronous|synchronous|in-person|face[- ]?to[- ]?face|online", ln, re.I):
            return ln.strip()
        if "Format" in ln or "Instruction Mode" in ln or "Mode" in ln:
            return ln.strip()
    return ""

def extract_grade_components(lines):
    for ln in lines:
        if "Grading" in ln or "Grade" in ln or "weight" in ln:
            if "%" in ln or "points" in ln.lower():
                return "Yes"
    return "No"

def extract_weekly_schedule(lines):
    block = "\n".join(lines)
    if re.search(r"Week\s*\d+|Module\s*\d+|Session\s*\d+|Date\s+", block, re.I):
        return "Yes"
    return "No"

def extract_50pct_in_person(lines):
    inperson_count = 0
    total_count = 0
    for ln in lines:
        if re.search(r"\b(Week|Module|Session)\b", ln) or re.search(r"\bIn[- ]?Person\b|\bF2F\b", ln, re.I):
            total_count += 1
            if re.search(r"In[- ]?Person|F2F|Face[- ]?to[- ]?Face", ln, re.I):
                inperson_count += 1
    if total_count >= 8 and inperson_count / total_count >= 0.5:
        return "Yes"
    return "No"

# ========== File Extraction Helpers ==========

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

# ========== Main Extraction Routine ==========

def analyze_one_file_v2(path, template_cols, assistant_name="Ismail"):
    text = extract_text_generic(path)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    row = {c: "" for c in template_cols}
    row["Student Assistants' Name (who works on the sheet)"] = assistant_name
    row["Course Name & Number"] = extract_course_name_number(lines)
    row["Faculty Name"] = extract_faculty_name(lines)
    row["Faculty CPP email included?"] = extract_email(lines)
    row["Class schedule (day and time)?"] = extract_schedule(lines)
    row["Class location (building number & classroom number)"] = extract_class_location(lines)
    row["Offic hours?"] = extract_office_hours(lines)
    row["Office location?"] = extract_office_location(lines)
    row["Course Learning Outcomes/Objectives included?"] = extract_learning_outcomes(lines)
    row["Course modality specified?"] = extract_modality(lines)
    row["Final Grade components explained"] = extract_grade_components(lines)
    row["Weekly Schedule included?"] = extract_weekly_schedule(lines)
    row["Min. 50% in person class dates?"] = extract_50pct_in_person(lines)
    row["Notes"] = ""
    return row

# ========== Streamlit App UI ==========

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
                row = analyze_one_file_v2(path, template_cols)
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
