import streamlit as st
import pandas as pd
import os
import re
import zipfile
from io import BytesIO
from tempfile import NamedTemporaryFile
from pypdf import PdfReader
from docx import Document as DocxDocument

ASSISTANT_NAME = "Ismail"
TOTAL_SESSIONS = 15
MIN_INPERSON_SESSIONS = 8

COLUMN_PATTERNS = {
    "Course Name & Number": [r"GBA\s*\d{4}[A-Za-z]?"],
    "Faculty Name": [r"Instructor", r"Professor", r"Faculty", r"Lecturer", r"Dr\."],
    "Faculty CPP email included?": [r"[a-zA-Z0-9._%+-]+@cpp\.edu"],
    "Class schedule (day and time)?": [r"Class schedule", r"Meeting Day", r"Meeting Time", r"Class Meetings", r"Schedule"],
    "Class location (building number & classroom number)": [r"Location", r"Room", r"Building", r"Bldg"],
    "Offic hours?": [r"Office Hours", r"student hours"],
    "Office location?": [r"Office Location", r"Office:", r"Office Bldg"],
    "Course Learning Outcomes/Objectives included?": [r"Learning Objectives", r"Learning Outcomes", r"Objectives"],
    "Course modality specified?": [r"modality", r"instruction mode", r"hybrid", r"in-person", r"asynchronous", r"synchronous", r"online", r"format"],
    "Final Grade components explained": [r"Grading", r"Grade", r"weight", r"points", r"Evaluation"],
    "Weekly Schedule included?": [r"Week\s*\d+", r"Session\s*\d+", r"Module\s*\d+", r"Schedule", r"Calendar"],
    "Min. 50% in person class dates?": [r"In[- ]?person", r"F2F", r"Face[- ]?to[- ]?Face", r"Session"],
}

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

# --- FIXED COURSE NAME AND FACULTY NAME EXTRACTION ---

def extract_course_name_number(lines):
    for i, ln in enumerate(lines):
        m = re.match(r"(GBA\s*\d{4}[A-Za-z]?)\s*[:\-–]?\s*(.+)?", ln)
        if m:
            code = m.group(1).replace(" ", "")
            rest = m.group(2).strip() if m.group(2) else ""
            bads = ["syllabus", "fall", "section", "spring", "schedule", "course number", "class", "p01", "p02"]
            if not rest or any(bad in rest.lower() for bad in bads) or len(rest.split()) < 3:
                for j in range(1, 4):
                    if i+j < len(lines):
                        next_line = lines[i+j].strip()
                        if len(next_line.split()) > 3 and not any(bad in next_line.lower() for bad in bads):
                            rest = next_line
                            break
            rest = re.sub(r"\(.*?\)", "", rest)
            rest = re.sub(r"Fall \d{4}|Spring \d{4}|Section\s*[A-Za-z0-9]+", "", rest, flags=re.I).strip()
            return f"{code}: {rest}".strip(": ").replace("  ", " ")
    return ""

def extract_faculty_name(lines):
    for ln in lines:
        m = re.search(r"(Instructor|Professor|Faculty|Lecturer)\s*[:\-]?\s*([A-Za-z\.\-\s']+)", ln, re.I)
        if m:
            name = m.group(2)
            name = re.split(r",|Office|Email|Contact|Class", name)[0].strip()
            name = re.sub(r"[^A-Za-z\s'\-]", "", name)
            if len(name.split()) > 1:
                return name
    for ln in lines:
        if ln.strip().startswith("Dr. "):
            name = ln.strip().replace("Dr. ", "")
            name = re.split(r",|Office|Email|Contact|Class", name)[0].strip()
            name = re.sub(r"[^A-Za-z\s'\-]", "", name)
            if len(name.split()) > 1:
                return name
    return ""

def extract_faculty_email(lines):
    for ln in lines:
        if re.search(r"[a-zA-Z0-9._%+-]+@cpp\.edu", ln):
            return "Yes"
    return "No"

def extract_schedule(lines):
    pats = COLUMN_PATTERNS["Class schedule (day and time)?"]
    for ln in lines:
        if any(re.search(pat, ln, re.I) for pat in pats):
            return "Yes"
    return "No"

def extract_class_location(lines):
    pats = COLUMN_PATTERNS["Class location (building number & classroom number)"]
    for ln in lines:
        if any(re.search(pat, ln, re.I) for pat in pats):
            if re.search(r"\d", ln):
                return "Yes"
    return "No"

def extract_office_hours(lines):
    pats = COLUMN_PATTERNS["Offic hours?"]
    for ln in lines:
        if any(re.search(pat, ln, re.I) for pat in pats):
            return "Yes"
    return "No"

def extract_office_location(lines):
    pats = COLUMN_PATTERNS["Office location?"]
    for ln in lines:
        if any(re.search(pat, ln, re.I) for pat in pats) and "Hours" not in ln:
            if re.search(r"\d", ln):
                return "Yes"
    return "No"

def extract_learning_outcomes(lines):
    pats = COLUMN_PATTERNS["Course Learning Outcomes/Objectives included?"]
    for ln in lines:
        if any(re.search(pat, ln, re.I) for pat in pats):
            return "Yes"
    return "No"

def extract_modality(lines):
    for ln in lines:
        if re.search(r"hybrid.*asynchronous", ln, re.I):
            return "Hybrid Asynchronous"
        if re.search(r"hybrid.*synchronous", ln, re.I):
            return "Hybrid Synchronous"
        if re.search(r"in[- ]?person", ln, re.I):
            return "In-person"
        if re.search(r"asynchronous", ln, re.I):
            return "Asynchronous"
        if re.search(r"synchronous", ln, re.I):
            return "Synchronous"
        if re.search(r"online", ln, re.I):
            return "Online"
        if re.search(r"hybrid", ln, re.I):
            return "Hybrid"
    return ""

def extract_final_grade_components(lines):
    pats = COLUMN_PATTERNS["Final Grade components explained"]
    for idx, ln in enumerate(lines):
        if any(re.search(pat, ln, re.I) for pat in pats):
            if "%" in ln or re.search(r"\d+\s*%", ln):
                return "Yes"
            for j in range(1, 3):
                if idx+j < len(lines):
                    next_ln = lines[idx+j]
                    if "%" in next_ln or re.search(r"\d+\s*%", next_ln):
                        return "Yes"
    return "No"

def extract_weekly_schedule(lines):
    pats = COLUMN_PATTERNS["Weekly Schedule included?"]
    for ln in lines:
        if any(re.search(pat, ln, re.I) for pat in pats):
            return "Yes"
    return "No"

def check_50pct_inperson(lines):
    inperson_count = 0
    total_count = 0
    week_lines = []
    for ln in lines:
        if re.search(r"(Week|Session)\s*\d+", ln, re.I):
            week_lines.append(ln)
    if not week_lines:
        week_lines = [ln for ln in lines if re.search(r"In[- ]?person|F2F|Face[- ]?to[- ]?Face", ln, re.I)]
    for ln in week_lines:
        total_count += 1
        if re.search(r"In[- ]?person|F2F|Face[- ]?to[- ]?Face", ln, re.I):
            inperson_count += 1
    if total_count < TOTAL_SESSIONS:
        return "No", "schedule/class dates not explicit"
    if inperson_count >= MIN_INPERSON_SESSIONS:
        return "Yes", ""
    else:
        if inperson_count == 0:
            return "No", "no in-person sessions"
        else:
            return "No", ""

def analyze_one_file_strict(path, template_cols, assistant_name=ASSISTANT_NAME):
    text = extract_text_generic(path)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    row = {c: "" for c in template_cols}
    notes = []
    row["Student Assistants' Name (who works on the sheet)"] = assistant_name
    for col in template_cols:
        if col == "Student Assistants' Name (who works on the sheet)":
            continue
        if col == "Course Name & Number":
            row[col] = extract_course_name_number(lines)
        elif col == "Faculty Name":
            row[col] = extract_faculty_name(lines)
        elif col == "Faculty CPP email included?":
            row[col] = extract_faculty_email(lines)
        elif col == "Class schedule (day and time)?":
            row[col] = extract_schedule(lines)
        elif col == "Class location (building number & classroom number)":
            row[col] = extract_class_location(lines)
        elif col == "Offic hours?":
            row[col] = extract_office_hours(lines)
        elif col == "Office location?":
            row[col] = extract_office_location(lines)
        elif col == "Course Learning Outcomes/Objectives included?":
            row[col] = extract_learning_outcomes(lines)
        elif col == "Course modality specified?":
            row[col] = extract_modality(lines)
        elif col == "Final Grade components explained":
            row[col] = extract_final_grade_components(lines)
        elif col == "Weekly Schedule included?":
            val = extract_weekly_schedule(lines)
            row[col] = val
            if val == "No":
                notes.append("Weekly schedule not explicit")
        elif col == "Min. 50% in person class dates?":
            val, note = check_50pct_inperson(lines)
            row[col] = val
            if note:
                notes.append(note)
        else:
            row[col] = ""
    row["Notes"] = " | ".join(notes) if notes else ""
    return row

# --- Streamlit UI ---

st.set_page_config(page_title="Graduate Business School Syllabus Reviewer", layout="centered")
st.title("Graduate Business School Syllabus Reviewer")
st.markdown("""
Upload your **Excel/ODS template** and then one or more **syllabus files** (.pdf, .docx, .txt, .zip).
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

        st.markdown("### 🔍 **Preview: See Extracted Text for Each Syllabus**")
        for fp in syllabus_paths:
            st.write(f"---\n##### {os.path.basename(fp)}")
            with st.expander("Show extracted text"):
                txt = extract_text_generic(fp)
                st.text(txt[:2000] + ("\n... (truncated)" if len(txt) > 2000 else ""))

        st.markdown("---")
        if st.button("Process Syllabi & Download Excel"):
            rows = []
            for path in syllabus_paths:
                st.write(f"Analyzing: `{os.path.basename(path)}`")
                try:
                    row = analyze_one_file_strict(path, template_cols)
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
