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

# --- Extraction logic ---

def merge_broken_lines(lines, threshold=30):
    merged_lines = []
    buf = ""
    for ln in lines:
        if len(ln) < threshold:
            buf += ln + " "
        else:
            if buf:
                merged_lines.append(buf.strip())
                buf = ""
            merged_lines.append(ln.strip())
    if buf:
        merged_lines.append(buf.strip())
    return merged_lines

def smart_cleanup_line(line):
    # Fix email addresses and other splits
    line = re.sub(r'\s*([a-zA-Z])\s+', r'\1', line)
    line = re.sub(r'@\s*cpp\s*\.?\s*edu', '@cpp.edu', line)
    line = re.sub(r'(\d)\s*:\s*(\d{2})\s*p\s*m', r'\1:\2 pm', line)
    line = re.sub(r'(\d)\s*:\s*(\d{2})\s*a\s*m', r'\1:\2 am', line)
    line = re.sub(r'\s+', ' ', line)
    return line.strip()

def extract_course_name_number(lines):
    # Try normal match
    for i, ln in enumerate(lines[:40]):
        m = re.search(r'(GBA\s*\d{4}[A-Za-z]?)[\s:]*([^\d:]+)', ln)
        if m:
            course_num = m.group(1).replace(" ", "")
            course_name = m.group(2).strip(":- ").replace(":", "").strip()
            if len(course_name) > 2:
                return f"{course_num}: {course_name}"
    # Backup: look for GBA line, grab next non-GBA, non-numeric line
    for i, ln in enumerate(lines):
        if "GBA" in ln and i+1 < len(lines):
            next_line = lines[i+1]
            if not re.search(r'\d', next_line) and "GBA" not in next_line:
                course_num = re.search(r'(GBA\s*\d{4}[A-Za-z]?)', ln)
                if course_num:
                    return f"{course_num.group(1).replace(' ', '')}: {next_line.strip()}"
    return ""

def extract_faculty_name(lines):
    stopwords = ["Class", "Office", "Schedule", "location", "Information", "Email", "Format"]
    for ln in lines[:60]:
        m = re.search(r'(Instructor|Professor)[:\s]*(Dr\.?\s*)?([A-Z][a-zA-Z]+)([A-Z][a-zA-Z]+)?', ln)
        if m:
            s = ln.split(m.group(0))[-1]
            for stop in stopwords:
                if stop in s:
                    s = s.split(stop, 1)[0]
            names = [m.group(3)]
            if m.group(4): names.append(m.group(4))
            split_names = []
            for name in names:
                split_names += re.findall(r'[A-Z][a-z]+', name)
            return " ".join(split_names).strip()
        m2 = re.search(r'(Dr\.?\s*[A-Z][a-zA-Z]+)', ln)
        if m2:
            namepart = m2.group(0).replace('Dr.', '').strip()
            name_parts = re.findall(r'[A-Z][a-z]+', namepart)
            if name_parts:
                return " ".join(name_parts)
    return ""

def extract_email(lines):
    for ln in lines[:40]:
        if re.search(r"[a-zA-Z0-9._%+-]+@cpp\.edu", ln.replace(" ", "")):
            return "Yes"
    return "No"

def extract_schedule(lines):
    for ln in lines[:80]:
        if re.search(r'(Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*\s*\d{1,2}[:]\d{2}\s*[ap]m', ln, re.I):
            return "Yes"
        if re.search(r'W\d{1,2}:\d{2} ?pm', ln):  # compressed
            return "Yes"
    return "No"

def extract_class_location(lines):
    for ln in lines[:80]:
        if re.search(r"(Location:|Room|Building|Classroom|Rm)\s*\d+", ln, re.I):
            return "Yes"
    return "No"

def extract_office_hours(lines):
    for ln in lines[:100]:
        if "office hours" in ln.lower():
            return "Yes"
    return "No"

def extract_office_location(lines):
    for ln in lines[:100]:
        if ("office location" in ln.lower() or "office:" in ln.lower()) and "office hours" not in ln.lower():
            return "Yes"
    return "No"

def extract_learning_outcomes(lines):
    for ln in lines:
        if any(kw in ln.lower() for kw in ["learning objectives", "learning outcomes", "course objectives", "expected outcomes"]):
            return "Yes"
    return "No"

def extract_modality(lines):
    has_inperson, has_online, has_hybrid = False, False, False
    for ln in lines[:100]:
        l = ln.lower()
        if any(x in l for x in ["in-person", "inperson", "face-to-face"]):
            has_inperson = True
        if any(x in l for x in ["online", "zoom", "canvas", "synchronous"]):
            has_online = True
        if "hybrid" in l:
            has_hybrid = True
        if "hybridsynchronous" in l or (has_inperson and has_online):
            return "Hybrid Synchronous"
    if has_hybrid:
        return "Hybrid"
    if has_inperson:
        return "In-Person"
    if has_online:
        return "Online"
    return ""

def extract_grade_components(lines):
    for ln in lines:
        if any(word in ln.lower() for word in ["grading", "grade", "weight", "percentage", "points"]):
            if "%" in ln or re.search(r"\bpoints\b", ln, re.I):
                return "Yes"
    return "No"

def extract_weekly_schedule(lines):
    joined = "\n".join(lines)
    if re.search(r"\bWeek\s*\d+|Module\s*\d+|Session\s*\d+|Date\s+", joined, re.I):
        return "Yes"
    return "No"

def extract_50pct_in_person(lines):
    inperson_count = 0
    total_count = 0
    week_pattern = re.compile(r"\b(Week|Module|Session)\b", re.I)
    inperson_pattern = re.compile(r"\b(In[- ]?Person|F2F|Face[- ]?to[- ]?Face)\b", re.I)
    for ln in lines:
        if week_pattern.search(ln):
            total_count += 1
            if inperson_pattern.search(ln):
                inperson_count += 1
    if total_count >= 13 and (inperson_count / total_count) >= 0.5:
        return "Yes"
    return "No"

# --- File text extraction ---

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

def load_template_columns(template_path):
    if template_path.endswith('.ods'):
        df = pd.read_excel(template_path, engine="odf")
    else:
        df = pd.read_excel(template_path)
    return list(df.columns)

def process_file_final(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        raw_text = extract_text_pdf(path)
    elif ext == ".docx":
        raw_text = extract_text_docx(path)
    else:
        raw_text = extract_text_txt(path)
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
    merged = merge_broken_lines(lines, threshold=30)
    cleaned = [smart_cleanup_line(ln) for ln in merged]
    return cleaned

def analyze_file_from_lines(cleaned, template_cols, assistant_name="Ismail"):
    return {
        "Student Assistants' Name (who works on the sheet)": assistant_name,
        "Course Name & Number": extract_course_name_number(cleaned),
        "Faculty Name": extract_faculty_name(cleaned),
        "Faculty CPP email included?": extract_email(cleaned),
        "Class schedule (day and time)?": extract_schedule(cleaned),
        "Class location (building number & classroom number)": extract_class_location(cleaned),
        "Offic hours?": extract_office_hours(cleaned),
        "Office location?": extract_office_location(cleaned),
        "Course Learning Outcomes/Objectives included?": extract_learning_outcomes(cleaned),
        "Course modality specified?": extract_modality(cleaned),
        "Final Grade components explained": extract_grade_components(cleaned),
        "Weekly Schedule included?": extract_weekly_schedule(cleaned),
        "Min. 50% in person class dates?": extract_50pct_in_person(cleaned),
        "Notes": "",
    }

# --- Streamlit UI ---

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

debug_mode = st.checkbox("Debug Mode: Show merged & cleaned lines for each syllabus?")

if uploaded_files:
    with st.spinner("Processing uploaded files..."):
        saved_paths = save_uploaded_files(uploaded_files)
        syllabus_paths = gather_syllabus_paths(saved_paths)
        st.info(f"Found {len(syllabus_paths)} syllabus files to process.")

        rows = []
        for path in syllabus_paths:
            st.write(f"Analyzing: `{os.path.basename(path)}`")
            cleaned = process_file_final(path)
            if debug_mode:
                st.write("**Merged & Cleaned Lines Preview:**")
                for ln in cleaned[:50]:
                    st.write(ln)
            try:
                row = analyze_file_from_lines(cleaned, template_cols)
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
