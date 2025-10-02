import streamlit as st
import pandas as pd
import os
import re
import zipfile
from io import BytesIO
from tempfile import NamedTemporaryFile
from pypdf import PdfReader
from docx import Document as DocxDocument

import openai

# ---------- API KEY ----------
openai.api_key = st.secrets["OPENAI_API_KEY"]
# For local testing, you could use:
# openai.api_key = st.text_input("Enter OpenAI API key", type="password")

COLUMN_PATTERNS = {
    "Course Name & Number": [r"GBA\s*\d{4}[A-Za-z]?", r"MSIS\s*\d{4}", r"MSHRL\s*\d{4}"],
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

def extract_faculty_name(lines, text):
    for ln in lines[:40]:
        m = re.search(r"(Instructor|Professor|Faculty|Lecturer)\s*[:\-]?\s*([A-Za-z\.\-\s']+)", ln, re.I)
        if m:
            name = m.group(2).strip()
            name = re.split(r"[,;|]|Email|Contact|Office|and", name, 1)[0].strip()
            name = re.sub(r"\b(Dr\.?|Prof\.?|Professor|Faculty|Lecturer)\b", "", name, flags=re.I).strip()
            if len(name.split()) >= 2 and all(w[0].isupper() for w in name.split()[:2]):
                return name
    for ln in lines[:20]:
        words = ln.split()
        for j in range(len(words)-1):
            if words[j][0].isupper() and words[j+1][0].isupper():
                possible_name = f"{words[j]} {words[j+1]}"
                if possible_name.lower() not in ["course title", "class meeting"]:
                    return possible_name
    return None

def fallback_gpt_faculty_name(text):
    try:
        resp = openai.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": f"Extract the main instructor's full name from the following syllabus text. Only return the name. If not found, reply Unknown.\n\n{text[:3500]}"}],
            max_tokens=50,
            temperature=0,
        )
        return resp.choices[0].message.content.strip()
    except Exception:
        return "Unknown"

def extract_course_name_number(lines, text, filename=""):
    code_pat = r"\b([A-Z]{3,6}\s*\d{4}[A-Za-z]?)\b"
    ignore_titles = ["syllabus", "fall", "spring", "section", "schedule", "course id", "p01", "p02", "video", "about", "research", "teaching", "office", "room", "canvas", "zoom", "assignment", "contact", "semester", "term"]
    for i, ln in enumerate(lines[:50]):
        m = re.search(code_pat, ln)
        if m:
            code = m.group(1).replace(" ", "")
            after = ln[m.end():].strip(" -:•")
            for sep in [":", "-", "–", "—"]:
                if sep in after:
                    title = after.split(sep, 1)[-1].strip()
                    if len(title.split()) > 2 and not any(b in title.lower() for b in ignore_titles):
                        return f"{code}: {title}"
            for j in range(i+1, min(i+4, len(lines))):
                next_ln = lines[j].strip()
                if next_ln and not any(b in next_ln.lower() for b in ignore_titles) and len(next_ln.split()) > 2:
                    return f"{code}: {next_ln}"
    m = re.search(code_pat, filename)
    if m:
        code = m.group(1).replace(" ", "")
        fn_title = filename[m.end():].replace("_", " ").replace("-", " ").strip(" ._")
        fn_title = re.sub(r"\.pdf|\.docx|\.txt", "", fn_title, flags=re.I)
        if fn_title and len(fn_title.split()) > 2 and not any(b in fn_title.lower() for b in ignore_titles):
            return f"{code}: {fn_title}"
    return None

def fallback_gpt_course_name_number(text):
    try:
        resp = openai.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": f"Extract the course code and full official title from the following syllabus, combine as 'CODE: Title'. Only return the result. If not found, reply Unknown.\n\n{text[:3500]}"}],
            max_tokens=80,
            temperature=0,
        )
        return resp.choices[0].message.content.strip()
    except Exception:
        return "Unknown"

def analyze_one_file_strict(path, template_cols):
    text = extract_text_generic(path)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    filename = os.path.basename(path)
    row = {c: "" for c in template_cols}
    for col in template_cols:
        if col == "Course Name & Number":
            val = extract_course_name_number(lines, text, filename)
            if not val or val.lower() == "unknown":
                val = fallback_gpt_course_name_number(text)
            row[col] = val if val and val.lower() != "unknown" else ""
        elif col == "Faculty Name":
            val = extract_faculty_name(lines, text)
            if not val or val.lower() == "unknown":
                val = fallback_gpt_faculty_name(text)
            row[col] = val if val and val.lower() != "unknown" else ""
        elif col == "Faculty CPP email included?":
            row[col] = "Yes" if re.search(r"[a-zA-Z0-9._%+-]+@cpp\.edu", text) else "No"
        elif col == "Class schedule (day and time)?":
            pats = COLUMN_PATTERNS["Class schedule (day and time)?"]
            found = any(re.search(pat, text, re.I) for pat in pats)
            row[col] = "Yes" if found else "No"
        elif col == "Class location (building number & classroom number)":
            pats = COLUMN_PATTERNS["Class location (building number & classroom number)"]
            found = any(re.search(pat, text, re.I) for pat in pats)
            row[col] = "Yes" if found else "No"
        elif col == "Offic hours?":
            pats = COLUMN_PATTERNS["Offic hours?"]
            found = any(re.search(pat, text, re.I) for pat in pats)
            row[col] = "Yes" if found else "No"
        elif col == "Office location?":
            pats = COLUMN_PATTERNS["Office location?"]
            found = any(re.search(pat, text, re.I) for pat in pats)
            row[col] = "Yes" if found else "No"
        elif col == "Course Learning Outcomes/Objectives included?":
            pats = COLUMN_PATTERNS["Course Learning Outcomes/Objectives included?"]
            found = any(re.search(pat, text, re.I) for pat in pats)
            row[col] = "Yes" if found else "No"
        elif col == "Course modality specified?":
            val = ""
            for kw in ["In-person", "Hybrid", "Online synchronous", "Online asynchronous", "Synchronous", "Asynchronous", "Remote", "Zoom"]:
                for ln in lines[:80]:
                    if kw.lower() in ln.lower():
                        val = ln.strip()
                        break
                if val:
                    break
            if not val:
                # fallback to LLM if nothing pattern matched
                try:
                    resp = openai.chat.completions.create(
                        model="gpt-4o",
                        messages=[{"role": "user", "content": f"Extract the exact course modality (e.g. In-person, Hybrid Synchronous, Online asynchronous) from the syllabus below. Only return the phrase.\n\n{text[:3500]}"}],
                        max_tokens=60,
                        temperature=0,
                    )
                    val = resp.choices[0].message.content.strip()
                except Exception:
                    val = ""
            row[col] = val
        elif col == "Final Grade components explained":
            pats = COLUMN_PATTERNS["Final Grade components explained"]
            found = any(re.search(pat, text, re.I) for pat in pats)
            row[col] = "Yes" if found else "No"
        elif col == "Weekly Schedule included?":
            pats = COLUMN_PATTERNS["Weekly Schedule included?"]
            found = any(re.search(pat, text, re.I) for pat in pats)
            row[col] = "Yes" if found else "No"
        elif col == "Min. 50% in person class dates?":
            session_lines = [ln for ln in lines if re.search(r"Week|Session", ln, re.I)]
            inperson = [ln for ln in session_lines if re.search(r"In[- ]?person|F2F|Face[- ]?to[- ]?Face", ln, re.I)]
            if session_lines and len(session_lines) >= 13:
                row[col] = "Yes" if len(inperson) >= 8 else "No"
            else:
                row[col] = "No"
        elif col == "Program":
            basename = os.path.basename(path).lower()
            programs = [
                ("MBA", "mba"),
                ("MSBA", "msba"),
                ("Digital Supply Chain", "digital supply chain"),
                ("MS Digital Supply Chain", "ms digital supply chain"),
                ("GBA", "gba"),
                ("MSDM", "msdm"),
                ("MSHRL", "mshrl"),
                ("MSIS", "msis"),
            ]
            val = ""
            for program_label, key in programs:
                if key in basename:
                    val = program_label
            row[col] = val
        else:
            row[col] = ""
    return row

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
    saved_paths = save_uploaded_files(uploaded_files)
    syllabus_paths = gather_syllabus_paths(saved_paths)
    st.info(f"Found {len(syllabus_paths)} syllabus files to process.")

    st.markdown("### 🔍 **Preview: See Extracted Text for Each Syllabus**")
    for fp in syllabus_paths:
        st.write(f"---\n##### {os.path.basename(fp)}")
        with st.expander("Show extracted text"):
            txt = extract_text_generic(fp)
            st.code(txt[:2000] + ("\n... (truncated)" if len(txt) > 2000 else ""), language="text")

    st.markdown("---")
    # Stateful processing & edit logic
    if st.button("Process Syllabi & Edit/Download Excel") or "df_out" not in st.session_state:
        with st.spinner("Processing uploaded files..."):
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
            st.session_state["df_out"] = df_out
    else:
        df_out = st.session_state["df_out"]

    # Editable table
    st.markdown("#### **Edit any cell below before downloading:**")
    edited_df = st.data_editor(
        df_out,
        use_container_width=True,
        num_rows="dynamic",
        key="editable_excel"
    )
    towrite = BytesIO()
    edited_df.to_excel(towrite, index=False)
    towrite.seek(0)
    st.download_button(
        label="Download Excel Output",
        data=towrite,
        file_name="syllabus_review_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload one or more syllabus files to begin.")
