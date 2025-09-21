import streamlit as st
import os
from tempfile import NamedTemporaryFile
from pypdf import PdfReader
from docx import Document as DocxDocument
import zipfile

def extract_text_pdf(path):
    text = []
    try:
        reader = PdfReader(path)
        for page in reader.pages:
            t = page.extract_text() or ""
            if t:
                text.append(t)
    except Exception:
        pass
    return "\n".join(text)

def extract_text_docx(path):
    try:
        doc = DocxDocument(path)
    except Exception:
        return ""
    parts = []
    for p in doc.paragraphs:
        if p.text:
            parts.append(p.text)
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

def save_and_list_files(uploaded_files):
    saved = []
    up_dir = "uploaded_files"
    os.makedirs(up_dir, exist_ok=True)
    for f in uploaded_files:
        out_path = os.path.join(up_dir, f.name)
        with open(out_path, "wb") as out_f:
            out_f.write(f.getbuffer())
        saved.append(out_path)
    # Expand ZIPs
    files = []
    for name in saved:
        if name.lower().endswith('.zip'):
            out_dir = os.path.join("unzipped", os.path.splitext(os.path.basename(name))[0])
            os.makedirs(out_dir, exist_ok=True)
            with zipfile.ZipFile(name) as z:
                z.extractall(out_dir)
            for root, _, fs in os.walk(out_dir):
                for f in fs:
                    if f.lower().endswith((".pdf", ".docx", ".txt")):
                        files.append(os.path.join(root, f))
        else:
            if name.lower().endswith((".pdf", ".docx", ".txt")):
                files.append(name)
    return files

st.title("Syllabus File Uploader & Reader")

uploaded_files = st.file_uploader("Upload syllabi files (PDF, DOCX, TXT, or ZIP)", type=['pdf', 'docx', 'txt', 'zip'], accept_multiple_files=True)
if uploaded_files:
    file_paths = save_and_list_files(uploaded_files)
    st.write(f"**Found {len(file_paths)} files:**")
    for fp in file_paths:
        st.write(f"---\n### {os.path.basename(fp)}")
        with st.expander("Show extracted text"):
            txt = extract_text_generic(fp)
            st.text(txt[:2000])  # Show first 2000 chars only for preview

    st.success("All files extracted and displayed above.")
    st.info("You can now process these for Excel output with your main app logic.")
else:
    st.info("Upload one or more files to begin.")
