Graduate Business School Syllabus Reviewer

A browser-based Streamlit app to batch-review graduate business school syllabi and output a detailed Excel summary for compliance or curriculum QA.

How to Use

Upload your Excel/ODS template
(with all your target review columns).

Upload one or more syllabus files
(.pdf, .docx, .txt, or a .zip of them).

App auto-extracts all fields using advanced pattern logic (and OpenAI GPT fallback for “Course Name & Number” and “Faculty Name” only if pattern fails).

Review, edit (optional), and download the output Excel.

Key Features

High-accuracy extraction for all compliance fields

Hybrid pattern + GPT logic for fields that are often missed by regex (e.g., instructor name, course code/title)

**No “guessing”—all outputs are either present in the document, filename, or returned as “Unknown” only if truly missing

100% editable table in the browser before you download

Requirements

See requirements.txt for details.
Works on Streamlit Cloud
 or locally with:

streamlit

pandas

openpyxl

odfpy

xlrd

python-docx

pypdf

openai

API Key Setup

For best results (LLM fallback), create an OpenAI API key
 and put it in .streamlit/secrets.toml like:

OPENAI_API_KEY = "sk-..."

Deployment

Push this repo to GitHub

Deploy to Streamlit Cloud or run locally:

streamlit run streamlit_app.py

Limitations

Only reads extractable (not scanned image-only) PDFs.

LLM fallback can be slow and may incur API costs.

Accuracy depends on syllabus formatting quality—edit in-browser if needed.

Credits

Built by Ismail Sholuke, Graduate Student Assistant.
For support or feature requests, open an Issue.
