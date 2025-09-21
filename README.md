# Graduate Business School Syllabus Reviewer

A Streamlit app to batch-review graduate business syllabi and output an Excel summary for compliance, accreditation, or curriculum review.

---

## How to Use

1. **Upload your Excel/ODS template** (with your review columns/questions).
2. **Upload one or more syllabus files** (`.pdf`, `.docx`, `.txt`, or a `.zip` of them).
3. **Process syllabi and review results:**
    - The app extracts all key information and displays an editable table.
    - **Edit any cell directly in the browser** before download (no need to open Excel separately!).
    - Use the “Extracted Text” preview for each file to help you verify or correct blanks.
4. **Download the final, reviewed Excel file** for records, reporting, or further QA.

---

## Features

- Upload and batch-analyze any number of syllabi (PDF, DOCX, TXT, or ZIP)
- Uses your Excel/ODS template for flexible, standards-driven review
- **Editable table**: Manually correct or fill any cell before download
- **Preview full extracted text** for each syllabus (for QA or to guide manual corrections)
- **Strict evidence-based auto-extraction**: Only fills “Yes” if information is explicit; blanks/unusual cases left for reviewer review
- Download final Excel for reporting, audit, or archiving

---

## Limitations

- **Scanned PDFs** (image files) are not supported—only digital PDFs, DOCX, or TXT files will be processed.
- For highly unusual or unstructured syllabi, **manual review in the browser is strongly recommended**.
- For full AI (semantic) reading, see local Colab/GPT-powered versions.

---

## Credits

Built by Ismail, Graduate Student Assistant.  
QA & documentation by [CBA].

---

## Contact

Questions? Raise an issue or contact [itsholuke@cpp.edu].
