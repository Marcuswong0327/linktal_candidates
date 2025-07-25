import streamlit as st
import docx
import re
import pandas as pd
from io import BytesIO

# Extract text from docx
def extract_text_from_docx(doc):
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text.strip())
    return full_text

# Parse candidate info per page
def parse_candidate_info(text_lines):
    first_name = text_lines[0].strip().title() if text_lines else "NA"

    cs = "NA"
    es = "NA"
    np = "NA"
    rfl = "NA"

    for line in text_lines:
        upper_line = line.upper()
        if upper_line.startswith("CS:"):
            cs = line.partition("CS:")[2].strip()
        elif upper_line.startswith("ES:"):
            es = line.partition("ES:")[2].strip()
        elif upper_line.startswith("NOTICE PERIOD:") or upper_line.startswith("NP:"):
            np = line.partition(":")[2].strip()

    full_text = "\n".join(text_lines)
    rfl_match = re.search(r"RFL[:\s]*(.*)", full_text, re.IGNORECASE | re.DOTALL)
    if rfl_match:
        rfl = rfl_match.group(1).strip()

    return {
        "First Name": first_name,
        "CS": cs,
        "ES": es,
        "Notice Period": np,
        "RFL/Reason for Leaving": rfl,
    }

# Streamlit app
st.title("Candidate Info Extractor from Word Document")

uploaded_file = st.file_uploader("Upload Word Document (.docx)", type="docx")

if uploaded_file:
    doc = docx.Document(uploaded_file)
    candidates = []

    page_lines = []
    for para in doc.paragraphs:
        line = para.text.strip()
        if not line:
            continue
        # Check for new page by capitalized name line (first line of page)
        if page_lines == []:
            page_lines.append(line)
        elif re.match(r"^[A-Z\s]{2,}$", line) and len(page_lines) >= 2:
            candidates.append(parse_candidate_info(page_lines))
            page_lines = [line]
        else:
            page_lines.append(line)

    if page_lines:
        candidates.append(parse_candidate_info(page_lines))

    df = pd.DataFrame(candidates)
    st.dataframe(df)

    output = BytesIO()
    df.to_excel(output, index=False)
    st.download_button("Download Excel", output.getvalue(), file_name="candidates.xlsx")
