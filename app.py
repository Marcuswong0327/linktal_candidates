import streamlit as st
import docx
import re
import pandas as pd
from io import BytesIO

# Function to extract text from a Word document
def extract_text_from_docx(doc):
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text.strip())
    return full_text

# Function to parse individual candidate page
def parse_candidate_info(text_lines):
    first_name = text_lines[0].strip().title() if text_lines else "NA"

    cs = "NA"
    es = "NA"
    np = "NA"
    rfl = "NA"

    joined_text = " ".join(text_lines)

    for line in text_lines:
        line_upper = line.upper()
        if line_upper.startswith("CS:"):
            cs = line.partition("CS:")[2].strip()
        elif line_upper.startswith("ES:"):
            es = line.partition("ES:")[2].strip()
        elif line_upper.startswith("NOTICE PERIOD:") or line_upper.startswith("NP:"):
            np = line.partition(":")[2].strip()

    rfl_match = re.search(r"RFL[:\s]*(.*?)(?=\n|$)", joined_text, re.IGNORECASE | re.DOTALL)
    if rfl_match:
        rfl = rfl_match.group(1).strip()

    summary = joined_text

    return {
        "First Name": first_name,
        "CS": cs,
        "ES": es,
        "Notice Period": np,
        "RFL/Reason for Leaving": rfl,
        "Summary": summary
    }

# Streamlit app
st.title("Candidate Info Extractor from Word Document")

uploaded_file = st.file_uploader("Upload Word Document (.docx)", type="docx")

if uploaded_file:
    doc = docx.Document(uploaded_file)

    # Split by pages or assume each page starts with a name line
    candidates = []
    page_lines = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            if re.match(r"^[A-Z\s]{2,}$", text) and page_lines:
                candidates.append(parse_candidate_info(page_lines))
                page_lines = [text]  # New page starts
            else:
                page_lines.append(text)

    if page_lines:
        candidates.append(parse_candidate_info(page_lines))

    df = pd.DataFrame(candidates)
    st.dataframe(df)

    output = BytesIO()
    df.to_excel(output, index=False)
    st.download_button("Download Excel", output.getvalue(), file_name="candidates.xlsx")
