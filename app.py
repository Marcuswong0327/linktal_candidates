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
    joined_text = " ".join(text_lines)

    def search(pattern):
        match = re.search(pattern, joined_text, re.IGNORECASE)
        return match.group(1).strip() if match else "NA"

    first_name = text_lines[0].strip() if text_lines else "NA"
    cs = search(r"CS[:\s]*([\d\.kK\+\s]+)")
    es = search(r"ES[:\s]*([\d\.kK\+\s]+)")
    notice_period = search(r"(?:Notice period|NP)[:\s]*((?:immediate|\d+\s*(?:day|week|month|year)s?)[\w\s]*)")
    rfl = search(r"RFL[:\s]*([\w\s,]+)")

    summary = joined_text

    return {
        "First Name": first_name,
        "CS": cs,
        "ES": es,
        "Notice Period": notice_period,
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
