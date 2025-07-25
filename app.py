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
    first_name = "NA"
    cs = "NA"
    es = "NA"
    np = "NA"
    rfl = "NA"

    # Look for the name after a line of underscores
    for i in range(len(text_lines) - 1):
        if re.match(r"^_+$", text_lines[i]):
            first_name = text_lines[i + 1].strip().title()
            break

    # Only extract CS, ES, NOTICE PERIOD from their specific line only
    for line in text_lines:
        line_upper = line.upper()
        if line_upper.startswith("CS:"):
            cs = line.partition("CS:")[2].strip()
        elif line_upper.startswith("ES:"):
            es = line.partition("ES:")[2].strip()
        elif line_upper.startswith("NOTICE PERIOD:") or line_upper.startswith("NP:"):
            np = line.partition(":")[2].strip()

    # RFL can be multiline
    joined_text = "\n".join(text_lines)
    rfl_match = re.search(r"RFL[:\s]*(.*)", joined_text, re.IGNORECASE | re.DOTALL)
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

    candidates = []
    page_lines = []
    first_line = True

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            if first_line:
                # Always treat the first non-empty line as a new candidate
                if page_lines:
                    candidates.append(parse_candidate_info(page_lines))
                page_lines = [text]
                first_line = False
            elif re.match(r"^[A-Z\s]{2,}$", text):
                # Treat capital line as possible new candidate only if there are enough prior lines
                if len(page_lines) >= 2:
                    candidates.append(parse_candidate_info(page_lines))
                    page_lines = [text]
                else:
                    page_lines.append(text)
            else:
                page_lines.append(text)

    if page_lines:
        candidates.append(parse_candidate_info(page_lines))

    df = pd.DataFrame(candidates)
    st.dataframe(df)

    output = BytesIO()
    df.to_excel(output, index=False)
    st.download_button("Download Excel", output.getvalue(), file_name="candidates.xlsx")
