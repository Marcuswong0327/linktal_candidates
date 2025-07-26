import streamlit as st
import pandas as pd
from docx import Document
import re
import io
from typing import List, Dict, Any
import base64

# Page configuration
st.set_page_config(
    page_title="Candidate Info Extractor from Word Document",
    page_icon="📄",
    layout="wide"
)

def extract_candidate_info_from_page(page_text: str) -> Dict[str, str]:
    """
    Extract candidate information from a single page text.
    Returns a dictionary with extracted fields.
    """
    lines = page_text.strip().split('\n')
    
    # Initialize result dictionary
    result = {
        'First Name': 'NA',
        'CS': 'NA',
        'ES': 'NA',
        'Notice Period': 'NA',
        'RFL': 'NA',
        'Summary': page_text.strip()
    }
    
    # Extract first name from first line
    if lines:
        first_line = lines[0].strip()
        if first_line:
            result['First Name'] = first_line
    
    # Process line by line for specific fields
    rfl_lines = []
    rfl_started = False
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Check for CS pattern
        cs_match = re.search(r'CS:\s*(.+)', line, re.IGNORECASE)
        if cs_match:
            result['CS'] = cs_match.group(1).strip()
            continue
            
        # Check for ES pattern
        es_match = re.search(r'ES:\s*(.+)', line, re.IGNORECASE)
        if es_match:
            result['ES'] = es_match.group(1).strip()
            continue
            
        # Check for Notice Period or NP pattern
        np_match = re.search(r'(?:NOTICE PERIOD|NP):\s*(.+)', line, re.IGNORECASE)
        if np_match:
            result['Notice Period'] = np_match.group(1).strip()
            continue
            
        # Check for RFL pattern
        rfl_match = re.search(r'RFL:\s*(.+)', line, re.IGNORECASE)
        if rfl_match:
            rfl_started = True
            rfl_lines.append(rfl_match.group(1).strip())
            continue
            
        # If RFL has started and current line doesn't contain other patterns, add to RFL
        if rfl_started and not any(pattern in line.upper() for pattern in ['CS:', 'ES:', 'NOTICE PERIOD:', 'NP:']):
            # Check if line starts with a new field pattern, if so stop RFL collection
            if re.match(r'^[A-Z]+:', line):
                rfl_started = False
            else:
                rfl_lines.append(line)
    
    # Combine RFL lines
    if rfl_lines:
        result['RFL'] = ' '.join(rfl_lines).strip()
    
    return result

def extract_pages_from_docx(docx_file) -> List[str]:
    """
    Extract text from each page of a DOCX file.
    Returns a list of page texts.
    """
    try:
        doc = Document(docx_file)
        pages = []
        current_page = []
        
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if text:
                current_page.append(text)
            
            # Check for page break (simple heuristic - empty paragraphs or page break)
            # Since python-docx doesn't easily detect page breaks, we'll use a different approach
            # We'll split based on patterns that indicate a new person/page
            
        # For now, we'll treat the entire document as pages separated by multiple empty lines
        # or when we encounter a pattern that looks like a new person
        full_text = '\n'.join([p.text for p in doc.paragraphs])
        
        # Split by multiple newlines or when we see a pattern that indicates new person
        sections = re.split(r'\n\s*\n\s*\n', full_text)
        
        # Filter out empty sections
        pages = [section.strip() for section in sections if section.strip()]
        
        # If no clear separation found, try to split by detecting names at start
        if len(pages) == 1:
            lines = full_text.split('\n')
            current_section = []
            pages = []
            
            for i, line in enumerate(lines):
                line = line.strip()
                if not line:
                    continue
                    
                # If this looks like a name (first line of a new section)
                # and we have content in current_section, start a new page
                if (i > 0 and 
                    len(current_section) > 5 and  # Minimum content threshold
                    re.match(r'^[A-Z][a-z]+ [A-Z][a-z]+$', line.strip())):  # Name pattern
                    
                    if current_section:
                        pages.append('\n'.join(current_section))
                        current_section = []
                
                current_section.append(line)
            
            # Add the last section
            if current_section:
                pages.append('\n'.join(current_section))
        
        return pages
        
    except Exception as e:
        st.error(f"Error reading DOCX file: {str(e)}")
        return []

def create_excel_download(df: pd.DataFrame) -> bytes:
    """
    Create an Excel file from DataFrame and return as bytes.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Candidate_Info')
    return output.getvalue()

def main():
    st.title("Candidate Info Extractor from Word Document")
    
    # File upload section
    st.subheader("Upload Word Document (.docx)")
    
    uploaded_file = st.file_uploader(
        "Drag and drop file here",
        type=['docx'],
        help="Limit 200MB per file • DOCX"
    )
    
    if uploaded_file is not None:
        # Display file info
        st.success(f"File uploaded: {uploaded_file.name}")
        st.info(f"File size: {uploaded_file.size / (1024*1024):.2f} MB")
        
        # Process button
        if st.button("Extract Information", type="primary"):
            with st.spinner("Processing document..."):
                # Extract pages from document
                pages = extract_pages_from_docx(uploaded_file)
                
                if not pages:
                    st.error("No content found in the document or unable to process the file.")
                    return
                
                st.success(f"Found {len(pages)} sections/pages in the document.")
                
                # Extract information from each page
                extracted_data = []
                
                for i, page_text in enumerate(pages):
                    if page_text.strip():  # Only process non-empty pages
                        candidate_info = extract_candidate_info_from_page(page_text)
                        extracted_data.append(candidate_info)
                
                if extracted_data:
                    # Create DataFrame
                    df = pd.DataFrame(extracted_data)
                    
                    # Reorder columns to match requirements
                    column_order = ['First Name', 'CS', 'ES', 'Notice Period', 'RFL', 'Summary']
                    df = df[column_order]
                    
                    # Display preview
                    st.subheader("Extracted Data Preview")
                    st.dataframe(df, use_container_width=True)
                    
                    # Download section
                    st.subheader("Download Extracted Data")
                    
                    # Create Excel file
                    excel_data = create_excel_download(df)
                    
                    # Download button
                    st.download_button(
                        label="Download Excel File",
                        data=excel_data,
                        file_name="candidate_information.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Display statistics
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Candidates", len(df))
                    with col2:
                        st.metric("Complete Profiles", len(df[df['First Name'] != 'NA']))
                    with col3:
                        cs_filled = len(df[df['CS'] != 'NA'])
                        st.metric("CS Information", f"{cs_filled}/{len(df)}")
                
                else:
                    st.warning("No candidate information could be extracted from the document.")
    
    # Instructions section
    with st.expander("How to use this tool"):
        st.markdown("""
        ### Instructions:
        1. **Upload a Word document (.docx)** containing candidate information
        2. **Document format**: One candidate per page/section
        3. **Required format**: 
           - First line should contain the candidate's name
           - Use patterns like `CS:`, `ES:`, `NOTICE PERIOD:` or `NP:`, `RFL:` for specific information
        4. **Click "Extract Information"** to process the document
        5. **Preview the extracted data** in the table
        6. **Download the Excel file** with all extracted information
        
        ### Extracted Fields:
        - **First Name**: Extracted from the first line of each section
        - **CS**: Compensation/salary information (looks for "CS:")
        - **ES**: Employment status or experience (looks for "ES:")
        - **Notice Period**: Notice period information (looks for "NOTICE PERIOD:" or "NP:")
        - **RFL**: Reason for leaving (looks for "RFL:", can span multiple lines)
        - **Summary**: Complete text of each candidate's section
        
        ### Notes:
        - Missing information will be marked as "NA"
        - The tool handles capitalized text format
        - Multiple sentences are supported for RFL field
        """)

if __name__ == "__main__":
    main()
