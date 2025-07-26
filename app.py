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

def extract_candidate_info_from_bold_text(section_data: Dict[str, str]) -> Dict[str, str]:
    """
    Extract candidate information specifically from bold text content.
    Returns a dictionary with extracted fields.
    """
    bold_text = section_data.get('bold_text', '')
    name = section_data.get('name', 'NA')
    full_text = section_data.get('full_text', '')
    
    # Initialize result dictionary
    result = {
        'First Name': name,
        'CS': 'NA',
        'ES': 'NA',
        'Notice Period': 'NA',
        'RFL': 'NA',
        'Summary': full_text
    }
    
    # Process bold text line by line for specific fields
    lines = bold_text.strip().split('\n')
    rfl_lines = []
    rfl_started = False
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Check for CS pattern (Current Salary)
        cs_match = re.search(r'CS:\s*(.+?)(?:\s+[A-Z]+:|$)', line, re.IGNORECASE)
        if cs_match:
            result['CS'] = cs_match.group(1).strip()
            continue
            
        # Check for ES pattern (Expected Salary)
        es_match = re.search(r'ES:\s*(.+?)(?:\s+[A-Z]+:|$)', line, re.IGNORECASE)
        if es_match:
            result['ES'] = es_match.group(1).strip()
            continue
            
        # Check for Notice Period patterns
        np_match = re.search(r'(?:NOTICE PERIOD|NOTICE|NP):\s*(.+?)(?:\s+[A-Z]+:|$)', line, re.IGNORECASE)
        if np_match:
            result['Notice Period'] = np_match.group(1).strip()
            continue
            
        # Check for RFL pattern (Reason for Leaving)
        rfl_match = re.search(r'RFL:\s*(.+)', line, re.IGNORECASE)
        if rfl_match:
            rfl_started = True
            rfl_content = rfl_match.group(1).strip()
            # Clean up - stop at next field if it appears on same line
            rfl_content = re.split(r'\s+[A-Z]+:', rfl_content)[0].strip()
            if rfl_content:
                rfl_lines.append(rfl_content)
            continue
            
        # If RFL has started, continue collecting lines until we hit another field
        if rfl_started:
            # Stop RFL collection if we encounter another field pattern
            if re.match(r'^[A-Z]+:\s', line, re.IGNORECASE):
                rfl_started = False
                # Process this line for other patterns by continuing the loop
                continue
            else:
                # Add this line to RFL if it doesn't start with a field pattern
                if not re.match(r'^[A-Z]+:$', line):
                    rfl_lines.append(line)
    
    # Combine RFL lines
    if rfl_lines:
        result['RFL'] = ' '.join(rfl_lines).strip()
    
    return result

def extract_bold_text_from_docx(docx_file) -> List[Dict[str, str]]:
    """
    Extract only bold text from DOCX file and organize by person/section.
    Returns a list of dictionaries with bold text content per person.
    """
    try:
        doc = Document(docx_file)
        sections = []
        current_section_bold_text = []
        current_section_all_text = []
        
        for paragraph in doc.paragraphs:
            paragraph_text = paragraph.text.strip()
            if not paragraph_text:
                continue
                
            # Extract bold text from this paragraph
            bold_text_parts = []
            for run in paragraph.runs:
                if run.bold and run.text.strip():
                    bold_text_parts.append(run.text.strip())
            
            # Add all text for context
            current_section_all_text.append(paragraph_text)
            
            # If we found bold text, add it to current section
            if bold_text_parts:
                bold_line = ' '.join(bold_text_parts)
                current_section_bold_text.append(bold_line)
            
            # Check if this might be a new person (name pattern at start)
            # Simple heuristic: if we have accumulated substantial content and see a name pattern
            if (len(current_section_all_text) > 10 and 
                re.match(r'^[A-Z][a-zA-Z]+\s+[A-Z][a-zA-Z]+', paragraph_text) and
                len(paragraph_text.split()) <= 4):
                
                # Save previous section if it has content
                if current_section_bold_text or current_section_all_text:
                    sections.append({
                        'bold_text': '\n'.join(current_section_bold_text),
                        'full_text': '\n'.join(current_section_all_text[:-1]),  # Exclude current line
                        'name': current_section_all_text[0] if current_section_all_text else 'Unknown'
                    })
                
                # Start new section
                current_section_bold_text = []
                current_section_all_text = [paragraph_text]
        
        # Add the last section
        if current_section_bold_text or current_section_all_text:
            sections.append({
                'bold_text': '\n'.join(current_section_bold_text),
                'full_text': '\n'.join(current_section_all_text),
                'name': current_section_all_text[0] if current_section_all_text else 'Unknown'
            })
        
        # If no clear sections found, treat whole document as one section
        if not sections:
            all_bold_text = []
            all_text = []
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    all_text.append(paragraph.text.strip())
                    
                    bold_parts = []
                    for run in paragraph.runs:
                        if run.bold and run.text.strip():
                            bold_parts.append(run.text.strip())
                    
                    if bold_parts:
                        all_bold_text.append(' '.join(bold_parts))
            
            if all_bold_text:
                sections.append({
                    'bold_text': '\n'.join(all_bold_text),
                    'full_text': '\n'.join(all_text),
                    'name': all_text[0] if all_text else 'Unknown'
                })
        
        return sections
        
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
                # Extract bold text sections from document
                sections = extract_bold_text_from_docx(uploaded_file)
                
                if not sections:
                    st.error("No content or bold text found in the document.")
                    return
                
                st.success(f"Found {len(sections)} sections with bold text in the document.")
                
                # Show bold text preview for debugging
                with st.expander("📋 Bold Text Preview (Debug Info)", expanded=False):
                    for i, section in enumerate(sections, 1):
                        st.write(f"**Section {i} - {section['name']}:**")
                        st.write("Bold text found:")
                        st.text_area(f"Bold content {i}:", section['bold_text'], height=100, key=f"bold_{i}")
                
                # Extract information from each section's bold text
                extracted_data = []
                
                for section in sections:
                    if section['bold_text'].strip():  # Only process sections with bold text
                        candidate_info = extract_candidate_info_from_bold_text(section)
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
        2. **Document format**: Information should be in **BOLD TEXT** in the Word document
        3. **Required format**: 
           - Candidate names should be at the beginning of each section
           - **BOLD TEXT ONLY**: Use patterns like `CS:`, `ES:`, `NOTICE PERIOD:` or `NP:`, `RFL:` in bold format
        4. **Click "Extract Information"** to process the document
        5. **Preview the extracted data** in the table
        6. **Download the Excel file** with all extracted information
        
        ### Extracted Fields (from BOLD text only):
        - **First Name**: Extracted from the first line of each section
        - **CS**: Current salary information (looks for "CS:" in bold)
        - **ES**: Expected salary (looks for "ES:" in bold)
        - **Notice Period**: Notice period information (looks for "NOTICE PERIOD:" or "NP:" in bold)
        - **RFL**: Reason for leaving (looks for "RFL:" in bold, can span multiple lines)
        - **Summary**: Complete text of each candidate's section
        
        ### Important Notes:
        - **Only BOLD text** is processed for CS, ES, NP, and RFL extraction
        - This significantly improves accuracy by focusing on formatted key information
        - Missing information will be marked as "NA"
        - The debug section shows exactly what bold text was found
        """)

if __name__ == "__main__":
    main()
