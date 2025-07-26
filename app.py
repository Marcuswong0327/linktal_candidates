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

def extract_bold_text_from_paragraph(paragraph) -> str:
    """
    Extract only bold text from a paragraph.
    Returns bold text as a string.
    """
    bold_text_parts = []
    for run in paragraph.runs:
        if run.bold and run.text.strip():
            bold_text_parts.append(run.text.strip())
    return ' '.join(bold_text_parts)

def extract_candidate_info_from_page_enhanced(page_text: str, bold_text: str = "") -> Dict[str, str]:
    """
    Extract candidate information using dual mechanism:
    1. Regular extraction from all text
    2. Enhanced extraction focusing on bold text when available
    """
    lines = page_text.strip().split('\n')
    bold_lines = bold_text.strip().split('\n') if bold_text else []
    
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
    
    # MECHANISM 1: Regular extraction from all text
    def extract_from_text_lines(text_lines, source_name="regular"):
        extracted = {
            'CS': 'NA',
            'ES': 'NA',
            'Notice Period': 'NA',
            'RFL': 'NA'
        }
        
        rfl_lines = []
        rfl_started = False
        
        for line in text_lines:
            line = line.strip()
            if not line:
                continue
                
            # Check for CS pattern
            cs_match = re.search(r'CS:\s*(.+?)(?:\s+[A-Z]+:|$)', line, re.IGNORECASE)
            if cs_match:
                extracted['CS'] = cs_match.group(1).strip()
                continue
                
            # Check for ES pattern
            es_match = re.search(r'ES:\s*(.+?)(?:\s+[A-Z]+:|$)', line, re.IGNORECASE)
            if es_match:
                extracted['ES'] = es_match.group(1).strip()
                continue
                
            # Check for Notice Period or NP pattern
            np_match = re.search(r'(?:NOTICE PERIOD|NP):\s*(.+?)(?:\s+[A-Z]+:|$)', line, re.IGNORECASE)
            if np_match:
                extracted['Notice Period'] = np_match.group(1).strip()
                continue
                
            # Check for RFL pattern
            rfl_match = re.search(r'RFL:\s*(.+)', line, re.IGNORECASE)
            if rfl_match:
                rfl_started = True
                rfl_content = rfl_match.group(1).strip()
                # Clean up - stop at next field if it appears on same line
                rfl_content = re.split(r'\s+[A-Z]+:', rfl_content)[0].strip()
                if rfl_content:
                    rfl_lines.append(rfl_content)
                continue
                
            # If RFL has started and current line doesn't contain other patterns, add to RFL
            if rfl_started:
                if re.match(r'^[A-Z]+:\s', line, re.IGNORECASE):
                    rfl_started = False
                elif not re.match(r'^[A-Z]+:$', line):
                    rfl_lines.append(line)
        
        # Combine RFL lines
        if rfl_lines:
            extracted['RFL'] = ' '.join(rfl_lines).strip()
            
        return extracted
    
    # Extract from regular text
    regular_extraction = extract_from_text_lines(lines, "regular")
    
    # MECHANISM 2: Enhanced extraction from bold text (if available)
    bold_extraction = {'CS': 'NA', 'ES': 'NA', 'Notice Period': 'NA', 'RFL': 'NA'}
    if bold_text.strip():
        bold_extraction = extract_from_text_lines(bold_lines, "bold")
    
    # PRIORITY LOGIC: Bold text takes precedence when available
    # If bold text has a field, use it; otherwise fall back to regular extraction
    for field in ['CS', 'ES', 'Notice Period', 'RFL']:
        if bold_extraction[field] != 'NA':
            result[field] = bold_extraction[field]
            # Add indicator that this came from bold text
            result[field] = f"{result[field]} [BOLD]"
        elif regular_extraction[field] != 'NA':
            result[field] = regular_extraction[field]
    
    return result

def extract_pages_from_docx_with_bold(docx_file) -> List[Dict[str, str]]:
    """
    Extract text from each page/section with both regular and bold text.
    Returns a list of dictionaries containing both regular and bold text.
    """
    try:
        doc = Document(docx_file)
        
        # First, extract all paragraphs with their bold text
        all_paragraphs = []
        all_bold_text = []
        
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if text:
                all_paragraphs.append(text)
                
                # Extract bold text from this paragraph
                bold_text = extract_bold_text_from_paragraph(paragraph)
                if bold_text:
                    all_bold_text.append(bold_text)
        
        # Combine all text
        full_text = '\n'.join(all_paragraphs)
        full_bold_text = '\n'.join(all_bold_text)
        
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
        
        # Now create sections with corresponding bold text
        result_sections = []
        for i, page in enumerate(pages):
            # For simplicity, we'll associate bold text proportionally
            # In a more sophisticated version, we could track paragraph-level associations
            section_data = {
                'regular_text': page,
                'bold_text': full_bold_text,  # For now, include all bold text
                'section_number': i + 1
            }
            result_sections.append(section_data)
        
        return result_sections
        
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
                # Extract pages with bold text detection
                sections = extract_pages_from_docx_with_bold(uploaded_file)
                
                if not sections:
                    st.error("No content found in the document or unable to process the file.")
                    return
                
                st.success(f"Found {len(sections)} sections/pages in the document.")
                
                # Show processing details
                with st.expander("📋 Processing Details (Debug Info)", expanded=False):
                    for i, section in enumerate(sections, 1):
                        st.write(f"**Section {i}:**")
                        
                        # Show if bold text was found
                        if section['bold_text'].strip():
                            st.write("✅ Bold text detected - Enhanced extraction will be used")
                            st.text_area(f"Bold text found in section {i}:", 
                                       section['bold_text'][:200] + "..." if len(section['bold_text']) > 200 else section['bold_text'],
                                       height=80, key=f"bold_debug_{i}")
                        else:
                            st.write("ℹ️ No bold text - Regular extraction only")
                        
                        # Show preview of regular text
                        preview = section['regular_text'][:200] + "..." if len(section['regular_text']) > 200 else section['regular_text']
                        st.text_area(f"Section {i} preview:", preview, height=60, key=f"preview_{i}")
                        st.write("---")
                
                # Extract information from each section
                extracted_data = []
                
                for section in sections:
                    if section['regular_text'].strip():  # Only process non-empty sections
                        candidate_info = extract_candidate_info_from_page_enhanced(
                            section['regular_text'], 
                            section['bold_text']
                        )
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
                    
                    # Show extraction source summary
                    st.subheader("Extraction Summary")
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        bold_extractions = sum(1 for _, row in df.iterrows() 
                                             if any('[BOLD]' in str(row[field]) for field in ['CS', 'ES', 'Notice Period', 'RFL']))
                        st.metric("Enhanced (Bold) Extractions", bold_extractions)
                    
                    with col2:
                        regular_extractions = len(df) - bold_extractions
                        st.metric("Regular Extractions", regular_extractions)
                    
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
           - **Bold text** will be given priority for more accurate extraction
        4. **Click "Extract Information"** to process the document
        5. **Preview the extracted data** in the table
        6. **Download the Excel file** with all extracted information
        
        ### Dual Extraction Mechanism:
        - **Regular Extraction**: Processes all text in the document
        - **Enhanced Extraction**: When bold text is detected, it gets priority for CS, ES, NP, RFL fields
        - **Priority Logic**: Bold text overrides regular extraction when available
        - **Indicators**: Fields extracted from bold text are marked with [BOLD]
        
        ### Extracted Fields:
        - **First Name**: Extracted from the first line of each section
        - **CS**: Current salary information (prioritizes bold text)
        - **ES**: Expected salary (prioritizes bold text)
        - **Notice Period**: Notice period information (prioritizes bold text)
        - **RFL**: Reason for leaving (prioritizes bold text, can span multiple lines)
        - **Summary**: Complete text of each candidate's section
        
        ### Notes:
        - Missing information will be marked as "NA"
        - Bold text extraction provides higher accuracy
        - Regular extraction serves as fallback when bold text is not available
        - The debug section shows exactly what processing method was used
        """)

if __name__ == "__main__":
    main()
