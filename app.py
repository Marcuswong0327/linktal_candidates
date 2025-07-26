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
    This function now processes ONLY the current page content.
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
    
    # Extract first name from first line (guaranteed to be from this page only)
    if lines:
        first_line = lines[0].strip()
        if first_line:
            result['First Name'] = first_line
    
    # Process line by line for specific fields - ONLY within this page
    rfl_lines = []
    rfl_started = False
    
    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue
            
        # Check for CS pattern - extract only from the same line
        cs_match = re.search(r'CS:\s*([^\n\r]+)', line, re.IGNORECASE)
        if cs_match:
            extracted_cs = cs_match.group(1).strip()
            # Clean up - stop at next colon or end of meaningful content
            extracted_cs = re.split(r'\s+[A-Z]+:', extracted_cs)[0].strip()
            result['CS'] = extracted_cs
            continue
            
        # Check for ES pattern - extract only from the same line
        es_match = re.search(r'ES:\s*([^\n\r]+)', line, re.IGNORECASE)
        if es_match:
            extracted_es = es_match.group(1).strip()
            # Clean up - stop at next colon or end of meaningful content
            extracted_es = re.split(r'\s+[A-Z]+:', extracted_es)[0].strip()
            result['ES'] = extracted_es
            continue
            
        # Check for Notice Period or NP pattern - extract only from the same line
        np_match = re.search(r'(?:NOTICE PERIOD|NP):\s*([^\n\r]+)', line, re.IGNORECASE)
        if np_match:
            extracted_np = np_match.group(1).strip()
            # Clean up - stop at next colon or end of meaningful content
            extracted_np = re.split(r'\s+[A-Z]+:', extracted_np)[0].strip()
            result['Notice Period'] = extracted_np
            continue
            
        # Check for RFL pattern - this can span multiple lines
        rfl_match = re.search(r'RFL:\s*(.+)', line, re.IGNORECASE)
        if rfl_match:
            rfl_started = True
            rfl_content = rfl_match.group(1).strip()
            # Clean up first line of RFL - stop at next field
            rfl_content = re.split(r'\s+[A-Z]+:', rfl_content)[0].strip()
            if rfl_content:
                rfl_lines.append(rfl_content)
            continue
            
        # If RFL has started, continue collecting lines until we hit another field
        if rfl_started:
            # Stop RFL collection if we encounter another field pattern
            if re.match(r'^[A-Z]+:\s', line, re.IGNORECASE):
                rfl_started = False
                # Process this line for other patterns
                continue
            else:
                # Add this line to RFL if it's meaningful content
                if line and not re.match(r'^[A-Z]+:$', line):  # Not just a field label
                    rfl_lines.append(line)
    
    # Combine RFL lines and clean up
    if rfl_lines:
        combined_rfl = ' '.join(rfl_lines).strip()
        # Remove any trailing field patterns that might have been caught
        combined_rfl = re.split(r'\s+[A-Z]+:\s', combined_rfl)[0].strip()
        result['RFL'] = combined_rfl
    
    return result

def extract_pages_from_docx(docx_file) -> List[str]:
    """
    Extract text from each page of a DOCX file.
    Returns a list of page texts, properly separated by page breaks.
    """
    try:
        doc = Document(docx_file)
        pages = []
        current_page_content = []
        
        # Process each paragraph and look for page breaks
        for paragraph in doc.paragraphs:
            # Check if paragraph contains a page break
            page_break_found = False
            
            # Check for page break in the paragraph's runs
            for run in paragraph.runs:
                if run._element.xml.find('w:br') != -1:
                    # Check if it's a page break
                    if 'type="page"' in run._element.xml or 'w:type="page"' in run._element.xml:
                        page_break_found = True
                        break
            
            # Add paragraph text to current page
            text = paragraph.text.strip()
            if text:
                current_page_content.append(text)
            
            # If page break found, save current page and start new one
            if page_break_found:
                if current_page_content:
                    pages.append('\n'.join(current_page_content))
                    current_page_content = []
        
        # Add the last page if it has content
        if current_page_content:
            pages.append('\n'.join(current_page_content))
        
        # If no page breaks were detected, try alternative splitting methods
        if len(pages) == 1 and pages[0]:
            # Try to split by detecting person names/sections
            full_text = pages[0]
            
            # Method 1: Split by multiple empty lines (3 or more line breaks)
            sections = re.split(r'\n\s*\n\s*\n+', full_text)
            sections = [section.strip() for section in sections if section.strip()]
            
            if len(sections) > 1:
                pages = sections
            else:
                # Method 2: Smart detection based on name patterns and content structure
                lines = full_text.split('\n')
                pages = []
                current_section = []
                
                for i, line in enumerate(lines):
                    line = line.strip()
                    if not line:
                        continue
                    
                    # Detect potential new person/page based on patterns
                    is_likely_new_person = False
                    
                    # Check if this line looks like a name (at start of document or after sufficient content)
                    if (i == 0 or len(current_section) > 10):  # First line or after substantial content
                        # Pattern 1: All caps name (JOHN DOE)
                        if re.match(r'^[A-Z][A-Z\s]+[A-Z]$', line) and len(line.split()) >= 2:
                            is_likely_new_person = True
                        # Pattern 2: Title case name (John Doe)
                        elif re.match(r'^[A-Z][a-z]+\s+[A-Z][a-z]+(\s+[A-Z][a-z]+)*$', line):
                            is_likely_new_person = True
                        # Pattern 3: Mixed case but looks like a name
                        elif re.match(r'^[A-Z][a-zA-Z]+\s+[A-Z][a-zA-Z]+', line) and len(line.split()) <= 4:
                            # Additional validation: next few lines should contain typical resume content
                            next_lines = lines[i+1:i+5] if i+1 < len(lines) else []
                            has_resume_content = any(
                                any(keyword in next_line.upper() for keyword in ['CS:', 'ES:', 'NP:', 'NOTICE', 'RFL:', 'YEAR', 'EXPERIENCE']) 
                                for next_line in next_lines if next_line.strip()
                            )
                            if has_resume_content:
                                is_likely_new_person = True
                    
                    # If we detected a new person and have existing content, save current section
                    if is_likely_new_person and current_section and len(current_section) > 5:
                        pages.append('\n'.join(current_section))
                        current_section = []
                    
                    current_section.append(line)
                
                # Add the last section
                if current_section:
                    pages.append('\n'.join(current_section))
        
        # Final validation: ensure each page has reasonable content
        valid_pages = []
        for page in pages:
            if page.strip() and len(page.split()) >= 10:  # At least 10 words
                valid_pages.append(page.strip())
        
        return valid_pages
        
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
    output.seek(0)
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
                
                # Debug information - show page breakdown
                with st.expander("📋 Page Breakdown (Debug Info)", expanded=False):
                    for i, page in enumerate(pages, 1):
                        st.write(f"**Page {i} (Words: {len(page.split())}):**")
                        # Show first line (name) and word count
                        first_line = page.split('\n')[0] if page.split('\n') else "Empty"
                        st.write(f"- First line: `{first_line}`")
                        st.write(f"- Total words: {len(page.split())}")
                        # Show preview of content
                        preview = page[:200] + "..." if len(page) > 200 else page
                        st.text_area(f"Preview of Page {i}:", preview, height=100, key=f"preview_{i}")
                
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
