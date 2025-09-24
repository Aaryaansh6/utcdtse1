# -*- coding: utf-8 -*-
"""
Streamlit App: Universal File-to-Text Converter

This application provides a web interface to upload a file, convert its
content to plain text or Markdown, display a preview, and offer the full
text as a downloadable file.

Supported formats:
- Word (.docx)
- Excel (.xlsx)
- PowerPoint (.pptx)
- HTML (.html)
- ZIP archives (.zip) containing supported files
- Plain Text (.txt)
"""

# ==============================================================================
# 1. IMPORTS
# ==============================================================================
import streamlit as st
import os
import zipfile
import docx
import openpyxl
from pptx import Presentation
from bs4 import BeautifulSoup
from markdownify import markdownify as md
import io

# ==============================================================================
# 2. CORE CONVERSION FUNCTION (Adapted from Colab version)
# ==============================================================================

def convert_file_to_text(file_name, file_stream):
    """
    A universal converter that processes a file stream and extracts text.

    It identifies the file type based on its extension and uses the appropriate
    library to extract text content. For ZIP archives, it recursively processes
    the files within.

    Args:
        file_name (str): The name of the file (e.g., 'document.docx').
        file_stream (streamlit.UploadedFile or io.BytesIO): A byte stream of the file's content.

    Returns:
        str: The extracted text content.
    """
    # Get the file extension to determine the processing method
    _, file_extension = os.path.splitext(file_name)
    file_extension = file_extension.lower()
    
    extracted_text = ""

    try:
        # --- Handle Word documents ---
        if file_extension == '.docx':
            doc = docx.Document(file_stream)
            full_text = [para.text for para in doc.paragraphs]
            extracted_text = '\n'.join(full_text)

        # --- Handle Excel spreadsheets ---
        elif file_extension == '.xlsx':
            workbook = openpyxl.load_workbook(file_stream)
            full_text = []
            for sheet in workbook.worksheets:
                for row in sheet.iter_rows():
                    row_text = '\t'.join([str(cell.value) for cell in row if cell.value is not None])
                    full_text.append(row_text)
            extracted_text = '\n'.join(full_text)

        # --- Handle PowerPoint presentations ---
        elif file_extension == '.pptx':
            presentation = Presentation(file_stream)
            full_text = []
            for slide in presentation.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        full_text.append(shape.text)
            extracted_text = '\n---\n'.join(full_text)

        # --- Handle HTML files (convert to Markdown) ---
        elif file_extension == '.html':
            html_content = file_stream.read().decode('utf-8', 'ignore')
            # Use markdownify to convert HTML body to Markdown
            extracted_text = md(html_content, heading_style="ATX")

        # --- Handle ZIP archives ---
        elif file_extension == '.zip':
            full_text = []
            with zipfile.ZipFile(file_stream, 'r') as zip_ref:
                for name in zip_ref.namelist():
                    if not name.startswith('__MACOSX/') and not name.endswith('/'):
                        with zip_ref.open(name) as member_file:
                            member_stream = io.BytesIO(member_file.read())
                            st.info(f"  -> Processing '{name}' from ZIP archive...")
                            member_text = convert_file_to_text(name, member_stream)
                            full_text.append(f"--- Content from: {name} ---\n{member_text}")
            extracted_text = "\n\n".join(full_text)
            
        # --- Handle plain text files ---
        elif file_extension == '.txt':
            extracted_text = file_stream.read().decode('utf-8', 'ignore')

        else:
            extracted_text = f"File type '{file_extension}' is not supported."

    except Exception as e:
        st.error(f"An error occurred while processing '{file_name}': {e}")
        return ""

    return extracted_text

# ==============================================================================
# 3. STREAMLIT UI
# ==============================================================================

def main():
    """Defines the Streamlit application's user interface and logic."""
    
    # --- Page Configuration ---
    st.set_page_config(
        page_title="Universal File Converter",
        page_icon="📄",
        layout="centered"
    )

    # --- Header ---
    st.title("📄 Universal File-to-Text Converter")
    st.markdown("Drag, drop, and download. Convert DOCX, XLSX, PPTX, HTML, and ZIP files to plain text instantly.")

    # --- File Uploader ---
    uploaded_file = st.file_uploader(
        "Drag and drop your file here",
        type=['docx', 'xlsx', 'pptx', 'html', 'zip', 'txt'],
        label_visibility="collapsed"
    )

    # --- Processing and Display Logic ---
    if uploaded_file is not None:
        file_name = uploaded_file.name
        
        # Show a spinner while the file is being processed
        with st.spinner(f"Converting '{file_name}'..."):
            converted_text = convert_file_to_text(file_name, uploaded_file)
        
        if converted_text:
            st.success("Conversion successful!")

            # --- Display Preview ---
            st.subheader("Preview (First 1000 characters)")
            st.text_area(
                "Preview",
                converted_text[:1000],
                height=250,
                label_visibility="collapsed"
            )

            # --- Download Button ---
            base_name, _ = os.path.splitext(file_name)
            download_filename = f"converted_{base_name}.txt"
            
            st.download_button(
                label="⬇️ Download Full Text",
                data=converted_text.encode('utf-8'),
                file_name=download_filename,
                mime='text/plain',
                use_container_width=True
            )

# ==============================================================================
# 4. MAIN EXECUTION BLOCK
# ==============================================================================

if __name__ == "__main__":
    main()
