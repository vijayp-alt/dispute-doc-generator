import streamlit as st
import pandas as pd
from docx import Document
from docx2pdf import convert
import os
import tempfile
import base64

st.set_page_config(page_title="Dispute Document Generator", layout="centered")

# Display Nium logo
def get_image_base64(image_path):
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode()

# Header with logo and title side by side
logo_path = "nium_logo.webp"  # Place your Nium logo image in the same directory as this script

if os.path.exists(logo_path):
    logo_b64 = get_image_base64(logo_path)
    st.markdown(
        f"""
        <div style="display: flex; align-items: center; gap: 16px; margin-bottom: 8px;">
            <img src="data:image/webp;base64,{logo_b64}" width="80"/>
            <h1 style="margin: 0; font-size: 2rem;">Dispute Document Generator</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    # Fallback: just show the title if image not found
    st.warning("Logo image 'nium_logo.webp' not found. Place it in the same directory as this script.")
    st.title("📄 Dispute Document Generator")

st.markdown("---")

# Upload files
excel_file = st.file_uploader("Upload Excel File", type=["xls", "xlsx"])
word_file = st.file_uploader("Upload Word Template", type=["docx"])

# Start processing
if excel_file and word_file:
    try:
        df = pd.read_excel(excel_file, sheet_name=0, engine='openpyxl')
        st.success("Excel file loaded successfully.")
        st.dataframe(df)

        if 'Field name' in df.columns and 'Value' in df.columns:
            field_map = dict(zip(df['Field name'], df['Value']))

            if st.button("Generate PDF"):
                with tempfile.TemporaryDirectory() as tmpdir:
                    temp_doc_path = os.path.join(tmpdir, "updated.docx")
                    temp_pdf_path = os.path.join(tmpdir, "output.pdf")

                    # Load and update Word document
                    doc = Document(word_file)

                    # Replace placeholders in paragraphs
                    for paragraph in doc.paragraphs:
                        for key, value in field_map.items():
                            placeholder = f"<{key}>"
                            if placeholder in paragraph.text:
                                paragraph.text = paragraph.text.replace(placeholder, str(value))

                    # Replace placeholders in tables
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for key, value in field_map.items():
                                    placeholder = f"<{key}>"
                                    if placeholder in cell.text:
                                        cell.text = cell.text.replace(placeholder, str(value))

                    doc.save(temp_doc_path)

                    # Convert to PDF
                    convert(temp_doc_path, temp_pdf_path)

                    with open(temp_pdf_path, "rb") as f:
                        st.download_button("📥 Download PDF", f, file_name="Dispute_Document.pdf", mime="application/pdf")
        else:
            st.error("Excel file must contain 'Field name' and 'Value' columns.")
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Please upload both Excel and Word files to begin.")