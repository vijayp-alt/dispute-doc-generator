import streamlit as st
import pandas as pd
from docx import Document
import os
import tempfile
import platform
import subprocess
from io import BytesIO

# Try to import docx2pdf (available on Windows/macOS)
try:
    from docx2pdf import convert as docx2pdf_convert
except Exception:
    docx2pdf_convert = None

st.set_page_config(page_title="Dispute Document Generator", layout="centered")
st.title("📄 Dispute Document Generator")

# ---- Helpers ----
def replace_placeholders_in_paragraph(paragraph, field_map):
    """
    Replace placeholders like <Field name> inside a paragraph without losing formatting.
    Works across runs.
    """
    # Build a plain string concatenating all runs
    full_text = "".join(run.text for run in paragraph.runs)
    replaced = False
    for key, value in field_map.items():
        placeholder = f"<{key}>"
        if placeholder in full_text:
            full_text = full_text.replace(placeholder, str(value))
            replaced = True

    if replaced:
        # Clear all runs and write back as a single run
        # (This still may merge formatting; preserving mixed formatting with exact
        # placeholder boundaries is non-trivial. For most placeholders this is OK.)
        for _ in range(len(paragraph.runs)):
            paragraph.runs[0].text = ""
            paragraph._p.remove(paragraph.runs[0]._r)  # remove run
        paragraph.add_run(full_text)

def replace_placeholders_in_doc(doc, field_map):
    # Paragraphs
    for paragraph in doc.paragraphs:
        replace_placeholders_in_paragraph(paragraph, field_map)

    # Tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_placeholders_in_paragraph(p, field_map)

def linux_soffice_available():
    if platform.system().lower() != "linux":
        return False
    try:
        subprocess.run(["soffice", "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        return True
    except Exception:
        return False

def convert_docx_to_pdf(input_path, output_path):
    """
    Cross-platform PDF conversion:
    - Windows/macOS: docx2pdf (requires Microsoft Word)
    - Linux: LibreOffice `soffice` if available
    Returns (success: bool, message: str)
    """
    system = platform.system().lower()

    # Windows / macOS via docx2pdf
    if system in ("windows", "darwin") and docx2pdf_convert is not None:
        try:
            docx2pdf_convert(input_path, output_path)
            return True, "Converted via docx2pdf."
        except Exception as e:
            return False, f"docx2pdf failed: {e}"

    # Linux via LibreOffice
    if system == "linux" and linux_soffice_available():
        try:
            outdir = os.path.dirname(output_path)
            # LibreOffice writes output into outdir with the same base name .pdf
            subprocess.run(
                ["soffice", "--headless", "--convert-to", "pdf", "--outdir", outdir, input_path],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True
            )
            # Ensure the expected output exists; if LO changes the name, find it
            expected_pdf = os.path.join(outdir, os.path.splitext(os.path.basename(input_path))[0] + ".pdf")
            if os.path.exists(expected_pdf):
                # Rename to the requested output_path if necessary
                if expected_pdf != output_path:
                    os.replace(expected_pdf, output_path)
                return True, "Converted via LibreOffice."
            return False, "LibreOffice reported success but PDF not found."
        except subprocess.CalledProcessError as e:
            return False, f"LibreOffice conversion failed: {e.stderr.decode('utf-8', errors='ignore')}"
        except Exception as e:
            return False, f"LibreOffice conversion error: {e}"

    # No supported converter
    return False, f"PDF conversion not available on {platform.system()}."

# ---- UI ----
excel_file = st.file_uploader("Upload Excel File", type=["xls", "xlsx"])
word_file = st.file_uploader("Upload Word Template (.docx)", type=["docx"])

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
                    replace_placeholders_in_doc(doc, field_map)
                    doc.save(temp_doc_path)

                    # Try to convert to PDF
                    success, msg = convert_docx_to_pdf(temp_doc_path, temp_pdf_path)

                    if success and os.path.exists(temp_pdf_path):
                        with open(temp_pdf_path, "rb") as f:
                            st.download_button(
                                "📥 Download PDF",
                                f,
                                file_name="Dispute_Document.pdf",
                                mime="application/pdf"
                            )
                        st.caption(msg)
                    else:
                        # If conversion fails, offer the updated .docx for download
                        with open(temp_doc_path, "rb") as f:
                            st.download_button(
                                "📥 Download updated .docx (PDF conversion unavailable)",
                                f,
                                file_name="Dispute_Document.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        st.warning(
                            "PDF conversion is not available in this environment. "
                            "If you're on Linux, install LibreOffice, or run the app on Windows/macOS with Microsoft Word installed.\n\n"
                            f"Details: {msg}"
                        )
        else:
            st.error("Excel file must contain 'Field name' and 'Value' columns.")
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("Please upload both Excel and Word files to begin.")