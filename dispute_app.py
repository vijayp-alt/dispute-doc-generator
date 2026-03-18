import streamlit as st
import pandas as pd
from docx import Document
import subprocess
import os
import tempfile

st.set_page_config(page_title="Dispute Document Generator", layout="centered")

LOGO_URL = "https://raw.githubusercontent.com/vijayp-alt/dispute-doc-generator/main/image.webp"

# ── Global Dark Theme + Background Pattern ──────────────────────────────────
st.markdown("""
<style>
/* ── Page background with subtle dot pattern ── */
[data-testid="stAppViewContainer"] {
    background-color: #0f1117;
    background-image: radial-gradient(circle, #2a2d3a 1px, transparent 1px);
    background-size: 28px 28px;
}

[data-testid="stHeader"] {
    background-color: transparent;
}

/* ── Hide default Streamlit decorations ── */
#MainMenu, footer {visibility: hidden;}

/* ── Card style ── */
.card {
    background: #1c1f2e;
    border: 1px solid #2e3250;
    border-radius: 16px;
    padding: 28px 32px;
    margin-bottom: 24px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.4);
}

/* ── Header banner ── */
.header-banner {
    background: linear-gradient(135deg, #1a1d2e 0%, #12152a 100%);
    border: 1px solid #2e3250;
    border-radius: 16px;
    padding: 24px 32px;
    margin-bottom: 28px;
    display: flex;
    align-items: center;
    gap: 20px;
    box-shadow: 0 4px 32px rgba(0,0,0,0.5);
}

.header-title {
    color: #ffffff;
    font-size: 1.9rem;
    font-weight: 700;
    margin: 0;
    letter-spacing: -0.5px;
}

.header-subtitle {
    color: #7b82a8;
    font-size: 0.85rem;
    margin: 4px 0 0 0;
}

/* ── Section label ── */
.section-label {
    color: #a0a8cc;
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    margin-bottom: 10px;
}

/* ── Card title ── */
.card-title {
    color: #e2e6ff;
    font-size: 1.05rem;
    font-weight: 600;
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    gap: 8px;
}

/* ── Override Streamlit file uploader ── */
[data-testid="stFileUploader"] {
    background: #12152a !important;
    border: 1.5px dashed #3a3f6e !important;
    border-radius: 12px !important;
}

div[data-testid="stButton"] > button {
    background: linear-gradient(135deg, #6c63ff, #a78bfa);
    color: white;
    border: none;
    border-radius: 10px;
    padding: 12px 32px;
    font-size: 1rem;
    font-weight: 600;
    width: 100%;
    cursor: pointer;
    transition: opacity 0.2s;
}

div[data-testid="stButton"] > button:hover {
    opacity: 0.88;
}

p, label, span, div {
    color: #c8cde8;
}
</style>
""", unsafe_allow_html=True)

# ── Header Card ─────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="header-banner">
    <img src="{LOGO_URL}" width="72" onerror="this.style.display='none'" style="border-radius:10px;"/>
    <div>
        <p class="header-title">Dispute Document Generator</p>
        <p class="header-subtitle">Upload your Excel data and Word template to generate a dispute PDF instantly.</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Upload Card ──────────────────────────────────────────────────────────────
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="card-title">📂 Upload Files</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="section-label">Excel Data File</div>', unsafe_allow_html=True)
    excel_file = st.file_uploader("", type=["xls", "xlsx"], key="excel")
with col2:
    st.markdown('<div class="section-label">Word Template</div>', unsafe_allow_html=True)
    word_file = st.file_uploader("", type=["docx"], key="word")

st.markdown('</div>', unsafe_allow_html=True)

# ── Processing Card ──────────────────────────────────────────────────────────
if excel_file and word_file:
    try:
        df = pd.read_excel(excel_file, sheet_name=0, engine='openpyxl')

        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">📊 Data Preview</div>', unsafe_allow_html=True)
        st.success("✅ Excel file loaded successfully.")
        st.dataframe(df, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        if 'Field name' in df.columns and 'Value' in df.columns:
            field_map = dict(zip(df['Field name'], df['Value']))

            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="card-title">⚙️ Generate Document</div>', unsafe_allow_html=True)

            if st.button("🚀 Generate PDF"):
                with st.spinner("Processing your document..."):
                    with tempfile.TemporaryDirectory() as tmpdir:
                        temp_doc_path = os.path.join(tmpdir, "updated.docx")
                        temp_pdf_path = os.path.join(tmpdir, "output.pdf")

                        doc = Document(word_file)

                        for paragraph in doc.paragraphs:
                            for key, value in field_map.items():
                                placeholder = f"<{key}>"
                                if placeholder in paragraph.text:
                                    paragraph.text = paragraph.text.replace(placeholder, str(value))

                        for table in doc.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    for key, value in field_map.items():
                                        placeholder = f"<{key}>"
                                        if placeholder in cell.text:
                                            cell.text = cell.text.replace(placeholder, str(value))

                        doc.save(temp_doc_path)

                        # Convert to PDF using LibreOffice (works on Linux/Streamlit Cloud)
                        result = subprocess.run(
                            ["libreoffice", "--headless", "--convert-to", "pdf",
                             "--outdir", tmpdir, temp_doc_path],
                            capture_output=True, text=True
                        )
                        if result.returncode != 0:
                            raise Exception(f"LibreOffice conversion failed: {result.stderr}")

                        # LibreOffice names the output after the input file: updated.pdf
                        temp_pdf_path = os.path.join(tmpdir, "updated.pdf")

                        with open(temp_pdf_path, "rb") as f:
                            st.download_button(
                                "📥 Download Dispute PDF",
                                f,
                                file_name="Dispute_Document.pdf",
                                mime="application/pdf"
                            )

            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.error("⚠️ Excel file must contain 'Field name' and 'Value' columns.")

    except Exception as e:
        st.error(f"❌ Error: {e}")

else:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.info("👆 Please upload both an Excel file and a Word template above to get started.")
    st.markdown('</div>', unsafe_allow_html=True)
