import streamlit as st
import pandas as pd
from docx import Document
from docx2pdf import convert
import os
import re
import tempfile

st.set_page_config(page_title="Dispute Document Generator", layout="wide")

LOGO_URL = "https://raw.githubusercontent.com/vijayp-alt/dispute-doc-generator/main/image.webp"

# ─── Per-tab theme config ──────────────────────────────────────────────────────
TAB_THEMES = {
    "pos07": {
        "color":       "#1a6eb5",   # Blue
        "light":       "#e8f1fb",
        "emoji":       "🏧",
        "label":       "POS 07 — Card Present Dispute",
        "pdf":         "POS07_Dispute.pdf",
        "mandatory":   ["Cardholder Name", "Card Number", "Transaction Date",
                        "Transaction Amount", "Merchant Name"],
        "date_fields": ["Transaction Date", "Dispute Date"],
    },
    "3ds": {
        "color":       "#7b2fa8",   # Purple
        "light":       "#f3eafc",
        "emoji":       "🔐",
        "label":       "3DS — 3D Secure Dispute",
        "pdf":         "3DS_Dispute.pdf",
        "mandatory":   ["Cardholder Name", "Card Number", "Transaction Date",
                        "Transaction Amount", "3DS Reference Number"],
        "date_fields": ["Transaction Date", "Authentication Date"],
    },
    "pos05": {
        "color":       "#b55a00",   # Amber
        "light":       "#fff4e6",
        "emoji":       "💳",
        "label":       "POS 05 — Card Absent Dispute",
        "pdf":         "POS05_Dispute.pdf",
        "mandatory":   ["Cardholder Name", "Card Number", "Transaction Date",
                        "Transaction Amount", "Merchant Name", "Order Reference"],
        "date_fields": ["Transaction Date", "Dispute Date"],
    },
}

# Common fields auto-filled across all tabs
COMMON_FIELDS = ["Cardholder Name", "Card Number", "Transaction Date", "Transaction Amount"]

# ─── Global CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
button[data-baseweb="tab"] { font-size: 1rem; font-weight: 700; padding: 10px 28px; }
.tab-badge {
    display: inline-block; padding: 7px 22px; border-radius: 22px;
    font-weight: 700; font-size: 1.1rem; color: #fff; margin-bottom: 16px;
}
.common-banner {
    background: #f0f7ff; border-left: 4px solid #1a6eb5; border-radius: 6px;
    padding: 12px 18px; margin-bottom: 4px; font-size: 0.93rem;
}
.info-box  { background:#e8f5e9; border-left:4px solid #2e7d32; border-radius:6px; padding:10px 16px; margin:6px 0; }
.warn-box  { background:#fff8e1; border-left:4px solid #f9a825; border-radius:6px; padding:10px 16px; margin:6px 0; }
.err-box   { background:#fdecea; border-left:4px solid #c62828; border-radius:6px; padding:10px 16px; margin:6px 0; }
</style>
""", unsafe_allow_html=True)

# ─── Header ───────────────────────────────────────────────────────────────────
st.markdown(
    f"""
    <div style="display:flex;align-items:center;gap:16px;margin-bottom:8px;">
        <img src="{LOGO_URL}" width="80" onerror="this.style.display='none'"/>
        <h1 style="margin:0;font-size:2rem;">Dispute Document Generator</h1>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown("---")

# ─── Common fields panel ──────────────────────────────────────────────────────
if "common_values" not in st.session_state:
    st.session_state.common_values = {f: "" for f in COMMON_FIELDS}

with st.expander("🔗 Auto-fill Common Fields  (shared across all tabs)", expanded=True):
    st.markdown(
        '<div class="common-banner">Fill these once and they will be automatically applied to '
        '<b>all three tabs</b>, overriding any matching values in the uploaded Excel.</div>',
        unsafe_allow_html=True,
    )
    cols = st.columns(len(COMMON_FIELDS))
    for i, field in enumerate(COMMON_FIELDS):
        with cols[i]:
            st.session_state.common_values[field] = st.text_input(
                field,
                value=st.session_state.common_values[field],
                key=f"common_{field}",
                placeholder=f"{field}…",
            )

st.markdown("---")


# ─── Helpers ──────────────────────────────────────────────────────────────────
DATE_PATTERNS = [
    r"^\d{2}/\d{2}/\d{4}$",    # DD/MM/YYYY
    r"^\d{4}-\d{2}-\d{2}$",    # YYYY-MM-DD
    r"^\d{2}-\d{2}-\d{4}$",    # DD-MM-YYYY
    r"^\d{2}\.\d{2}\.\d{4}$",  # DD.MM.YYYY
]

def is_valid_date(value: str) -> bool:
    return any(re.match(p, value.strip()) for p in DATE_PATTERNS)


def replace_in_paragraph(paragraph, field_map):
    for key, value in field_map.items():
        placeholder = f"<{key}>"
        if placeholder in paragraph.text:
            full_text = "".join(run.text for run in paragraph.runs)
            if placeholder in full_text:
                new_text = full_text.replace(placeholder, str(value))
                if paragraph.runs:
                    paragraph.runs[0].text = new_text
                    for run in paragraph.runs[1:]:
                        run.text = ""


def replace_placeholders(doc, field_map):
    for para in doc.paragraphs:
        replace_in_paragraph(para, field_map)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_paragraph(para, field_map)
                for nested in cell.tables:
                    for nrow in nested.rows:
                        for ncell in nrow.cells:
                            for para in ncell.paragraphs:
                                replace_in_paragraph(para, field_map)


# ─── Tab renderer ─────────────────────────────────────────────────────────────
def render_tab(tab_key: str):
    t         = TAB_THEMES[tab_key]
    color     = t["color"]
    light     = t["light"]
    label     = t["label"]
    pdf_name  = t["pdf"]
    mandatory = t["mandatory"]
    date_flds = t["date_fields"]

    # Colored badge header
    st.markdown(
        f'<div class="tab-badge" style="background:{color};">{t["emoji"]} {label}</div>',
        unsafe_allow_html=True,
    )

    # Upload area with tab accent border
    st.markdown(
        f'<div style="background:{light};border:2px solid {color};'
        f'border-radius:10px;padding:20px;margin-bottom:14px;">',
        unsafe_allow_html=True,
    )
    c1, c2 = st.columns(2)
    with c1:
        excel_file = st.file_uploader(
            "📊 Upload Excel File (Field name / Value)",
            type=["xls", "xlsx"], key=f"{tab_key}_excel",
        )
    with c2:
        word_file = st.file_uploader(
            "📝 Upload Word Template (.docx)",
            type=["docx"], key=f"{tab_key}_word",
        )
    st.markdown("</div>", unsafe_allow_html=True)

    # Mandatory field reference
    with st.expander(f"ℹ️ Required fields for {label.split('—')[0].strip()}", expanded=False):
        st.markdown(
            f'<div style="background:{light};border-radius:8px;padding:12px;">'
            + "".join(f'<span style="display:inline-block;background:{color};color:#fff;'
                      f'border-radius:12px;padding:3px 12px;margin:3px;font-size:0.85rem;">'
                      f'{f}</span>' for f in mandatory)
            + "</div>",
            unsafe_allow_html=True,
        )

    if not (excel_file and word_file):
        msg = ("📎 Also upload a Word template." if excel_file
               else "📊 Also upload an Excel file." if word_file
               else "👆 Upload both files above to get started.")
        st.info(msg)
        return

    try:
        df = pd.read_excel(excel_file, sheet_name=0, engine="openpyxl")

        if "Field name" not in df.columns or "Value" not in df.columns:
            st.error("❌ Excel must contain **'Field name'** and **'Value'** columns.")
            return

        field_map = dict(zip(df["Field name"].astype(str), df["Value"].astype(str)))

        # ── 1. Auto-fill common fields ─────────────────────────────────────
        autofilled = [
            f for f, v in st.session_state.common_values.items() if v.strip()
        ]
        for field in autofilled:
            field_map[field] = st.session_state.common_values[field].strip()

        if autofilled:
            st.markdown(
                f'<div class="info-box">✅ <b>Auto-filled:</b> {", ".join(autofilled)}</div>',
                unsafe_allow_html=True,
            )

        # ── 2. Mandatory field checks ──────────────────────────────────────
        missing = [
            f for f in mandatory
            if not str(field_map.get(f, "")).strip()
            or str(field_map.get(f, "")).strip().lower() in ("nan", "none")
        ]
        if missing:
            st.markdown(
                '<div class="err-box">⛔ <b>Missing mandatory fields:</b><br>'
                + "".join(f"&nbsp;&nbsp;• {m}<br>" for m in missing)
                + "Add them to your Excel or fill the common fields above.</div>",
                unsafe_allow_html=True,
            )

        # ── 3. Date format validation ──────────────────────────────────────
        date_issues = []
        for df_field in date_flds:
            val = str(field_map.get(df_field, "")).strip()
            if val and val.lower() not in ("", "nan", "none") and not is_valid_date(val):
                date_issues.append(
                    f"<b>{df_field}</b>: <code>{val}</code> — use DD/MM/YYYY or YYYY-MM-DD"
                )
        if date_issues:
            st.markdown(
                '<div class="warn-box">⚠️ <b>Date format warnings:</b><br>'
                + "".join(f"&nbsp;&nbsp;• {w}<br>" for w in date_issues)
                + "</div>",
                unsafe_allow_html=True,
            )

        # ── 4. Field preview ───────────────────────────────────────────────
        with st.expander("🔍 Preview all field mappings", expanded=False):
            preview_df = pd.DataFrame(list(field_map.items()), columns=["Field Name", "Value"])

            def highlight(row):
                if row["Field Name"] in missing:
                    return ["background-color:#fdecea"] * 2
                if row["Field Name"] in mandatory:
                    return [f"background-color:{light}"] * 2
                return [""] * 2

            st.dataframe(
                preview_df.style.apply(highlight, axis=1),
                use_container_width=True, hide_index=True,
            )

        # ── 5. Generate button ─────────────────────────────────────────────
        if missing:
            st.warning("⛔ Resolve missing mandatory fields before generating the PDF.")

        short_label = label.split("—")[0].strip()
        if st.button(
            f"🖨️ Generate {short_label} PDF",
            key=f"{tab_key}_btn",
            disabled=bool(missing),
            use_container_width=True,
        ):
            with st.spinner("Generating PDF… please wait."):
                with tempfile.TemporaryDirectory() as tmpdir:
                    temp_doc = os.path.join(tmpdir, "updated.docx")
                    temp_pdf = os.path.join(tmpdir, pdf_name)
                    doc = Document(word_file)
                    replace_placeholders(doc, field_map)
                    doc.save(temp_doc)
                    convert(temp_doc, temp_pdf)
                    with open(temp_pdf, "rb") as f:
                        st.download_button(
                            label=f"📥 Download {pdf_name}",
                            data=f,
                            file_name=pdf_name,
                            mime="application/pdf",
                            key=f"{tab_key}_download",
                        )
            st.markdown(
                f'<div class="info-box">✅ <b>{pdf_name}</b> generated successfully!</div>',
                unsafe_allow_html=True,
            )

    except Exception as e:
        st.error(f"❌ Unexpected error: {e}")


# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab_pos07, tab_3ds, tab_pos05 = st.tabs(["🏧 POS 07", "🔐 3DS", "💳 POS 05"])

with tab_pos07:
    render_tab("pos07")

with tab_3ds:
    render_tab("3ds")

with tab_pos05:
    render_tab("pos05")
