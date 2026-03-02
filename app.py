import streamlit as st
import pandas as pd
import anthropic
import base64
import json
import re
import io
from PIL import Image
from pathlib import Path

# ─── Page Config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DataScan → Excel",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&display=swap');

    .main { background-color: #0a0a0f; }
    .stApp { background-color: #0a0a0f; }

    h1, h2, h3 { color: #00e5a0 !important; font-family: 'Space Mono', monospace; }

    .hero-banner {
        background: linear-gradient(135deg, #13131a 0%, #1a1a2e 100%);
        border: 1px solid #2a2a3a;
        border-left: 4px solid #00e5a0;
        border-radius: 12px;
        padding: 1.5rem 2rem;
        margin-bottom: 1.5rem;
    }
    .hero-title {
        font-family: 'Space Mono', monospace;
        font-size: 1.8rem;
        font-weight: 700;
        color: #e8e8f0;
        margin: 0;
    }
    .hero-title span { color: #00e5a0; }
    .hero-sub {
        font-family: 'Space Mono', monospace;
        color: #6b6b8a;
        font-size: 0.8rem;
        margin-top: 0.4rem;
    }

    .metric-card {
        background: #13131a;
        border: 1px solid #2a2a3a;
        border-radius: 10px;
        padding: 1rem 1.2rem;
        text-align: center;
    }
    .metric-val {
        font-family: 'Space Mono', monospace;
        font-size: 1.8rem;
        font-weight: 700;
        color: #00e5a0;
    }
    .metric-label {
        font-family: 'Space Mono', monospace;
        color: #6b6b8a;
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }

    .tip-box {
        background: rgba(0,229,160,0.08);
        border: 1px solid rgba(0,229,160,0.2);
        border-radius: 8px;
        padding: 0.8rem 1rem;
        font-family: 'Space Mono', monospace;
        font-size: 0.78rem;
        color: #00e5a0;
        margin-bottom: 1rem;
    }

    .stDataFrame { border: 1px solid #2a2a3a; border-radius: 8px; }
    div[data-testid="stDownloadButton"] button {
        background: #7b5ef8;
        color: white;
        border: none;
        font-weight: 700;
        padding: 0.6rem 1.5rem;
        border-radius: 8px;
    }
    div[data-testid="stDownloadButton"] button:hover { background: #9470ff; }
</style>
""", unsafe_allow_html=True)

# ─── Header ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-banner">
  <div class="hero-title">⚡ Data<span>Scan</span> → Excel</div>
  <div class="hero-sub">// AI-powered image & PDF data extractor · High accuracy · Ditto mark smart-fill</div>
</div>
""", unsafe_allow_html=True)

# ─── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")

    api_key = st.text_input(
        "Anthropic API Key",
        type="password",
        help="Get your key at console.anthropic.com",
        placeholder="sk-ant-...",
    )

    st.markdown("---")
    st.markdown("### 📋 Column Headers")
    default_cols = "SKU,QTY,BIN"
    col_input = st.text_area(
        "Columns (comma-separated)",
        value=default_cols,
        height=80,
        help="Enter the column names matching your document",
    )
    columns = [c.strip() for c in col_input.split(",") if c.strip()]

    st.markdown("---")
    st.markdown("### 🎯 Extraction Options")

    ditto_handling = st.checkbox(
        "Smart ditto fill (\" → copy above)",
        value=True,
        help='If a cell contains " or ditto marks, copy the value from the cell above',
    )

    enhance_image = st.checkbox(
        "Pre-process images for clarity",
        value=True,
        help="Auto-enhance contrast before sending to AI",
    )

    st.markdown("---")
    st.markdown("""
    <div style='font-family:Space Mono,monospace;font-size:0.7rem;color:#6b6b8a;'>
    <b style='color:#00e5a0'>Ditto Mark Logic:</b><br>
    When the BIN column (or any column) shows <code>"</code>, <code>,,</code>, <code>ditto</code>, or
    <code>//</code> — the app automatically copies the last real value above it.
    </div>
    """, unsafe_allow_html=True)

# ─── Helpers ─────────────────────────────────────────────────────────────────

def encode_image(uploaded_file, enhance=True) -> tuple[str, str]:
    """Returns (base64_data, media_type). Optionally enhances image."""
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)

    # Detect type
    fname = uploaded_file.name.lower()
    if fname.endswith(".pdf"):
        return base64.b64encode(file_bytes).decode(), "application/pdf"

    if enhance:
        try:
            img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
            # Increase contrast slightly for handwritten docs
            from PIL import ImageEnhance, ImageFilter
            img = ImageEnhance.Contrast(img).enhance(1.4)
            img = ImageEnhance.Sharpness(img).enhance(1.3)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=95)
            file_bytes = buf.getvalue()
        except Exception:
            pass

    media_type = "image/jpeg"
    if fname.endswith(".png"):
        media_type = "image/png"
    elif fname.endswith(".webp"):
        media_type = "image/webp"

    return base64.b64encode(file_bytes).decode(), media_type


def build_extraction_prompt(columns: list[str]) -> str:
    col_list = ", ".join(columns)
    ditto_explain = """
DITTO MARK RULE (CRITICAL):
In handwritten ledgers/sheets, a ditto mark means "same as above".
Ditto marks can look like: " (double quote), '' (two single quotes), // (double slash), ,, (two commas), the word "ditto", or repeated tick marks (11, 11, 11).
When you see these in a cell, output the EXACT SAME VALUE as the cell directly above in the same column.
Do NOT output the ditto symbol itself — output the actual repeated value.
"""

    return f"""You are an expert OCR and data extraction system specialising in handwritten warehouse/inventory ledger sheets.

Your task: Extract EVERY data row from this document into a structured JSON array.

COLUMNS TO EXTRACT: {col_list}

{ditto_explain}

EXTRACTION RULES:
1. Extract ALL rows — do not skip any, even if partially legible.
2. For each row, output an object with keys exactly matching: {col_list}
3. Apply the ditto mark rule above for ALL columns.
4. If a value is genuinely missing/blank, use "".
5. Preserve original formatting (e.g. "1955-XL", "T10-R11-A4").
6. Numbers: preserve as strings (e.g. "02", not 2).
7. Do NOT include the header row in output.
8. Be very careful with handwritten characters: 0 vs O, 1 vs I vs l, 5 vs S, etc.
9. Read every single row — handwritten sheets may have 50+ rows, capture all.

OUTPUT FORMAT:
Return ONLY a valid JSON array, nothing else. No markdown, no explanation, no code fences.

Example output format:
[
  {{"SKU": "1955-XL", "QTY": "2", "BIN": "T10-R11-A4"}},
  {{"SKU": "1613-XL", "QTY": "1", "BIN": "T10-R11-A4"}},
  {{"SKU": "AK-103-XL", "QTY": "2", "BIN": "T10-R11-A4"}}
]
"""


def apply_ditto_fill(df: pd.DataFrame) -> pd.DataFrame:
    """Post-process: fill ditto marks with value from above row."""
    DITTO_PATTERNS = re.compile(
        r"""^(\s*["''`]{1,3}\s*|,,|//|ditto|do|〃|\"\"|'{2}|`{2}|11|„)$""",
        re.IGNORECASE
    )
    df = df.copy()
    for col in df.columns:
        for i in range(1, len(df)):
            val = str(df.at[i, col]).strip()
            if DITTO_PATTERNS.match(val) or val in ('"', "''", "//", ",,", "``"):
                # Copy from above
                df.at[i, col] = df.at[i - 1, col]
    return df


def extract_from_file(client, uploaded_file, columns, enhance) -> list[dict]:
    b64, media_type = encode_image(uploaded_file, enhance)
    is_pdf = media_type == "application/pdf"

    content_block = (
        {"type": "document", "source": {"type": "base64", "media_type": media_type, "data": b64}}
        if is_pdf
        else {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": b64}}
    )

    response = client.messages.create(
        model="claude-opus-4-5",          # Use Opus for highest accuracy
        max_tokens=8192,
        temperature=0,                     # Deterministic for data extraction
        system="You are a precision OCR and data extraction engine. Extract data with 100% accuracy. Never skip rows. Apply ditto mark rules exactly as instructed.",
        messages=[
            {
                "role": "user",
                "content": [
                    content_block,
                    {"type": "text", "text": build_extraction_prompt(columns)},
                ],
            }
        ],
    )

    raw = "".join(b.text for b in response.content if hasattr(b, "text"))
    # Strip any accidental markdown
    raw = re.sub(r"```(?:json)?|```", "", raw).strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        # Try to salvage partial JSON
        match = re.search(r"\[.*\]", raw, re.DOTALL)
        if match:
            data = json.loads(match.group())
        else:
            raise ValueError(f"Could not parse JSON from response:\n{raw[:500]}")

    return data if isinstance(data, list) else []


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted Data")
        ws = writer.sheets["Extracted Data"]

        # Style header
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        header_fill = PatternFill("solid", fgColor="1A1A2E")
        accent_fill = PatternFill("solid", fgColor="00E5A0")
        header_font = Font(bold=True, color="00E5A0", name="Consolas")
        border = Border(
            bottom=Side(style="thin", color="2A2A3A"),
            right=Side(style="thin", color="2A2A3A"),
        )

        for col_idx, cell in enumerate(ws[1], 1):
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            # Auto column width
            max_len = max(
                (len(str(ws.cell(r, col_idx).value or "")) for r in range(1, ws.max_row + 1)),
                default=10,
            )
            ws.column_dimensions[cell.column_letter].width = min(max_len + 4, 40)

        # Zebra striping
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), 2):
            fill_color = "0F0F18" if row_idx % 2 == 0 else "13131E"
            row_fill = PatternFill("solid", fgColor=fill_color)
            for cell in row:
                cell.fill = row_fill
                cell.font = Font(name="Consolas", color="E8E8F0")
                cell.border = border
                cell.alignment = Alignment(vertical="center")

        ws.row_dimensions[1].height = 22
        ws.freeze_panes = "A2"

    return buf.getvalue()


# ─── Main UI ─────────────────────────────────────────────────────────────────

tab1, tab2 = st.tabs(["📤 Upload & Extract", "📊 Data & Export"])

with tab1:
    st.markdown("""
    <div class="tip-box">
    💡 <b>Ditto Mark Support:</b> When the BIN column shows <code>"</code> or <code>,,</code> marks
    (meaning "same as above"), the app automatically fills the correct value from the row above.
    </div>
    """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Drop images or PDFs here",
        type=["jpg", "jpeg", "png", "webp", "pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if uploaded_files:
        st.markdown(f"**{len(uploaded_files)} file(s) loaded:**")
        for f in uploaded_files:
            size_kb = len(f.read()) / 1024
            f.seek(0)
            icon = "📄" if f.name.endswith(".pdf") else "🖼️"
            st.markdown(f"- {icon} `{f.name}` · {size_kb:.1f} KB")

        col_go, col_info = st.columns([1, 3])
        with col_go:
            go = st.button("⚡ Extract All", type="primary", use_container_width=True)

        if go:
            if not api_key:
                st.error("⚠️ Please enter your Anthropic API key in the sidebar.")
            elif not columns:
                st.error("⚠️ Please define at least one column in the sidebar.")
            else:
                client = anthropic.Anthropic(api_key=api_key)
                all_rows = []
                progress = st.progress(0, text="Starting extraction...")
                errors = []

                for i, uf in enumerate(uploaded_files):
                    progress.progress((i) / len(uploaded_files), text=f"Processing {uf.name}...")
                    try:
                        rows = extract_from_file(client, uf, columns, enhance_image)
                        for r in rows:
                            r["__source"] = uf.name
                        all_rows.extend(rows)
                        st.success(f"✅ `{uf.name}` → {len(rows)} rows extracted")
                    except Exception as e:
                        errors.append((uf.name, str(e)))
                        st.error(f"❌ `{uf.name}`: {e}")

                progress.progress(1.0, text="Done!")

                if all_rows:
                    df = pd.DataFrame(all_rows)
                    # Ensure all expected columns exist
                    for c in columns:
                        if c not in df.columns:
                            df[c] = ""
                    source_col = df.pop("__source") if "__source" in df.columns else None
                    df = df[columns]  # reorder

                    # Apply ditto fill
                    if ditto_handling:
                        df = apply_ditto_fill(df)

                    if source_col is not None:
                        df["Source File"] = source_col.values

                    st.session_state["df"] = df
                    st.session_state["ready"] = True
                    st.balloons()
                    st.info(f"✨ Total: **{len(df)} rows** extracted. Go to the **Data & Export** tab.")

with tab2:
    if st.session_state.get("ready") and "df" in st.session_state:
        df = st.session_state["df"]

        # Metrics row
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f'<div class="metric-card"><div class="metric-val">{len(df)}</div><div class="metric-label">Total Rows</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="metric-card"><div class="metric-val">{len(df.columns)}</div><div class="metric-label">Columns</div></div>', unsafe_allow_html=True)
        with c3:
            non_empty = df.iloc[:, 0].replace("", pd.NA).notna().sum()
            st.markdown(f'<div class="metric-card"><div class="metric-val">{non_empty}</div><div class="metric-label">Valid Rows</div></div>', unsafe_allow_html=True)
        with c4:
            sources = df["Source File"].nunique() if "Source File" in df.columns else 1
            st.markdown(f'<div class="metric-card"><div class="metric-val">{sources}</div><div class="metric-label">Source Files</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Editable table
        st.markdown("#### 📋 Extracted Data *(click any cell to edit)*")
        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=False,
        )
        st.session_state["df"] = edited_df

        st.markdown("---")
        col_dl1, col_dl2, col_dl3 = st.columns([1, 1, 2])
        with col_dl1:
            excel_bytes = to_excel_bytes(edited_df)
            st.download_button(
                label="⬇️ Download Excel (.xlsx)",
                data=excel_bytes,
                file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col_dl2:
            csv_data = edited_df.to_csv(index=False).encode()
            st.download_button(
                label="⬇️ Download CSV",
                data=csv_data,
                file_name="extracted_data.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with col_dl3:
            if st.button("🗑️ Clear Data", use_container_width=True):
                del st.session_state["df"]
                del st.session_state["ready"]
                st.rerun()

    else:
        st.markdown("""
        <div style='text-align:center;padding:4rem 2rem;color:#6b6b8a;font-family:Space Mono,monospace;'>
            <div style='font-size:3rem;margin-bottom:1rem'>📊</div>
            <div>No data yet. Upload files and click <b style='color:#00e5a0'>⚡ Extract All</b></div>
        </div>
        """, unsafe_allow_html=True)
