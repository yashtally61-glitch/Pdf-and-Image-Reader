import streamlit as st
import pandas as pd
import pytesseract
import cv2
import numpy as np
import re
import io
from PIL import Image, ImageEnhance, ImageFilter

st.set_page_config(page_title="DataScan → Excel", page_icon="⚡", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&display=swap');
    .stApp { background-color: #0a0a0f; }
    h1,h2,h3 { color: #00e5a0 !important; font-family: 'Space Mono', monospace; }
    .hero { background: linear-gradient(135deg, #13131a, #1a1a2e); border-left:4px solid #00e5a0; border-radius:12px; padding:1.5rem 2rem; margin-bottom:1.5rem; }
    .hero-title { font-family:'Space Mono',monospace; font-size:1.8rem; font-weight:700; color:#e8e8f0; }
    .hero-title span { color:#00e5a0; }
    .hero-sub { font-family:'Space Mono',monospace; color:#6b6b8a; font-size:0.8rem; margin-top:0.4rem; }
    .tip { background:rgba(0,229,160,0.08); border:1px solid rgba(0,229,160,0.2); border-radius:8px; padding:0.8rem 1rem; font-family:'Space Mono',monospace; font-size:0.78rem; color:#00e5a0; margin-bottom:1rem; }
    .metric-card { background:#13131a; border:1px solid #2a2a3a; border-radius:10px; padding:1rem; text-align:center; }
    .metric-val { font-family:'Space Mono',monospace; font-size:1.8rem; font-weight:700; color:#00e5a0; }
    .metric-label { font-family:'Space Mono',monospace; color:#6b6b8a; font-size:0.7rem; text-transform:uppercase; }
    div[data-testid="stDownloadButton"] button { background:#7b5ef8 !important; color:white !important; border:none !important; font-weight:700 !important; border-radius:8px !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <div class="hero-title">⚡ Data<span>Scan</span> → Excel</div>
  <div class="hero-sub">// No API · No Cost · Offline OCR · Ditto mark smart-fill</div>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Column Headers")
    col_input = st.text_area("Columns (comma-separated)", value="SKU,QTY,BIN", height=80)
    columns = [c.strip() for c in col_input.split(",") if c.strip()]

    st.markdown("---")
    st.markdown("### 🎯 Options")
    ditto_fill   = st.checkbox('Smart ditto fill (" → copy above)', value=True)
    show_preview = st.checkbox("Show processed image", value=False)

    st.markdown("---")
    st.markdown("""
    <div style='font-family:Space Mono,monospace;font-size:0.7rem;color:#6b6b8a;'>
    <b style='color:#00e5a0'>No API needed!</b><br>
    Uses Tesseract OCR.<br><br>
    <b style='color:#00e5a0'>Ditto marks filled:</b><br>
    <code>"</code> <code>,,</code> <code>//</code> <code>ditto</code>
    </div>
    """, unsafe_allow_html=True)


# ── Core Functions ─────────────────────────────────────────────────────────────

def prepare_image(pil_img: Image.Image) -> np.ndarray:
    """Aggressive preprocessing for handwritten ledger sheets."""
    # Convert and upscale
    img = pil_img.convert("RGB")
    w, h = img.size
    # Upscale small images
    if w < 1500:
        scale = 1500 / w
        img = img.resize((int(w*scale), int(h*scale)), Image.LANCZOS)

    # Enhance
    img = ImageEnhance.Contrast(img).enhance(2.0)
    img = ImageEnhance.Sharpness(img).enhance(2.0)
    img = ImageEnhance.Brightness(img).enhance(1.1)

    # Convert to numpy for OpenCV
    arr = np.array(img)
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)

    # Denoise
    gray = cv2.bilateralFilter(gray, 9, 75, 75)

    # Otsu threshold
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    # Morphological cleanup
    kernel = np.ones((1, 1), np.uint8)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)

    return thresh


def ocr_image(processed: np.ndarray) -> str:
    """Run Tesseract with optimal settings for ledger data."""
    # PSM 6 = assume single uniform block of text (good for tables)
    # PSM 4 = assume single column (also good)
    config = r'--oem 3 --psm 6 -c tessedit_char_whitelist="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-/. ,\"' + "'"
    text = pytesseract.image_to_string(processed, config=config)
    return text


def parse_to_rows(raw_text: str, columns: list[str]) -> list[dict]:
    """
    Parse OCR output into structured rows.
    Strategy: split each line into N parts matching column count.
    """
    rows = []
    lines = raw_text.splitlines()

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Skip lines that look like pure headers
        lower_line = line.lower().replace(" ", "")
        col_lower = [c.lower() for c in columns]
        if all(c in lower_line for c in col_lower):
            continue

        # Try splitting by 2+ spaces first
        parts = re.split(r'\s{2,}|\t', line)
        parts = [p.strip() for p in parts if p.strip()]

        # If we got too few parts, try single space split
        if len(parts) < len(columns):
            # For 3-column sheets: try to extract last known patterns
            # e.g. "1955-XL 2 T10-R11-A4" → ["1955-XL", "2", "T10-R11-A4"]
            parts = line.split()

        if not parts:
            continue

        row = {}
        # Map parts to columns as best as possible
        if len(parts) >= len(columns):
            # Assign first N-1 columns normally, rest goes to last column
            for i, col in enumerate(columns[:-1]):
                row[col] = parts[i] if i < len(parts) else ""
            # Last column gets everything remaining joined
            row[columns[-1]] = " ".join(parts[len(columns)-1:]) if len(parts) >= len(columns) else ""
        else:
            # Fewer parts than columns — fill what we can
            for i, col in enumerate(columns):
                row[col] = parts[i] if i < len(parts) else ""

        # Skip if first column is empty or looks like noise
        first_val = row.get(columns[0], "").strip()
        if not first_val or len(first_val) < 2:
            continue

        rows.append(row)

    return rows


def apply_ditto(df: pd.DataFrame) -> pd.DataFrame:
    """Replace ditto marks with value from the row above."""
    DITTO = re.compile(
        r'^(\s*["\'`]{1,3}\s*|,,|//|ditto|do|11|〃|\u3003)$',
        re.IGNORECASE
    )
    df = df.copy()
    for col in df.columns:
        for i in range(1, len(df)):
            val = str(df.at[i, col]).strip()
            if DITTO.match(val) or val in ('"', "''", "//", ",,", "``", '""'):
                df.at[i, col] = df.at[i - 1, col]
    return df


def to_excel(df: pd.DataFrame) -> bytes:
    """Export DataFrame to styled Excel bytes."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted Data")
        ws = writer.sheets["Extracted Data"]
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

        hfill  = PatternFill("solid", fgColor="1A1A2E")
        hfont  = Font(bold=True, color="00E5A0", name="Consolas")
        border = Border(
            bottom=Side(style="thin", color="2A2A3A"),
            right=Side(style="thin",  color="2A2A3A"),
        )

        for ci, cell in enumerate(ws[1], 1):
            cell.fill      = hfill
            cell.font      = hfont
            cell.alignment = Alignment(horizontal="center", vertical="center")
            max_len = max(
                (len(str(ws.cell(r, ci).value or "")) for r in range(1, ws.max_row + 1)),
                default=10,
            )
            ws.column_dimensions[cell.column_letter].width = min(max_len + 4, 40)

        for ri, row in enumerate(ws.iter_rows(min_row=2), 2):
            rf = PatternFill("solid", fgColor="0F0F18" if ri % 2 == 0 else "13131E")
            for cell in row:
                cell.fill      = rf
                cell.font      = Font(name="Consolas", color="E8E8F0")
                cell.border    = border
                cell.alignment = Alignment(vertical="center")

        ws.row_dimensions[1].height = 22
        ws.freeze_panes = "A2"

    buf.seek(0)
    return buf.read()


# ── UI Tabs ────────────────────────────────────────────────────────────────────

tab1, tab2 = st.tabs(["📤 Upload & Extract", "📊 Data & Export"])

with tab1:
    st.markdown(
        '<div class="tip">💡 <b>No API key needed.</b> '
        'OCR runs directly on the server. '
        'Ditto marks (<code>"</code>) are auto-filled from the row above.</div>',
        unsafe_allow_html=True,
    )

    uploaded_files = st.file_uploader(
        "Drop images here",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    if uploaded_files:
        col_go, _ = st.columns([1, 3])
        with col_go:
            go = st.button("⚡ Extract All", type="primary", use_container_width=True)

        if go:
            if not columns:
                st.error("⚠️ Add at least one column in the sidebar.")
            else:
                all_rows  = []
                progress  = st.progress(0, text="Starting OCR...")

                for idx, uf in enumerate(uploaded_files):
                    progress.progress(idx / len(uploaded_files), text=f"Processing {uf.name}…")
                    try:
                        img = Image.open(uf)

                        processed = prepare_image(img)

                        if show_preview:
                            st.image(processed, caption=f"Processed: {uf.name}", use_column_width=True)

                        raw_text = ocr_image(processed)

                        # Show raw OCR for debugging
                        with st.expander(f"📄 Raw OCR text — {uf.name}"):
                            st.text(raw_text)

                        rows = parse_to_rows(raw_text, columns)

                        for r in rows:
                            r["__source"] = uf.name

                        all_rows.extend(rows)
                        st.success(f"✅ `{uf.name}` → {len(rows)} rows found")

                    except Exception as e:
                        st.error(f"❌ `{uf.name}`: {e}")

                progress.progress(1.0, text="Done!")

                if all_rows:
                    df = pd.DataFrame(all_rows)
                    for c in columns:
                        if c not in df.columns:
                            df[c] = ""

                    source_col = df.pop("__source") if "__source" in df.columns else None
                    df = df[columns]

                    if ditto_fill:
                        df = apply_ditto(df)

                    if source_col is not None:
                        df["Source File"] = source_col.values

                    st.session_state["df"]    = df
                    st.session_state["ready"] = True
                    st.balloons()
                    st.info(f"✨ **{len(df)} rows** extracted — go to **Data & Export** tab.")
                else:
                    st.warning("⚠️ No rows were extracted. Try enabling 'Show processed image' to check OCR quality.")

with tab2:
    if st.session_state.get("ready") and "df" in st.session_state:
        df = st.session_state["df"]

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f'<div class="metric-card"><div class="metric-val">{len(df)}</div><div class="metric-label">Total Rows</div></div>', unsafe_allow_html=True)
        with c2:
            ne = df.iloc[:, 0].replace("", pd.NA).notna().sum()
            st.markdown(f'<div class="metric-card"><div class="metric-val">{ne}</div><div class="metric-label">Valid Rows</div></div>', unsafe_allow_html=True)
        with c3:
            src = df["Source File"].nunique() if "Source File" in df.columns else 1
            st.markdown(f'<div class="metric-card"><div class="metric-val">{src}</div><div class="metric-label">Source Files</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📋 Extracted Data *(click any cell to edit)*")

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=False,
        )
        st.session_state["df"] = edited_df

        st.markdown("---")

        # ── Download buttons ──────────────────────────────────────────────────
        d1, d2, d3 = st.columns([1, 1, 2])

        excel_data = to_excel(edited_df)

        with d1:
            st.download_button(
                label="⬇️ Download Excel (.xlsx)",
                data=excel_data,
                file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                label="⬇️ Download CSV",
                data=edited_df.to_csv(index=False).encode("utf-8"),
                file_name="extracted_data.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with d3:
            if st.button("🗑️ Clear Data", use_container_width=True):
                del st.session_state["df"]
                del st.session_state["ready"]
                st.rerun()

    else:
        st.markdown(
            '<div style="text-align:center;padding:4rem;color:#6b6b8a;'
            'font-family:Space Mono,monospace;">'
            '<div style="font-size:3rem">📊</div><br>'
            'No data yet — upload files and click <b style="color:#00e5a0">⚡ Extract All</b>'
            '</div>',
            unsafe_allow_html=True,
        )
