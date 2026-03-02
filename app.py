import streamlit as st
import pandas as pd
import pytesseract
import cv2
import numpy as np
import re
import io
from PIL import Image
import pdf2image

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
    div[data-testid="stDownloadButton"] button { background:#7b5ef8; color:white; border:none; font-weight:700; border-radius:8px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <div class="hero-title">⚡ Data<span>Scan</span> → Excel</div>
  <div class="hero-sub">// No API · No Cost · 100% Offline OCR · Ditto mark smart-fill</div>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📋 Column Headers")
    col_input = st.text_area("Columns (comma-separated)", value="SKU,QTY,BIN", height=80)
    columns = [c.strip() for c in col_input.split(",") if c.strip()]

    st.markdown("---")
    st.markdown("### 🎯 Options")
    ditto_fill   = st.checkbox('Smart ditto fill (" → copy above)', value=True)
    enhance      = st.checkbox("Enhance image before OCR", value=True)
    show_preview = st.checkbox("Show processed image preview", value=False)

    st.markdown("---")
    st.markdown("""
    <div style='font-family:Space Mono,monospace;font-size:0.7rem;color:#6b6b8a;'>
    <b style='color:#00e5a0'>No API needed!</b><br>
    Uses Tesseract OCR locally.<br><br>
    <b style='color:#00e5a0'>Ditto marks auto-filled:</b><br>
    <code>"</code> <code>,,</code> <code>//</code> <code>ditto</code> <code>11</code>
    </div>
    """, unsafe_allow_html=True)


# ── Image Processing ──────────────────────────────────────────────────────────
def enhance_image(img_pil: Image.Image) -> np.ndarray:
    """Convert to grayscale, denoise, threshold for best OCR."""
    img = np.array(img_pil.convert("RGB"))
    gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
    # Upscale for better OCR on small text
    scale = 2.0
    gray = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    # Denoise
    gray = cv2.fastNlMeansDenoising(gray, h=10)
    # Adaptive threshold — works great for handwritten sheets
    thresh = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 31, 10
    )
    # Deskew
    coords = np.column_stack(np.where(thresh < 128))
    if len(coords) > 100:
        angle = cv2.minAreaRect(coords)[-1]
        if angle < -45: angle = 90 + angle
        if abs(angle) > 0.5:
            (h, w) = thresh.shape
            M = cv2.getRotationMatrix2D((w//2, h//2), angle, 1.0)
            thresh = cv2.warpAffine(thresh, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    return thresh


def run_ocr(img_array: np.ndarray) -> str:
    """Run Tesseract with best config for tabular data."""
    custom_config = r'--oem 3 --psm 6 -c preserve_interword_spaces=1'
    return pytesseract.image_to_string(img_array, config=custom_config)


def parse_rows(raw_text: str, columns: list[str]) -> list[dict]:
    """Parse OCR text into rows based on column count."""
    rows = []
    lines = [l.strip() for l in raw_text.splitlines() if l.strip()]

    # Skip header line if it matches column names
    skip_next = False
    for line in lines:
        lower = line.lower()
        # Skip lines that look like headers
        if any(c.lower() in lower for c in columns) and lower.count(' ') < 5:
            col_matches = sum(1 for c in columns if c.lower() in lower)
            if col_matches >= len(columns) - 1:
                continue

        # Split by 2+ spaces (column separator in OCR output)
        parts = re.split(r'\s{2,}|\t', line)
        parts = [p.strip() for p in parts if p.strip()]

        if not parts:
            continue

        row = {}
        for i, col in enumerate(columns):
            row[col] = parts[i] if i < len(parts) else ""

        # Skip rows where ALL fields are empty
        if all(v == "" for v in row.values()):
            continue

        rows.append(row)

    return rows


def apply_ditto(df: pd.DataFrame) -> pd.DataFrame:
    DITTO = re.compile(r'^(\s*["\'`]{1,3}\s*|,,|//|ditto|do|11|〃)$', re.IGNORECASE)
    df = df.copy()
    for col in df.columns:
        for i in range(1, len(df)):
            val = str(df.at[i, col]).strip()
            if DITTO.match(val) or val in ('"', "''", "//", ",,"):
                df.at[i, col] = df.at[i - 1, col]
    return df


def to_excel(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted Data")
        ws = writer.sheets["Extracted Data"]
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        hfill = PatternFill("solid", fgColor="1A1A2E")
        hfont = Font(bold=True, color="00E5A0", name="Consolas")
        border = Border(bottom=Side(style="thin", color="2A2A3A"), right=Side(style="thin", color="2A2A3A"))
        for ci, cell in enumerate(ws[1], 1):
            cell.fill = hfill; cell.font = hfont
            cell.alignment = Alignment(horizontal="center", vertical="center")
            max_len = max((len(str(ws.cell(r, ci).value or "")) for r in range(1, ws.max_row+1)), default=10)
            ws.column_dimensions[cell.column_letter].width = min(max_len+4, 40)
        for ri, row in enumerate(ws.iter_rows(min_row=2), 2):
            rf = PatternFill("solid", fgColor="0F0F18" if ri%2==0 else "13131E")
            for cell in row:
                cell.fill = rf; cell.font = Font(name="Consolas", color="E8E8F0")
                cell.border = border; cell.alignment = Alignment(vertical="center")
        ws.row_dimensions[1].height = 22
        ws.freeze_panes = "A2"
    return buf.getvalue()


# ── Main UI ───────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📤 Upload & Extract", "📊 Data & Export"])

with tab1:
    st.markdown('<div class="tip">💡 <b>No API key needed.</b> Upload your image or PDF — OCR runs directly on the server. Ditto marks (<code>"</code>) are auto-filled from the row above.</div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Drop images or PDFs",
        type=["jpg", "jpeg", "png", "webp", "pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

    if uploaded_files:
        col_go, _ = st.columns([1, 3])
        with col_go:
            go = st.button("⚡ Extract All", type="primary", use_container_width=True)

        if go:
            all_rows = []
            progress = st.progress(0, text="Starting OCR...")

            for idx, uf in enumerate(uploaded_files):
                progress.progress(idx / len(uploaded_files), text=f"OCR: {uf.name}...")
                try:
                    file_bytes = uf.read()
                    fname = uf.name.lower()

                    # Handle PDF → images
                    if fname.endswith(".pdf"):
                        pages = pdf2image.convert_from_bytes(file_bytes, dpi=300)
                        images = pages
                    else:
                        images = [Image.open(io.BytesIO(file_bytes))]

                    file_rows = []
                    for page_num, img in enumerate(images):
                        if enhance:
                            processed = enhance_image(img)
                        else:
                            processed = np.array(img.convert("L"))

                        if show_preview:
                            st.image(processed, caption=f"{uf.name} — page {page_num+1} (processed)", use_column_width=True)

                        raw_text = run_ocr(processed)
                        rows = parse_rows(raw_text, columns)
                        for r in rows:
                            r["__source"] = uf.name
                        file_rows.extend(rows)

                    all_rows.extend(file_rows)
                    st.success(f"✅ `{uf.name}` → {len(file_rows)} rows extracted")

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

                st.session_state["df"] = df
                st.session_state["ready"] = True
                st.balloons()
                st.info(f"✨ **{len(df)} rows** extracted. Go to **Data & Export** tab.")

with tab2:
    if st.session_state.get("ready") and "df" in st.session_state:
        df = st.session_state["df"]

        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="metric-card"><div class="metric-val">{len(df)}</div><div class="metric-label">Total Rows</div></div>', unsafe_allow_html=True)
        with c2:
            ne = df.iloc[:,0].replace("", pd.NA).notna().sum()
            st.markdown(f'<div class="metric-card"><div class="metric-val">{ne}</div><div class="metric-label">Valid Rows</div></div>', unsafe_allow_html=True)
        with c3:
            src = df["Source File"].nunique() if "Source File" in df.columns else 1
            st.markdown(f'<div class="metric-card"><div class="metric-val">{src}</div><div class="metric-label">Source Files</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📋 Extracted Data *(click any cell to edit)*")
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic", hide_index=False)
        st.session_state["df"] = edited_df

        st.markdown("---")
        d1, d2, d3 = st.columns([1, 1, 2])
        with d1:
            st.download_button("⬇️ Excel (.xlsx)", data=to_excel(edited_df), file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with d2:
            st.download_button("⬇️ CSV", data=edited_df.to_csv(index=False).encode(),
                file_name="extracted_data.csv", mime="text/csv", use_container_width=True)
        with d3:
            if st.button("🗑️ Clear", use_container_width=True):
                del st.session_state["df"]; del st.session_state["ready"]; st.rerun()
    else:
        st.markdown('<div style="text-align:center;padding:4rem;color:#6b6b8a;font-family:Space Mono,monospace;"><div style="font-size:3rem">📊</div><br>No data yet. Upload and click ⚡ Extract All</div>', unsafe_allow_html=True)
