import streamlit as st
import pandas as pd
import google.generativeai as genai
import base64
import json
import re
import io
from PIL import Image

st.set_page_config(page_title="DataScan → Excel", page_icon="⚡", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&display=swap');
    .stApp { background-color: #0a0a0f; }
    h1,h2,h3 { color: #00e5a0 !important; font-family: 'Space Mono', monospace; }
    .hero-banner { background: linear-gradient(135deg, #13131a 0%, #1a1a2e 100%); border:1px solid #2a2a3a; border-left:4px solid #00e5a0; border-radius:12px; padding:1.5rem 2rem; margin-bottom:1.5rem; }
    .hero-title { font-family:'Space Mono',monospace; font-size:1.8rem; font-weight:700; color:#e8e8f0; margin:0; }
    .hero-title span { color:#00e5a0; }
    .hero-sub { font-family:'Space Mono',monospace; color:#6b6b8a; font-size:0.8rem; margin-top:0.4rem; }
    .free-badge { display:inline-block; background:linear-gradient(135deg,#00e5a0,#00c87a); color:#000; font-family:'Space Mono',monospace; font-size:0.7rem; font-weight:700; padding:0.2rem 0.7rem; border-radius:20px; margin-left:0.5rem; }
    .metric-card { background:#13131a; border:1px solid #2a2a3a; border-radius:10px; padding:1rem 1.2rem; text-align:center; }
    .metric-val { font-family:'Space Mono',monospace; font-size:1.8rem; font-weight:700; color:#00e5a0; }
    .metric-label { font-family:'Space Mono',monospace; color:#6b6b8a; font-size:0.72rem; text-transform:uppercase; letter-spacing:0.08em; }
    .tip-box { background:rgba(0,229,160,0.08); border:1px solid rgba(0,229,160,0.2); border-radius:8px; padding:0.8rem 1rem; font-family:'Space Mono',monospace; font-size:0.78rem; color:#00e5a0; margin-bottom:1rem; }
    .info-box { background:rgba(123,94,248,0.08); border:1px solid rgba(123,94,248,0.25); border-radius:8px; padding:0.8rem 1rem; font-family:'Space Mono',monospace; font-size:0.78rem; color:#b8a0ff; margin-bottom:1rem; }
    div[data-testid="stDownloadButton"] button { background:#7b5ef8; color:white; border:none; font-weight:700; padding:0.6rem 1.5rem; border-radius:8px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-banner">
  <div class="hero-title">⚡ Data<span>Scan</span> → Excel <span class="free-badge">FREE</span></div>
  <div class="hero-sub">// Powered by Google Gemini (free) · Image & PDF extractor · Ditto mark smart-fill</div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    st.markdown("""
    <div class="info-box">
    🆓 <b>FREE API — No Credit Card</b><br><br>
    1. Go to <a href="https://aistudio.google.com" target="_blank" style="color:#00e5a0"><b>aistudio.google.com</b></a><br>
    2. Sign in with Google<br>
    3. Click <b>Get API Key</b> → <b>Create API Key</b><br>
    4. Paste it below<br><br>
    <small>✅ 1,500 free requests/day</small>
    </div>
    """, unsafe_allow_html=True)

    default_key = ""
    try:
        default_key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        pass

    api_key = st.text_input("Google Gemini API Key", value=default_key, type="password", placeholder="AIzaSy...")

    st.markdown("---")
    st.markdown("### 📋 Column Headers")
    col_input = st.text_area("Columns (comma-separated)", value="SKU,QTY,BIN", height=80)
    columns = [c.strip() for c in col_input.split(",") if c.strip()]

    st.markdown("---")
    st.markdown("### 🎯 Options")
    ditto_handling = st.checkbox('Smart ditto fill (" → copy above)', value=True)
    enhance_image = st.checkbox("Pre-process images for clarity", value=True)

    st.markdown("---")
    st.markdown('<div style="font-family:Space Mono,monospace;font-size:0.7rem;color:#6b6b8a;"><b style="color:#00e5a0">Ditto Marks Detected:</b><br><code>"</code> <code>,,</code> <code>//</code> <code>ditto</code> <code>11</code><br>All filled automatically from row above.</div>', unsafe_allow_html=True)


def preprocess_image(file_bytes, enhance):
    if not enhance:
        return file_bytes
    try:
        from PIL import ImageEnhance
        img = Image.open(io.BytesIO(file_bytes)).convert("RGB")
        img = ImageEnhance.Contrast(img).enhance(1.5)
        img = ImageEnhance.Sharpness(img).enhance(1.4)
        img = ImageEnhance.Brightness(img).enhance(1.05)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=97)
        return buf.getvalue()
    except Exception:
        return file_bytes


def build_prompt(columns):
    col_list = ", ".join(columns)
    return f"""You are an expert OCR system for handwritten warehouse/inventory ledger sheets.

Extract EVERY data row into a JSON array.

COLUMNS: {col_list}

DITTO MARK RULE (CRITICAL):
Ditto marks mean "same as the cell above". They look like: " (double quote), '' (two singles), // (double slash), ,, (two commas), the word ditto, or tick marks like 11.
When you see a ditto mark → output the ACTUAL VALUE from the row above, NOT the symbol.

RULES:
1. Extract ALL rows — do not skip any.
2. Each row = JSON object with keys: {col_list}
3. Apply ditto rule for every column.
4. Blank/unreadable cell = ""
5. Keep original format: "1955-XL", "T10-R11-A4", "02"
6. Do NOT include the header row.
7. Distinguish: 0 vs O, 1 vs I vs l, 5 vs S, 8 vs B.
8. Sheets often have 50+ rows — capture every single one.

Return ONLY a raw JSON array. No markdown, no explanation, no code fences.

Example:
[
  {{"SKU": "1955-XL", "QTY": "2", "BIN": "T10-R11-A4"}},
  {{"SKU": "1613-XL", "QTY": "1", "BIN": "T10-R11-A4"}}
]"""


def extract_from_file(api_key, uploaded_file, columns, enhance):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(
        model_name="gemini-1.5-pro",
        generation_config=genai.GenerationConfig(temperature=0.0, max_output_tokens=8192)
    )

    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    fname = uploaded_file.name.lower()

    if fname.endswith(".pdf"):
        part = {"inline_data": {"mime_type": "application/pdf", "data": base64.b64encode(file_bytes).decode()}}
    else:
        file_bytes = preprocess_image(file_bytes, enhance)
        mime = "image/png" if fname.endswith(".png") else ("image/webp" if fname.endswith(".webp") else "image/jpeg")
        part = {"inline_data": {"mime_type": mime, "data": base64.b64encode(file_bytes).decode()}}

    response = model.generate_content([part, build_prompt(columns)])
    raw = re.sub(r"```(?:json)?|```", "", response.text.strip()).strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\[.*\]", raw, re.DOTALL)
        if m:
            data = json.loads(m.group())
        else:
            raise ValueError(f"Could not parse response:\n{raw[:500]}")

    return data if isinstance(data, list) else []


def apply_ditto_fill(df):
    DITTO = re.compile(r'^(\s*["\'`]{1,3}\s*|,,|//|ditto|do|11|,,)$', re.IGNORECASE)
    df = df.copy()
    for col in df.columns:
        for i in range(1, len(df)):
            val = str(df.at[i, col]).strip()
            if DITTO.match(val) or val in ('"', "''", "//", ",,", "``"):
                df.at[i, col] = df.at[i - 1, col]
    return df


def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted Data")
        ws = writer.sheets["Extracted Data"]
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        hf = PatternFill("solid", fgColor="1A1A2E")
        hfont = Font(bold=True, color="00E5A0", name="Consolas")
        border = Border(bottom=Side(style="thin", color="2A2A3A"), right=Side(style="thin", color="2A2A3A"))
        for ci, cell in enumerate(ws[1], 1):
            cell.fill = hf; cell.font = hfont
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


tab1, tab2 = st.tabs(["📤 Upload & Extract", "📊 Data & Export"])

with tab1:
    st.markdown('<div class="tip-box">💡 <b>Ditto Mark Support:</b> When BIN shows <code>"</code> or <code>,,</code> (same as above), the app fills the correct value automatically.</div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader("Drop images or PDFs here", type=["jpg","jpeg","png","webp","pdf"], accept_multiple_files=True, label_visibility="collapsed")

    if uploaded_files:
        st.markdown(f"**{len(uploaded_files)} file(s) loaded:**")
        for f in uploaded_files:
            size_kb = len(f.read()) / 1024; f.seek(0)
            st.markdown(f"- {'📄' if f.name.lower().endswith('.pdf') else '🖼️'} `{f.name}` · {size_kb:.1f} KB")

        col_go, _ = st.columns([1, 3])
        with col_go:
            go = st.button("⚡ Extract All", type="primary", use_container_width=True)

        if go:
            if not api_key:
                st.error("⚠️ Enter your **Google Gemini API Key** in the sidebar. Get it FREE at aistudio.google.com")
            elif not columns:
                st.error("⚠️ Add at least one column in the sidebar.")
            else:
                all_rows = []
                progress = st.progress(0, text="Starting...")
                for i, uf in enumerate(uploaded_files):
                    progress.progress(i / len(uploaded_files), text=f"Processing {uf.name}...")
                    try:
                        rows = extract_from_file(api_key, uf, columns, enhance_image)
                        for r in rows: r["__source"] = uf.name
                        all_rows.extend(rows)
                        st.success(f"✅ `{uf.name}` → {len(rows)} rows")
                    except Exception as e:
                        st.error(f"❌ `{uf.name}`: {e}")
                progress.progress(1.0, text="Done!")

                if all_rows:
                    df = pd.DataFrame(all_rows)
                    for c in columns:
                        if c not in df.columns: df[c] = ""
                    source_col = df.pop("__source") if "__source" in df.columns else None
                    df = df[columns]
                    if ditto_handling: df = apply_ditto_fill(df)
                    if source_col is not None: df["Source File"] = source_col.values
                    st.session_state["df"] = df
                    st.session_state["ready"] = True
                    st.balloons()
                    st.info(f"✨ **{len(df)} rows** extracted. Go to **Data & Export** tab.")

with tab2:
    if st.session_state.get("ready") and "df" in st.session_state:
        df = st.session_state["df"]
        c1,c2,c3,c4 = st.columns(4)
        with c1: st.markdown(f'<div class="metric-card"><div class="metric-val">{len(df)}</div><div class="metric-label">Total Rows</div></div>', unsafe_allow_html=True)
        with c2: st.markdown(f'<div class="metric-card"><div class="metric-val">{len(df.columns)}</div><div class="metric-label">Columns</div></div>', unsafe_allow_html=True)
        with c3:
            ne = df.iloc[:,0].replace("",pd.NA).notna().sum()
            st.markdown(f'<div class="metric-card"><div class="metric-val">{ne}</div><div class="metric-label">Valid Rows</div></div>', unsafe_allow_html=True)
        with c4:
            src = df["Source File"].nunique() if "Source File" in df.columns else 1
            st.markdown(f'<div class="metric-card"><div class="metric-val">{src}</div><div class="metric-label">Source Files</div></div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📋 Extracted Data *(click any cell to edit)*")
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic", hide_index=False)
        st.session_state["df"] = edited_df
        st.markdown("---")
        d1,d2,d3 = st.columns([1,1,2])
        with d1:
            st.download_button("⬇️ Download Excel (.xlsx)", data=to_excel_bytes(edited_df), file_name="extracted_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with d2:
            st.download_button("⬇️ Download CSV", data=edited_df.to_csv(index=False).encode(), file_name="extracted_data.csv", mime="text/csv", use_container_width=True)
        with d3:
            if st.button("🗑️ Clear Data", use_container_width=True):
                del st.session_state["df"]; del st.session_state["ready"]; st.rerun()
    else:
        st.markdown('<div style="text-align:center;padding:4rem 2rem;color:#6b6b8a;font-family:Space Mono,monospace;"><div style="font-size:3rem;margin-bottom:1rem">📊</div><div>No data yet. Upload files and click <b style="color:#00e5a0">⚡ Extract All</b></div></div>', unsafe_allow_html=True)
