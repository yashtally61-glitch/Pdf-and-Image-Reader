import streamlit as st
import pandas as pd
import base64
import json
import re
import io
import requests
from PIL import Image, ImageEnhance

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
    .tip  { background:rgba(0,229,160,0.08); border:1px solid rgba(0,229,160,0.2); border-radius:8px; padding:0.8rem 1rem; font-family:'Space Mono',monospace; font-size:0.78rem; color:#00e5a0; margin-bottom:1rem; }
    .info { background:rgba(123,94,248,0.08); border:1px solid rgba(123,94,248,0.25); border-radius:8px; padding:0.8rem 1rem; font-family:'Space Mono',monospace; font-size:0.78rem; color:#b8a0ff; margin-bottom:1rem; }
    .metric-card { background:#13131a; border:1px solid #2a2a3a; border-radius:10px; padding:1rem; text-align:center; }
    .metric-val   { font-family:'Space Mono',monospace; font-size:1.8rem; font-weight:700; color:#00e5a0; }
    .metric-label { font-family:'Space Mono',monospace; color:#6b6b8a; font-size:0.7rem; text-transform:uppercase; }
    div[data-testid="stDownloadButton"] button { background:#7b5ef8 !important; color:white !important; border:none !important; font-weight:700 !important; border-radius:8px !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <div class="hero-title">⚡ Data<span>Scan</span> → Excel</div>
  <div class="hero-sub">// Free AI Vision · High Accuracy · No Credit Card · Ditto mark smart-fill</div>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuration")
    st.markdown("""
    <div class="info">
    🆓 <b>FREE — No Credit Card</b><br><br>
    1. Go to <a href="https://console.groq.com" target="_blank" style="color:#00e5a0"><b>console.groq.com</b></a><br>
    2. Sign up free with Google/GitHub<br>
    3. Click <b>API Keys → Create API Key</b><br>
    4. Paste below<br><br>
    ✅ Free tier · No card needed
    </div>
    """, unsafe_allow_html=True)

    default_key = ""
    try:
        default_key = st.secrets.get("GROQ_API_KEY", "")
    except Exception:
        pass

    api_key = st.text_input("Groq API Key", value=default_key, type="password", placeholder="gsk_...")

    st.markdown("---")
    st.markdown("### 📋 Column Headers")
    col_input = st.text_area("Columns (comma-separated)", value="SKU,QTY,BIN", height=80)
    columns = [c.strip() for c in col_input.split(",") if c.strip()]

    st.markdown("---")
    st.markdown("### 🎯 Options")
    ditto_fill = st.checkbox('Smart ditto fill (" → copy above)', value=True)
    enhance    = st.checkbox("Enhance image before extraction", value=True)

    st.markdown("---")
    st.markdown("""
    <div style='font-family:Space Mono,monospace;font-size:0.7rem;color:#6b6b8a;'>
    <b style='color:#00e5a0'>Ditto marks auto-filled:</b><br>
    <code>"</code> <code>,,</code> <code>//</code> <code>ditto</code> <code>11</code>
    </div>
    """, unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def enhance_image(pil_img: Image.Image) -> Image.Image:
    img = pil_img.convert("RGB")
    w, h = img.size
    if w < 1200:
        scale = 1200 / w
        img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
    img = ImageEnhance.Contrast(img).enhance(1.6)
    img = ImageEnhance.Sharpness(img).enhance(1.5)
    img = ImageEnhance.Brightness(img).enhance(1.05)
    return img


def image_to_base64(pil_img: Image.Image) -> str:
    buf = io.BytesIO()
    pil_img.save(buf, format="JPEG", quality=95)
    return base64.b64encode(buf.getvalue()).decode()


def build_prompt(columns: list[str]) -> str:
    col_list = ", ".join(columns)
    return f"""You are an expert OCR system for handwritten warehouse/inventory ledger sheets.

Extract EVERY data row from this image into a JSON array.

COLUMNS: {col_list}

DITTO MARK RULE (CRITICAL):
In handwritten sheets, ditto marks mean "same as the cell above in that column".
Ditto marks look like: " (double quote), '' two singles, // double slash, ,, two commas, the word ditto, tick marks like 11.
When you see a ditto mark → output the ACTUAL VALUE from the row above — NOT the symbol itself.

RULES:
1. Extract ALL rows — do not skip any row.
2. Each row = JSON object with keys exactly: {col_list}
3. Apply ditto rule to every column.
4. Blank/unreadable = ""
5. Keep original format: "1955-XL", "T10-R11-A4", "02"
6. Do NOT include the header row.
7. Be careful: 0 vs O, 1 vs I vs l, 5 vs S, 8 vs B
8. This sheet has 50+ rows — capture every single one.

Return ONLY a raw JSON array. No markdown, no explanation, no code fences.

Example:
[
  {{"SKU": "1955-XL", "QTY": "2", "BIN": "T10-R11-A4"}},
  {{"SKU": "1613-XL", "QTY": "1", "BIN": "T10-R11-A4"}}
]"""


def extract_with_groq(api_key: str, pil_img: Image.Image, columns: list[str]) -> list[dict]:
    b64 = image_to_base64(pil_img)

    payload = {
        "model": "meta-llama/llama-4-scout-17b-16e-instruct",
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{b64}"}
                    },
                    {
                        "type": "text",
                        "text": build_prompt(columns)
                    }
                ]
            }
        ],
        "max_tokens": 8192,
        "temperature": 0
    }

    resp = requests.post(
        "https://api.groq.com/openai/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        },
        json=payload,
        timeout=120
    )

    if resp.status_code != 200:
        raise ValueError(f"Groq API error {resp.status_code}: {resp.text[:300]}")

    raw = resp.json()["choices"][0]["message"]["content"].strip()
    raw = re.sub(r"```(?:json)?|```", "", raw).strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\[.*\]", raw, re.DOTALL)
        if m:
            data = json.loads(m.group())
        else:
            raise ValueError(f"Could not parse response:\n{raw[:400]}")

    return data if isinstance(data, list) else []


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
        hfill  = PatternFill("solid", fgColor="1A1A2E")
        hfont  = Font(bold=True, color="00E5A0", name="Consolas")
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
    buf.seek(0)
    return buf.read()


# ── UI ────────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📤 Upload & Extract", "📊 Data & Export"])

with tab1:
    st.markdown('<div class="tip">💡 Uses <b>Groq AI Vision</b> for near-perfect accuracy on handwritten sheets. Free — no credit card needed.</div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Drop images here",
        type=["jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

    if uploaded_files:
        col_go, _ = st.columns([1, 3])
        with col_go:
            go = st.button("⚡ Extract All", type="primary", use_container_width=True)

        if go:
            if not api_key:
                st.error("⚠️ Enter your **Groq API Key** in the sidebar. Get it FREE at console.groq.com")
            elif not columns:
                st.error("⚠️ Add at least one column.")
            else:
                all_rows = []
                progress = st.progress(0, text="Starting...")

                for idx, uf in enumerate(uploaded_files):
                    progress.progress(idx / len(uploaded_files), text=f"Processing {uf.name}…")
                    try:
                        img = Image.open(uf)
                        if enhance:
                            img = enhance_image(img)

                        rows = extract_with_groq(api_key, img, columns)
                        for r in rows:
                            r["__source"] = uf.name
                        all_rows.extend(rows)
                        st.success(f"✅ `{uf.name}` → {len(rows)} rows extracted")

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
                    st.info(f"✨ **{len(df)} rows** extracted — go to **Data & Export** tab.")
                else:
                    st.warning("No rows extracted. Check image quality and try again.")

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
            st.download_button("⬇️ Download Excel (.xlsx)", data=to_excel(edited_df),
                file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with d2:
            st.download_button("⬇️ Download CSV", data=edited_df.to_csv(index=False).encode(),
                file_name="extracted_data.csv", mime="text/csv", use_container_width=True)
        with d3:
            if st.button("🗑️ Clear Data", use_container_width=True):
                del st.session_state["df"]; del st.session_state["ready"]; st.rerun()
    else:
        st.markdown('<div style="text-align:center;padding:4rem;color:#6b6b8a;font-family:Space Mono,monospace;"><div style="font-size:3rem">📊</div><br>No data yet — upload and click <b style="color:#00e5a0">⚡ Extract All</b></div>', unsafe_allow_html=True)
