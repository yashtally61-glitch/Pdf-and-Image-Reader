import streamlit as st
import pandas as pd
import base64
import json
import re
import io
import requests
from PIL import Image, ImageEnhance
from difflib import get_close_matches

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
    .warn { background:rgba(255,180,0,0.08); border:1px solid rgba(255,180,0,0.25); border-radius:8px; padding:0.8rem 1rem; font-family:'Space Mono',monospace; font-size:0.78rem; color:#ffcc44; margin-bottom:0.5rem; }
    .metric-card { background:#13131a; border:1px solid #2a2a3a; border-radius:10px; padding:1rem; text-align:center; }
    .metric-val   { font-family:'Space Mono',monospace; font-size:1.8rem; font-weight:700; color:#00e5a0; }
    .metric-label { font-family:'Space Mono',monospace; color:#6b6b8a; font-size:0.7rem; text-transform:uppercase; }
    div[data-testid="stDownloadButton"] button { background:#7b5ef8 !important; color:white !important; border:none !important; font-weight:700 !important; border-radius:8px !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <div class="hero-title">⚡ Data<span>Scan</span> → Excel</div>
  <div class="hero-sub">// AI Vision · SKU Validation · Ditto Mark Smart-Fill · Free</div>
</div>
""", unsafe_allow_html=True)

# ── Load Master SKU List ──────────────────────────────────────────────────────
@st.cache_resource
def load_master_skus():
    """Load master SKU list from uploaded Test.xlsx or embedded JSON."""
    try:
        # Try to load from secrets path first
        import os
        paths = [
            "Test.xlsx",
            "data/Test.xlsx",
            "/app/Test.xlsx",
        ]
        for p in paths:
            if os.path.exists(p):
                wb = openpyxl.load_workbook(p)
                ws = wb.active
                skus = [str(r[0]).strip() for r in ws.iter_rows(min_row=4, values_only=True) if r[0] and str(r[0]).strip()]
                return set(skus)
    except Exception:
        pass
    return set()

# Build SKU set from uploaded file if available in session
def get_sku_set():
    if "master_skus" in st.session_state:
        return st.session_state["master_skus"]
    return set()


# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Groq API Key")
    st.markdown("""
    <div class="info">
    🆓 <b>FREE — No Credit Card</b><br><br>
    1. <a href="https://console.groq.com" target="_blank" style="color:#00e5a0"><b>console.groq.com</b></a><br>
    2. Sign up free → API Keys → Create<br>
    3. Paste below ↓
    </div>
    """, unsafe_allow_html=True)

    default_key = ""
    try:
        default_key = st.secrets.get("GROQ_API_KEY", "")
    except Exception:
        pass

    api_key = st.text_input("Groq API Key", value=default_key, type="password", placeholder="gsk_...")

    st.markdown("---")
    st.markdown("### 📋 Columns")
    col_input = st.text_area("Columns (comma-separated)", value="SKU,QTY,BIN", height=80)
    columns = [c.strip() for c in col_input.split(",") if c.strip()]

    st.markdown("---")
    st.markdown("### 📂 Master SKU File")
    sku_file = st.file_uploader("Upload Test.xlsx (master SKUs)", type=["xlsx"])
    if sku_file:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(sku_file)
            ws = wb.active
            skus = [str(r[0]).strip() for r in ws.iter_rows(min_row=4, values_only=True) if r[0] and str(r[0]).strip()]
            st.session_state["master_skus"] = set(skus)
            st.success(f"✅ {len(skus):,} SKUs loaded!")
        except Exception as e:
            st.error(f"Error: {e}")

    st.markdown("---")
    st.markdown("### 🎯 Options")
    ditto_fill   = st.checkbox('Smart ditto fill (" → copy above)', value=True)
    validate_sku = st.checkbox("Validate & fix SKUs against master list", value=True)
    enhance_img  = st.checkbox("Enhance image before extraction", value=True)

    st.markdown("---")
    st.markdown("""<div style='font-family:Space Mono,monospace;font-size:0.7rem;color:#6b6b8a;'>
    <b style='color:#00e5a0'>SKU Formats Supported:</b><br>
    <code>1001YKBEIGE-XL</code><br>
    <code>AK-103BLUE-XXL</code><br>
    <code>1003KDMUSTARD-11-12</code><br>
    <code>7001YKBLS-L-XL</code><br>
    <code>4001DRSRED-S</code><br>
    <code>6003SKDGREEN-XL</code><br>
    </div>""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────────────────────────────────

def enhance_image(pil_img: Image.Image) -> Image.Image:
    img = pil_img.convert("RGB")
    w, h = img.size
    if w < 1200:
        scale = 1200 / w
        img = img.resize((int(w*scale), int(h*scale)), Image.LANCZOS)
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
    return f"""You are an expert OCR system for handwritten warehouse/inventory ledger sheets from an Indian clothing brand.

Extract EVERY data row from this image into a JSON array.

COLUMNS: {col_list}

═══════════════════════════════
SKU FORMAT RULES (CRITICAL):
═══════════════════════════════
The SKU column follows these exact patterns — read VERY carefully:

PATTERN 1: {{NUMBER}}YK{{COLOR}}-{{SIZE}}
  Examples: 1001YKBEIGE-XL, 1057YKBLUE-3XL, 1189YKMAROON-8XL

PATTERN 2: {{NUMBER}}KD{{COLOR}}-{{SIZE}} (kids sizes like 7-8, 9-10, 11-12, 13-14)
  Examples: 1003KDMUSTARD-11-12, 1006KDBLUE-7-8

PATTERN 3: {{NUMBER}}YK{{NUMBER}}{{COLOR}}-{{SIZE}} (has YK + another number)
  Examples: 108YK148PINKRAY-L, 119YK148YELLOW-XL, 182YK305MUSTARD-XXL

PATTERN 4: AK-{{NUMBER}}{{COLOR}}-{{SIZE}}
  Examples: AK-103BLUE-XL, AK-120BLACK-XXL, AK-141PINK-3XL

PATTERN 5: {{NUMBER}}YKBLS{{COLOR}}-{{SIZE}} (BLS = blouse)
  Examples: 7001YKBLS-L-XL, 7001YKBLS-S-M

PATTERN 6: {{NUMBER}}DRS{{COLOR}}-{{SIZE}}
  Examples: 4001DRSRED-S, 4006DRSRED-XXL

PATTERN 7: {{NUMBER}}SKD{{COLOR}}-{{SIZE}}
  Examples: 6003SKDGREEN-XL, 6004SKDRED-S

PATTERN 8: {{NUMBER}}MW{{COLOR}}-{{SIZE}}
  Examples: 8001MWRED-S, 8002MWGREEN-XXL

PATTERN 9: KD{{NUMBER}}{{COLOR}}-{{SIZE}}
  Examples: KD001SKYBLUE-7-8, KD0010PINK-11-12

PATTERN 10: {{NUMBER}}DPT{{NUMBER}}{{COLOR}}-{{SIZE}}
  Examples: 1379DPT22MAROON-S, 1338DPT9BLACK-XL

PATTERN 11: TB{{NUMBER}}YK{{COLOR}} (no size suffix)
  Examples: TB1YKLAVENDER, TB9YKPINK

EXACT COLOR NAMES — never abbreviate or change:
BEIGE, BLUE, BLACK, BROWN, CREAM, DARKGREY, DENIM, FIROZI, GREEN, GREY,
INDIGO, KHAKI, LAVENDER, LEMON, MAROON, MEHROON, MINT, MULTI, MUSTARD,
MUSTRAD, NAVY, NAVYBLUE, OFFWHITE, OWHITE, OLIVE, ORANGE, PEACH, PINK,
PINKRAY, PISTAGREEN, POWDERBLUE, PURPLE, RANI, RED, RUST, SEAGREEN,
SHARARA, SKYBLUE, TEAL, TURQ, TYEDYE, WHITE, WINE, YELLOW,
AZKBLUE, AZKBLU, BGDY, BBNDJ, BNDJ, CHECKS, AJARAKH, LEHRIYA, BOTTELGREEN,
CRYSTALTEAL, HUNTERGREEN, DEEPGREEN, BUBBLEGUMPINK, HOTPINK, PASTELBLUE,
PISTAGREEN, DGREEN, DBLUE, DWHITE, DPINK, DMULTI, DMUSTARD, DLAVENDER,
DGREY, DYELLOW, PBLUE, SKYBLUE, LBLUE, LGREEN, OWHITE, MEHROON, MUAVE,
MAUVE, INDIGO, FIROZI, BLU

EXACT SIZE SUFFIXES:
- Standard: XS, S, M, L, XL, XXL, 3XL, 4XL, 5XL, 6XL, 7XL, 8XL
- Kids: 0-3, 3-6, 6-9, 7-8, 9-10, 11-12, 13-14, 7-14
- Combined: S-M, L-XL, XXL-3XL, 4XL-5XL, S/M, L/XL, XS/S, M/L
- Free size: F, F-S/XXL, F-3xl/5xl
- Other: L-XL, S-M-L, XL-XXL-3XL

IMPORTANT RULES:
1. Extract ALL rows — do not skip any
2. CASE SENSITIVE — preserve exact uppercase as shown above
3. Ditto marks (", '', //, ,,, ditto, 11) → replace with value from row above
4. Keep exact format: 1001YKBEIGE-XL not 1001-YK-BEIGE-XL
5. Never add spaces inside SKU
6. BIN column format like T10-R11-A4, keep exact
7. QTY is a number — read carefully (0 vs O, 1 vs l, 6 vs G)
8. If SKU is unclear, give best match based on patterns above

Return ONLY a raw JSON array. No markdown, no explanation.

Example output:
[
  {{"SKU": "1001YKBEIGE-XL", "QTY": "2", "BIN": "T10-R11-A4"}},
  {{"SKU": "1003KDMUSTARD-11-12", "QTY": "5", "BIN": "T10-R11-A4"}}
]"""


def extract_with_groq(api_key: str, pil_img: Image.Image, columns: list[str]) -> list[dict]:
    b64 = image_to_base64(pil_img)
    payload = {
        "model": "meta-llama/llama-4-scout-17b-16e-instruct",
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                {"type": "text", "text": build_prompt(columns)}
            ]
        }],
        "max_tokens": 8192,
        "temperature": 0
    }
    resp = requests.post(
        "https://api.groq.com/openai/v1/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json=payload, timeout=120
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
            raise ValueError(f"Could not parse JSON:\n{raw[:400]}")
    return data if isinstance(data, list) else []


SIZE_SUFFIX_RE = re.compile(
    r'-(XS|S|M|L|XL|XXL|[3-8]XL|F|F-S/XXL|F-3XL/5XL|\d+-\d+|S-M|L-XL|XXL-3XL|4XL-5XL|XS-S|M-L|S/M|L/XL|XS/S|M/L)$',
    re.IGNORECASE
)

def build_base_map(master_skus: set) -> dict:
    """Build uppercase-base → [full SKUs] mapping from master list."""
    base_map = {}
    for s in master_skus:
        base = SIZE_SUFFIX_RE.sub('', s).upper()
        base_map.setdefault(base, []).append(s)
    return base_map

def validate_and_fix_sku(sku: str, master_skus: set, base_map: dict) -> tuple[str, bool, str, list]:
    """
    Returns (corrected_sku, was_changed, suggestion, expanded_rows).
    - If SKU matches exactly → return as-is
    - If SKU is partial (no size suffix) → expand to ALL matching full SKUs from master
    - If SKU is wrong → fuzzy match to best full SKU
    expanded_rows: list of full SKUs to expand this row into (empty = no expansion)
    """
    sku = str(sku).strip()
    if not sku or sku == "nan":
        return sku, False, "", []

    # Exact match in master
    if sku in master_skus:
        return sku, False, "", []

    sku_upper = sku.upper()

    # Exact match after uppercase
    if sku_upper in master_skus:
        return sku_upper, True, sku_upper, []

    # Check if it's a PARTIAL SKU (base matches a known base → expand)
    base_of_input = SIZE_SUFFIX_RE.sub('', sku_upper)
    if base_of_input in base_map:
        full_skus = sorted(base_map[base_of_input])
        return full_skus[0], True, f"EXPANDED → {len(full_skus)} SKUs", full_skus

    # Fuzzy match against all bases (partial input fuzzy)
    all_bases = list(base_map.keys())
    base_matches = get_close_matches(base_of_input, all_bases, n=1, cutoff=0.82)
    if base_matches:
        full_skus = sorted(base_map[base_matches[0]])
        return full_skus[0], True, f"EXPANDED (fuzzy) → {len(full_skus)} SKUs", full_skus

    # Fuzzy match against full SKU list
    matches = get_close_matches(sku_upper, master_skus, n=1, cutoff=0.85)
    if matches:
        return matches[0], True, matches[0], []

    return sku, False, "", []  # Not found, keep as-is


def apply_ditto(df: pd.DataFrame) -> pd.DataFrame:
    DITTO = re.compile(r'^(\s*["\'`]{1,3}\s*|,,|//|ditto|do|11|〃)$', re.IGNORECASE)
    df = df.copy()
    for col in df.columns:
        for i in range(1, len(df)):
            val = str(df.at[i, col]).strip()
            if DITTO.match(val) or val in ('"', "''", "//", ",,"):
                df.at[i, col] = df.at[i-1, col]
    return df


def to_excel(df: pd.DataFrame, corrections: dict = None) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted Data")
        ws = writer.sheets["Extracted Data"]
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, PatternFill
        from openpyxl.styles.fills import PatternFill as PF

        hfill  = PatternFill("solid", fgColor="1A1A2E")
        hfont  = Font(bold=True, color="00E5A0", name="Consolas")
        border = Border(bottom=Side(style="thin", color="2A2A3A"), right=Side(style="thin", color="2A2A3A"))
        warn_fill = PatternFill("solid", fgColor="2A1A00")

        for ci, cell in enumerate(ws[1], 1):
            cell.fill = hfill; cell.font = hfont
            cell.alignment = Alignment(horizontal="center", vertical="center")
            max_len = max((len(str(ws.cell(r, ci).value or "")) for r in range(1, ws.max_row+1)), default=10)
            ws.column_dimensions[cell.column_letter].width = min(max_len+4, 45)

        sku_col_idx = None
        for ci, cell in enumerate(ws[1], 1):
            if str(cell.value).strip().upper() == "SKU":
                sku_col_idx = ci
                break

        for ri, row in enumerate(ws.iter_rows(min_row=2), 2):
            row_idx = ri - 2
            rf = PatternFill("solid", fgColor="0F0F18" if ri%2==0 else "13131E")
            for cell in row:
                cell.fill = rf
                cell.font = Font(name="Consolas", color="E8E8F0")
                cell.border = border
                cell.alignment = Alignment(vertical="center")

            # Highlight auto-corrected SKUs in amber
            if sku_col_idx and corrections and row_idx in corrections:
                ws.cell(ri, sku_col_idx).fill = PatternFill("solid", fgColor="3A2A00")
                ws.cell(ri, sku_col_idx).font = Font(name="Consolas", color="FFCC44")

        ws.row_dimensions[1].height = 22
        ws.freeze_panes = "A2"

    buf.seek(0)
    return buf.read()


# ── UI ────────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["📤 Upload & Extract", "📊 Data & Export"])

with tab1:
    master_skus = get_sku_set()
    if master_skus:
        st.markdown(f'<div class="tip">✅ Master SKU list loaded: <b>{len(master_skus):,} SKUs</b> — extracted SKUs will be validated & auto-corrected.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="warn">⚠️ No master SKU file loaded. Upload <b>Test.xlsx</b> in the sidebar to enable SKU validation.</div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader("Drop images here", type=["jpg","jpeg","png","webp"], accept_multiple_files=True, label_visibility="collapsed")

    if uploaded_files:
        col_go, _ = st.columns([1, 3])
        with col_go:
            go = st.button("⚡ Extract All", type="primary", use_container_width=True)

        if go:
            if not api_key:
                st.error("⚠️ Enter your Groq API Key in the sidebar.")
            elif not columns:
                st.error("⚠️ Add at least one column.")
            else:
                all_rows = []
                corrections = {}
                progress = st.progress(0, text="Starting...")

                for idx, uf in enumerate(uploaded_files):
                    progress.progress(idx / len(uploaded_files), text=f"Processing {uf.name}…")
                    try:
                        img = Image.open(uf)
                        if enhance_img:
                            img = enhance_image(img)

                        rows = extract_with_groq(api_key, img, columns)
                        for r in rows:
                            r["__source"] = uf.name
                        all_rows.extend(rows)
                        st.success(f"✅ `{uf.name}` → {len(rows)} rows")
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

                    # SKU Validation + Partial SKU Expansion
                    if validate_sku and master_skus and "SKU" in df.columns:
                        base_map = build_base_map(master_skus)
                        fixed_count = 0
                        expanded_count = 0
                        not_found = []
                        expanded_rows = []  # Collect all final rows after expansion

                        for i, row in df.iterrows():
                            sku = str(row.get("SKU", "")).strip()
                            if not sku:
                                expanded_rows.append(row.to_dict())
                                continue

                            corrected, changed, suggestion, full_skus = validate_and_fix_sku(sku, master_skus, base_map)

                            if full_skus:
                                # Expand this one row into multiple rows (one per size)
                                for full_sku in full_skus:
                                    new_row = row.to_dict()
                                    new_row["SKU"] = full_sku
                                    new_row["__expanded"] = True
                                    expanded_rows.append(new_row)
                                expanded_count += len(full_skus)
                                fixed_count += 1
                            elif changed:
                                r = row.to_dict()
                                r["SKU"] = corrected
                                r["__expanded"] = False
                                expanded_rows.append(r)
                                corrections[i] = corrected
                                fixed_count += 1
                            else:
                                r = row.to_dict()
                                r["__expanded"] = False
                                expanded_rows.append(r)
                                if corrected not in master_skus and corrected:
                                    not_found.append(sku)

                        # Rebuild df from expanded rows
                        df = pd.DataFrame(expanded_rows)
                        expanded_flag = df.pop("__expanded") if "__expanded" in df.columns else None
                        # Mark all expanded rows as corrections for amber highlight
                        corrections = {}
                        if expanded_flag is not None:
                            for idx_pos in range(len(df)):
                                if expanded_flag.iloc[idx_pos]:
                                    corrections[df.index[idx_pos]] = df.iloc[idx_pos]["SKU"]

                        if fixed_count:
                            st.warning(f"🔧 **{fixed_count} partial/wrong SKUs** matched → **{expanded_count} full SKU rows** generated (highlighted amber in Excel).")
                        if not_found:
                            with st.expander(f"⚠️ {len(not_found)} SKUs not found in master list"):
                                for s in not_found[:50]:
                                    st.code(s)

                    # Ensure columns exist and are ordered correctly after potential expansion
                    for c in columns:
                        if c not in df.columns:
                            df[c] = ""
                    extra_cols = [c for c in df.columns if c not in columns]
                    df = df[columns + extra_cols]

                    if source_col is not None and len(source_col) != len(df):
                        pass  # source col size mismatch after expansion — skip
                    elif source_col is not None:
                        df["Source File"] = source_col.values

                    st.session_state["df"] = df
                    st.session_state["corrections"] = corrections
                    st.session_state["ready"] = True
                    st.balloons()
                    st.info(f"✨ **{len(df)} rows** extracted — go to **Data & Export** tab.")
                else:
                    st.warning("No rows extracted. Check image quality.")

with tab2:
    if st.session_state.get("ready") and "df" in st.session_state:
        df = st.session_state["df"]
        corrections = st.session_state.get("corrections", {})

        c1, c2, c3, c4 = st.columns(4)
        with c1: st.markdown(f'<div class="metric-card"><div class="metric-val">{len(df)}</div><div class="metric-label">Total Rows</div></div>', unsafe_allow_html=True)
        with c2:
            valid = df.iloc[:,0].replace("", pd.NA).notna().sum()
            st.markdown(f'<div class="metric-card"><div class="metric-val">{valid}</div><div class="metric-label">Valid Rows</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="metric-card"><div class="metric-val">{len(corrections)}</div><div class="metric-label">SKUs Fixed</div></div>', unsafe_allow_html=True)
        with c4:
            src = df["Source File"].nunique() if "Source File" in df.columns else 1
            st.markdown(f'<div class="metric-card"><div class="metric-val">{src}</div><div class="metric-label">Source Files</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("#### 📋 Extracted Data *(click any cell to edit)*")
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic", hide_index=False)
        st.session_state["df"] = edited_df

        st.markdown("---")
        d1, d2, d3 = st.columns([1,1,2])
        with d1:
            st.download_button(
                "⬇️ Download Excel (.xlsx)",
                data=to_excel(edited_df, corrections),
                file_name="extracted_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with d2:
            st.download_button(
                "⬇️ Download CSV",
                data=edited_df.to_csv(index=False).encode(),
                file_name="extracted_data.csv",
                mime="text/csv",
                use_container_width=True
            )
        with d3:
            if st.button("🗑️ Clear Data", use_container_width=True):
                for k in ["df","ready","corrections"]: st.session_state.pop(k, None)
                st.rerun()
    else:
        st.markdown('<div style="text-align:center;padding:4rem;color:#6b6b8a;font-family:Space Mono,monospace;"><div style="font-size:3rem">📊</div><br>No data yet — upload images and click <b style="color:#00e5a0">⚡ Extract All</b></div>', unsafe_allow_html=True)
