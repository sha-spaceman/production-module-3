import streamlit as st
import pandas as pd
import io

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Examinent · File Parts Extractor",
    page_icon="📄",
    layout="centered",
)

# ── Styling ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=DM+Serif+Display&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* Top banner */
.banner {
    background: #0f2a4a;
    border-radius: 12px;
    padding: 2rem 2.2rem 1.6rem;
    margin-bottom: 2rem;
}
.banner h1 {
    font-family: 'DM Serif Display', serif;
    color: #ffffff;
    font-size: 1.9rem;
    margin: 0 0 0.35rem 0;
    letter-spacing: -0.01em;
}
.banner p {
    color: #8fb3d9;
    font-size: 0.92rem;
    margin: 0;
    line-height: 1.5;
}

/* Section labels */
.section-label {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #0f2a4a;
    margin-bottom: 0.4rem;
}

/* Board cards */
.board-hint {
    font-size: 0.82rem;
    color: #607a96;
    margin-top: 0.2rem;
}

/* Result stat strip */
.stat-strip {
    display: flex;
    gap: 1.2rem;
    margin: 1.4rem 0 1rem;
}
.stat-card {
    flex: 1;
    background: #f0f5fb;
    border: 1px solid #c9daea;
    border-radius: 8px;
    padding: 0.85rem 1rem;
    text-align: center;
}
.stat-card .num {
    font-family: 'DM Serif Display', serif;
    font-size: 1.8rem;
    color: #0f2a4a;
    line-height: 1;
}
.stat-card .lbl {
    font-size: 0.78rem;
    color: #607a96;
    margin-top: 0.25rem;
    font-weight: 500;
}

/* Download button override */
.stDownloadButton > button {
    background: #0f2a4a !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 7px !important;
    font-weight: 500 !important;
    padding: 0.55rem 1.4rem !important;
    width: 100%;
}
.stDownloadButton > button:hover {
    background: #1a3f6f !important;
}

/* Divider */
hr { border-color: #d6e4f0; margin: 1.6rem 0; }

/* Tab strip */
[data-baseweb="tab-list"] {
    gap: 0.5rem;
    border-bottom: 2px solid #d6e4f0;
}
[data-baseweb="tab"] {
    font-size: 0.85rem !important;
    font-weight: 500 !important;
}

/* Expander */
details summary {
    font-size: 0.85rem;
    color: #0f2a4a;
    font-weight: 500;
}

/* Footer */
.footer {
    text-align: center;
    font-size: 0.75rem;
    color: #aabfcf;
    margin-top: 3rem;
    padding-top: 1rem;
    border-top: 1px solid #e2edf5;
}
</style>
""", unsafe_allow_html=True)

# ── Header banner ─────────────────────────────────────────────────────────────
st.markdown("""
<div class="banner">
  <h1>File Parts Extractor</h1>
  <p>Upload your PNG filename list, choose an exam board, and download<br>
  the QParts and MSParts breakdown as a ready-to-use Excel file.</p>
</div>
""", unsafe_allow_html=True)


# ── Parsers ───────────────────────────────────────────────────────────────────
def parse_caie(name):
    q_paper  = name[:14]
    q_number = name[15:17]
    q_part   = name[18:19] if len(name) == 23 else ""
    return q_paper, q_number, q_part


def parse_edexcel(name):
    if len(name) > 7 and name[7] == "r":
        q_paper  = name[:21]
        q_number = name[22:24]
        len_short, len_long   = 29, 30
        pos_short, pos_long   = 24, 25
    else:
        q_paper  = name[:20]
        q_number = name[21:23]
        len_short, len_long   = 28, 29
        pos_short, pos_long   = 23, 24

    if len(name) == len_short:
        q_part = name[pos_short:pos_short + 1]
    elif len(name) == len_long:
        q_part = name[pos_long:pos_long + 1]
    else:
        q_part = ""
    return q_paper, q_number, q_part


def parse_ibdp(name):
    q_paper  = name[:29]
    q_number = name[30:32]
    q_part   = name[33:34] if len(name) == 38 else ""
    return q_paper, q_number, q_part


PARSERS = {
    "Cambridge (CAIE)":   parse_caie,
    "Pearson Edexcel":    parse_edexcel,
    "IBDP":               parse_ibdp,
}

QP_PATTERNS = {
    "Cambridge (CAIE)":  "*qp*",
    "Pearson Edexcel":   "*que*",
    "IBDP":              "*qp*",
}

MS_PATTERNS = {
    "Cambridge (CAIE)":  "*ms*",
    "Pearson Edexcel":   "*rms*",
    "IBDP":              "*ms*",
}


def like_filter(series, pattern):
    inner = pattern.strip("*").lower()
    s = series.str.lower()
    if pattern.startswith("*") and pattern.endswith("*"):
        return s.str.contains(inner, na=False)
    elif pattern.startswith("*"):
        return s.str.endswith(inner, na=False)
    elif pattern.endswith("*"):
        return s.str.startswith(inner, na=False)
    return s == inner


def run_query(df, pattern):
    mask     = like_filter(df["Q paper/Mark Scheme"], pattern)
    filtered = df[mask]
    return (
        filtered
        .groupby(["Q paper/Mark Scheme", "Q Number"], sort=False)["Q part"]
        .max()
        .reset_index()
        .rename(columns={"Q part": "MaxOfQ part"})
    )


def build_excel(qp_df, ms_df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        qp_df.to_excel(writer, sheet_name="QParts",  index=False)
        ms_df.to_excel(writer, sheet_name="MSParts", index=False)
    return buf.getvalue()


# ── Step 1 — Board selection ──────────────────────────────────────────────────
st.markdown('<div class="section-label">Step 1 — Examination Board</div>', unsafe_allow_html=True)

board = st.radio(
    label="board",
    options=list(PARSERS.keys()),
    label_visibility="collapsed",
    horizontal=False,
)

HINTS = {
    "Cambridge (CAIE)":  "Filenames follow the CAIE 14-character paper code convention.",
    "Pearson Edexcel":   "Handles both standard and revision-paper ('r') filename formats.",
    "IBDP":              "Filenames follow the 29-character IB paper code convention.",
}
st.markdown(f'<div class="board-hint">ℹ &nbsp;{HINTS[board]}</div>', unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

# ── Step 2 — File upload ──────────────────────────────────────────────────────
st.markdown('<div class="section-label">Step 2 — Upload Filename List</div>', unsafe_allow_html=True)
st.markdown(
    "<p style='font-size:0.83rem;color:#607a96;margin-bottom:0.6rem;'>"
    "Upload the Excel file containing a single column of PNG filenames (no header row).</p>",
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    label="Choose an Excel file",
    type=["xlsx", "xls"],
    label_visibility="collapsed",
)

# ── Step 3 — Process ──────────────────────────────────────────────────────────
if uploaded:
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="section-label">Step 3 — Results</div>', unsafe_allow_html=True)

    try:
        raw       = pd.read_excel(uploaded, header=None, dtype=str)
        filenames = raw.iloc[:, 0].dropna().str.strip().tolist()
        filenames = [f for f in filenames if f]
    except Exception as e:
        st.error(f"Could not read the file: {e}")
        st.stop()

    if not filenames:
        st.warning("The uploaded file appears to be empty. Please check it and try again.")
        st.stop()

    parser = PARSERS[board]
    records = []
    errors  = []
    for name in filenames:
        try:
            qp, qn, qt = parser(name)
            records.append({
                "Q paper/Mark Scheme": qp.strip(),
                "Q Number":            qn.strip(),
                "Q part":              qt.strip(),
            })
        except Exception:
            errors.append(name)

    if errors:
        with st.expander(f"⚠ {len(errors)} filename(s) could not be parsed — click to review"):
            st.dataframe(pd.DataFrame(errors, columns=["Filename"]), use_container_width=True)

    file_parts = pd.DataFrame(records, columns=["Q paper/Mark Scheme", "Q Number", "Q part"])

    qp_df = run_query(file_parts, QP_PATTERNS[board])
    ms_df = run_query(file_parts, MS_PATTERNS[board])

    # Stat strip
    st.markdown(f"""
    <div class="stat-strip">
      <div class="stat-card">
        <div class="num">{len(filenames)}</div>
        <div class="lbl">Filenames loaded</div>
      </div>
      <div class="stat-card">
        <div class="num">{len(qp_df)}</div>
        <div class="lbl">QParts rows</div>
      </div>
      <div class="stat-card">
        <div class="num">{len(ms_df)}</div>
        <div class="lbl">MSParts rows</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # Preview tabs
    tab_qp, tab_ms, tab_raw = st.tabs(["📋 QParts", "📋 MSParts", "🗂 File Parts Info"])

    with tab_qp:
        if qp_df.empty:
            st.info("No question-paper records matched the filter for this board.")
        else:
            st.dataframe(qp_df, use_container_width=True, hide_index=True)

    with tab_ms:
        if ms_df.empty:
            st.info("No mark-scheme records matched the filter for this board.")
        else:
            st.dataframe(ms_df, use_container_width=True, hide_index=True)

    with tab_raw:
        st.caption("Full 'File Parts Info' table before grouping/filtering.")
        st.dataframe(file_parts, use_container_width=True, hide_index=True)

    # Download
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="section-label">Step 4 — Download</div>', unsafe_allow_html=True)

    excel_bytes = build_excel(qp_df, ms_df)
    st.download_button(
        label="⬇  Download QP_MS_Parts.xlsx",
        data=excel_bytes,
        file_name="QP_MS_Parts.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Upload an Excel file above to get started.")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="footer">Examinent · File Parts Extractor · '
    'Built for CAIE, Edexcel &amp; IBDP workflows</div>',
    unsafe_allow_html=True,
)
