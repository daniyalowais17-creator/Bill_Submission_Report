import streamlit as st
import tempfile
from pathlib import Path
import traceback
import importlib.util

st.set_page_config(
    page_title="Billing Submission Automation",
    page_icon="📊",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Styling — single glassmorphic hero, Apple-style restraint
# ---------------------------------------------------------------------------
st.markdown("""
<style>
.block-container {max-width: 880px; padding-top: 2.6rem;}

/* Soft ambient backdrop so the glass has something to refract */
.stApp {
    background:
        radial-gradient(900px 500px at 20% -10%, rgba(63,94,119,.08), transparent 60%),
        radial-gradient(700px 500px at 100% 10%, rgba(28,140,107,.06), transparent 55%),
        #FAFAF8;
}

/* ---- Hero: one unified glass card ---- */
.hero-card{
    background: rgba(255,255,255,.62);
    backdrop-filter: blur(22px) saturate(160%);
    -webkit-backdrop-filter: blur(22px) saturate(160%);
    border: 1px solid rgba(255,255,255,.7);
    border-radius: 28px;
    padding: 56px 48px 40px;
    text-align: center;
    box-shadow:
        0 30px 60px -30px rgba(23,27,34,.18),
        0 1px 0 rgba(255,255,255,.8) inset;
    margin-bottom: 22px;
    margin-top:20px;
}

.eyebrow{
    display:inline-flex; align-items:center; gap:7px;
    color:#8A8F99; font-size:11px; font-weight:700; letter-spacing:.12em; text-transform:uppercase;
    margin-bottom: 14px;
}

.hero-title{
    font-size: 2.4rem; font-weight: 750; letter-spacing:-.025em; line-height:1.15;
    color:#14171D; margin: 0 0 12px;
}

.hero-sub{
    font-size: 15px; color:#6B7280; line-height:1.6; max-width:520px;
    margin: 0 auto 30px;
}

/* ---- Feature pill grid: 4 items, evenly aligned ---- */
.pill-grid{
    display:grid; grid-template-columns: repeat(4, 1fr); gap:12px;
    max-width: 640px; margin: 0 auto 30px;
}
@media (max-width: 700px){ .pill-grid{ grid-template-columns: repeat(2, 1fr); } }

.pill{
    display:flex; align-items:center; justify-content:center; gap:8px;
    background: rgba(255,255,255,.72);
    border: 1px solid rgba(23,27,34,.08);
    border-radius: 999px;
    padding: 10px 14px;
    font-size: 12.5px; font-weight: 600; color:#3A3F47;
    white-space: nowrap;
    box-shadow: 0 1px 0 rgba(255,255,255,.9) inset;
}

.pill-dot{
    width:7px; height:7px; border-radius:50%; flex:0 0 auto;
    background:#1C8C6B;
    box-shadow: 0 0 0 0 rgba(28,140,107,.55);
    animation: breathe 2.1s ease-in-out infinite;
}
@keyframes breathe{
    0%   { box-shadow: 0 0 0 0 rgba(28,140,107,.45); opacity:1; }
    70%  { box-shadow: 0 0 0 6px rgba(28,140,107,0); opacity:.55; }
    100% { box-shadow: 0 0 0 0 rgba(28,140,107,0); opacity:1; }
}

.hero-divider{
    height:1px; width:100%;
    background: linear-gradient(90deg, transparent, rgba(23,27,34,.10), transparent);
    margin: 4px 0 0;
}

/* ---- Functional glass sub-panel (uploads + process) ---- */
.glass-card{
    background: rgba(255,255,255,.72);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid rgba(23,27,34,.07);
    border-radius: 20px;
    box-shadow: 0 20px 40px -28px rgba(23,27,34,.20);
}

.step-label {font-weight: 650; font-size: 0.85rem; margin-bottom: 0.2rem; color:#171B22;}
.step-hint {opacity: 0.6; font-size: 0.78rem; margin-top: -0.1rem; margin-bottom: 0.4rem; color:#171B22;}

/* Process button: dark ink pill */
div.stButton > button[kind="primary"] {
    background-color: #171B22;
    border: 1px solid #171B22;
    color: #ffffff;
    font-weight: 650;
    border-radius: 10px;
    transition: background-color 0.15s ease, border-color 0.15s ease;
}
div.stButton > button[kind="primary"]:hover:not(:disabled) {
    background-color: #000000;
    border-color: #000000;
}
div.stButton > button[kind="primary"]:disabled {
    background-color: #E4E1D9;
    border: 1px solid #E4E1D9;
    color: rgba(23, 27, 34, 0.45);
    opacity: 1;
}
div.stButton > button[kind="primary"]:disabled:hover {
    cursor: not-allowed;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Hero: single glass card, heading, subtitle, 4 feature pills
# ---------------------------------------------------------------------------
st.markdown("""
<div class="hero-card">
    <div class="eyebrow">Automated Reporting Engine</div>
    <div class="hero-title">Billing Submission Automation</div>
    <div class="hero-sub">Upload BDR, Case_AR and Master — we handle matching, formatting,
    and submission-ready output.</div>
    <div class="pill-grid">
        <div class="pill"><span class="pill-dot"></span>Zero-code automation</div>
        <div class="pill"><span class="pill-dot"></span>Instant Excel output</div>
        <div class="pill"><span class="pill-dot"></span>Real-time processing</div>
        <div class="pill"><span class="pill-dot"></span>Local &amp; private</div>
    </div>
    <div class="hero-divider"></div>
</div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Load the processing engine from web_app.py.
# ---------------------------------------------------------------------------
ENGINE_PATH = Path(__file__).with_name("web_app.py")

if not ENGINE_PATH.exists():
    st.error("`web_app.py` wasn't found next to `streamlit_app.py`. Make sure both files sit in the same folder in your repo.")
    st.stop()


@st.cache_resource
def load_engine():
    spec = importlib.util.spec_from_file_location("billing_engine", ENGINE_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


try:
    engine = load_engine()
except ModuleNotFoundError as e:
    missing_pkg = str(e).split("'")[1] if "'" in str(e) else str(e)
    st.error(
        f"The app couldn't start because the **`{missing_pkg}`** package isn't installed.\n\n"
        f"Fix: open `requirements.txt` in your repo (the file must be named exactly "
        f"`requirements.txt`, not `requirements_streamlit.txt`) and make sure `{missing_pkg}` "
        f"is listed in it. Then push the change and, in Streamlit Cloud, use "
        f"**Manage app → Reboot** to reinstall dependencies."
    )
    with st.expander("Technical details"):
        st.code(traceback.format_exc())
    st.stop()
except Exception:
    st.error("Something went wrong while loading the processing engine.")
    with st.expander("Technical details"):
        st.code(traceback.format_exc())
    st.stop()

# ---------------------------------------------------------------------------
# Uploads + process — glass sub-panel below the hero
# ---------------------------------------------------------------------------
st.markdown('<div class="glass-card">', unsafe_allow_html=True)

st.markdown('<div class="step-label">1. BDR</div>', unsafe_allow_html=True)
st.markdown('<div class="step-hint">Billing Detail Report — CSV or TXT</div>', unsafe_allow_html=True)
bdr_file = st.file_uploader("BDR", type=["csv", "txt"], key="bdr", label_visibility="collapsed")

st.markdown('<div class="step-label">2. Case_AR</div>', unsafe_allow_html=True)
st.markdown('<div class="step-hint">Case AR workbook — Excel</div>', unsafe_allow_html=True)
case_file = st.file_uploader("Case_AR", type=["xlsx", "xls"], key="case_ar", label_visibility="collapsed")

st.markdown('<div class="step-label">3. Master</div>', unsafe_allow_html=True)
st.markdown('<div class="step-hint">Billing Submission Master — Excel</div>', unsafe_allow_html=True)
master_file = st.file_uploader("Master", type=["xlsx", "xls"], key="master", label_visibility="collapsed")

ready = bdr_file is not None and case_file is not None and master_file is not None

missing_files = []
if bdr_file is None:
    missing_files.append("BDR")
if case_file is None:
    missing_files.append("Case_AR")
if master_file is None:
    missing_files.append("Master")

if not ready:
    st.caption("Upload all three files above to enable processing.")

button_help = None if ready else f"Upload {', '.join(missing_files)} to enable processing"

process_clicked = st.button(
    "Process files",
    type="primary",
    disabled=not ready,
    use_container_width=True,
    help=button_help,
)

st.markdown('</div>', unsafe_allow_html=True)
st.write("")

if process_clicked:
    temp_dir = Path(tempfile.mkdtemp(prefix="billing_submission_"))
    bdr_path = temp_dir / bdr_file.name
    case_path = temp_dir / case_file.name
    master_path = temp_dir / master_file.name
    output_path = temp_dir / "Billing_Submission_Master_File.xlsx"

    try:
        bdr_path.write_bytes(bdr_file.getvalue())
        case_path.write_bytes(case_file.getvalue())
        master_path.write_bytes(master_file.getvalue())

        with st.status("Processing your files...", expanded=True) as status:
            st.write(f"BDR: `{bdr_file.name}`")
            st.write(f"Case_AR: `{case_file.name}`")
            st.write(f"Master: `{master_file.name}`")

            engine.process_one(bdr_path, case_path, master_path, output_path)

            if not output_path.exists() or output_path.stat().st_size == 0:
                raise RuntimeError("Processing finished but no output file was created.")

            status.update(label="Done", state="complete", expanded=False)

        st.success("Your Billing Submission Master File is ready.")

        data = output_path.read_bytes()
        st.download_button(
            "Download Billing_Submission_Master_File.xlsx",
            data=data,
            file_name="Billing_Submission_Master_File.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(f"File size: {len(data) / 1024 / 1024:.2f} MB")

    except Exception:
        st.error("Processing failed — no file was created.")
        with st.expander("Technical details"):
            st.code(traceback.format_exc())
