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
# Styling — dark rainy-blue glass hero, shadowed upload cards
# ---------------------------------------------------------------------------
st.markdown("""
<style>
.block-container {max-width: 880px; padding-top: 2.6rem;}

.stApp {
    background:
        radial-gradient(1000px 600px at 15% -10%, rgba(59,88,113,.20), transparent 60%),
        radial-gradient(800px 500px at 90% 0%, rgba(28,66,92,.14), transparent 55%),
        #EEF1F5;
}

/* ---- Hero: dark rainy-blue gradient glass card ---- */
.hero-card{
    background:
        radial-gradient(650px 320px at 12% 0%, rgba(255,255,255,.12), transparent 60%),
        radial-gradient(550px 420px at 100% 100%, rgba(122,166,199,.28), transparent 55%),
        linear-gradient(160deg, #223B51 0%, #2F5372 45%, #3E6E90 100%);
    backdrop-filter: blur(22px) saturate(160%);
    -webkit-backdrop-filter: blur(22px) saturate(160%);
    border: 1px solid rgba(255,255,255,.14);
    border-radius: 28px;
    padding: 56px 48px 40px;
    text-align: center;
    box-shadow:
        0 30px 70px -30px rgba(10,20,35,.55),
        0 1px 0 rgba(255,255,255,.12) inset;
    margin-bottom: 22px;
    margin-top: 20px;
}

.eyebrow{
    display:inline-flex; align-items:center; gap:7px;
    color:rgba(255,255,255,.75); font-size:11px; font-weight:700; letter-spacing:.12em; text-transform:uppercase;
    margin-bottom: 14px;
}

.hero-title{
    font-size: 2.4rem; font-weight: 750; letter-spacing:-.025em; line-height:1.15;
    color:#F5F8FC; margin: 0 0 12px;
}

.hero-sub{
    font-size: 15px; color:rgba(255,255,255,.72); line-height:1.6; max-width:520px;
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
    background: rgba(255,255,255,.10);
    border: 1px solid rgba(255,255,255,.22);
    border-radius: 999px;
    padding: 10px 14px;
    font-size: 12.5px; font-weight: 600; color:#EAF0F6;
    white-space: nowrap;
    box-shadow: 0 1px 0 rgba(255,255,255,.08) inset;
}

.pill-dot{
    width:7px; height:7px; border-radius:50%; flex:0 0 auto;
    background:#2FE39B;
    box-shadow: 0 0 0 0 rgba(47,227,155,.55);
    animation: breathe 2.1s ease-in-out infinite;
}
@keyframes breathe{
    0%   { box-shadow: 0 0 0 0 rgba(47,227,155,.45); opacity:1; }
    70%  { box-shadow: 0 0 0 6px rgba(47,227,155,0); opacity:.55; }
    100% { box-shadow: 0 0 0 0 rgba(47,227,155,0); opacity:1; }
}

.hero-divider{
    height:1px; width:100%;
    background: linear-gradient(90deg, transparent, rgba(255,255,255,.18), transparent);
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

/* Each upload step gets its own shadowed card, matching the hero's depth */
.st-key-upload-bdr, .st-key-upload-case_ar, .st-key-upload-master {
    background: rgba(255,255,255,.9);
    border: 1px solid rgba(23,27,34,.06);
    border-radius: 16px;
    padding: 16px 18px 10px;
    margin-bottom: 14px;
    box-shadow: 0 16px 32px -22px rgba(23,27,34,.30);
}

.step-label {font-weight: 650; font-size: 0.85rem; margin-bottom: 0.2rem; color:#171B22;}
.step-hint {opacity: 0.6; font-size: 0.78rem; margin-top: -0.1rem; margin-bottom: 0.4rem; color:#171B22;}

/* Process button: dark ink pill, its own shadow to match the theme */
div.stButton > button[kind="primary"] {
    background-color: #171B22;
    border: 1px solid #171B22;
    color: #ffffff;
    font-weight: 650;
    border-radius: 10px;
    box-shadow: 0 16px 30px -18px rgba(23,27,34,.55);
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
    box-shadow: none;
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
# Uploads + process — glass sub-panel below the hero, each step its own
# shadowed card (targeted via container key -> .st-key-<key> CSS class)
# ---------------------------------------------------------------------------
st.markdown('<div class="glass-card">', unsafe_allow_html=True)

with st.container(key="upload-bdr"):
    st.markdown('<div class="step-label">1. BDR</div>', unsafe_allow_html=True)
    st.markdown('<div class="step-hint">Billing Detail Report — CSV or TXT</div>', unsafe_allow_html=True)
    bdr_file = st.file_uploader("BDR", type=["csv", "txt"], key="bdr", label_visibility="collapsed")

with st.container(key="upload-case_ar"):
    st.markdown('<div class="step-label">2. Case_AR</div>', unsafe_allow_html=True)
    st.markdown('<div class="step-hint">Case AR workbook — Excel</div>', unsafe_allow_html=True)
    case_file = st.file_uploader("Case_AR", type=["xlsx", "xls"], key="case_ar", label_visibility="collapsed")

with st.container(key="upload-master"):
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
