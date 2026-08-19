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
# Styling — cream hero left, gradient "glass" panel right (Voxai-inspired)
# ---------------------------------------------------------------------------
st.markdown("""
<style>
.block-container {max-width: 1080px; padding-top: 2.4rem;}

/* ---- Left hero panel ---- */
.hero-card{
    background:#FAF7F2; border:1px solid rgba(23,27,34,.08); border-radius:20px;
    padding:38px 34px; height:100%; margin-top: 20px;
}
.eyebrow{
    display:inline-flex; align-items:center; gap:7px;
    border:1px solid rgba(23,27,34,.10); background:rgba(255,255,255,.6);
    color:#5B6472; font-size:11px; font-weight:650; letter-spacing:.04em; text-transform:uppercase;
    padding:6px 12px; border-radius:999px; margin-bottom:18px;
}
.eyebrow .dot{width:6px; height:6px; border-radius:50%; background:#1C8C6B; box-shadow:0 0 8px #1C8C6B; display:inline-block;}
.hero-title{
    font-size:1.85rem; font-weight:700; letter-spacing:-.02em; line-height:1.18;
    color:#171B22; margin:0 0 12px;
}
.hero-sub{font-size:14.5px; color:#5B6472; line-height:1.6; margin:0 0 22px;}
.privacy-note{font-size:12px; color:#5B6472; margin-top:20px;}

/* ---- Right functional panel ---- */
.right-wrap{
    background:linear-gradient(160deg,#3F5E77,#22303F);
    border-radius:20px; padding:24px; height:100%;
}
.right-badge{
    display:inline-flex; align-items:center; gap:7px;
    border:1px solid rgba(255,255,255,.28); background:rgba(255,255,255,.10);
    color:#fff; font-size:11.5px; padding:6px 12px; border-radius:999px; margin-bottom:16px;
}
.right-badge .dot{width:6px; height:6px; border-radius:50%; background:#7FE6C6; box-shadow:0 0 8px #7FE6C6; display:inline-block;}

.glass-card{
    background:rgba(255,255,255,.92); border-radius:16px; padding:20px 20px 6px;
}

.step-label {font-weight: 650; font-size: 0.85rem; margin-bottom: 0.2rem; color:#171B22;}
.step-hint {opacity: 0.6; font-size: 0.78rem; margin-top: -0.1rem; margin-bottom: 0.4rem; color:#171B22;}

/* Process button: dark ink pill, matches hero accent */
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
    background-color: #DCD8CF;
    border: 1px solid #DCD8CF;
    color: rgba(23, 27, 34, 0.45);
    opacity: 1;
}
div.stButton > button[kind="primary"]:disabled:hover {
    cursor: not-allowed;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Hero: left marketing copy / right decorative "glass" intro
# ---------------------------------------------------------------------------
hero_left, hero_right = st.columns([1.1, 0.9], gap="medium")

with hero_left:
    st.markdown("""
    <div class="hero-card">
      <div class="eyebrow"><span class="dot"></span>Runs entirely in this session</div>
      <div class="hero-title">Billing Submission Report</div>
      <div class="hero-sub">Upload BDR, Case_AR and Master File &
      Get back a ready-to-submit Billing Submission Master File.</div>
    <div class="right-wrap">
      <div class="right-badge"><span class="dot"></span>Local &amp; private</div>
    </div>
      <div class="privacy-note"></div>
    </div>
    """, unsafe_allow_html=True)

with hero_right:
    st.markdown("""
    <div class="right-wrap">
      <div class="right-badge"><span class="dot"></span>Local &amp; private</div>
    </div>
    """, unsafe_allow_html=True)

st.write("")

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
# Uploads + process — inside a "glass card" to keep the visual language
# ---------------------------------------------------------------------------
st.markdown('<div class="glass-card">', unsafe_allow_html=True)

st.markdown('<div class="step-label">1. BDR</div>', unsafe_allow_html=True)
st.markdown('<div class="step-hint">Billing Detail Report (.CSV) </div>', unsafe_allow_html=True)
bdr_file = st.file_uploader("BDR", type="csv", key="bdr", label_visibility="collapsed")

st.markdown('<div class="step-label">2. Case_AR</div>', unsafe_allow_html=True)
st.markdown('<div class="step-hint">Case AR workbook (.xlsx)</div>', unsafe_allow_html=True)
case_file = st.file_uploader("Case_AR", type="xlsx", key="case_ar", label_visibility="collapsed")

st.markdown('<div class="step-label">3. Master</div>', unsafe_allow_html=True)
st.markdown('<div class="step-hint">Billing Submission Master (.xlsx)</div>', unsafe_allow_html=True)
master_file = st.file_uploader("Master", type="xlsx", key="master", label_visibility="collapsed")

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
