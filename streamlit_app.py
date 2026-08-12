import streamlit as st
import tempfile
from pathlib import Path
import traceback
import importlib.util

st.set_page_config(
    page_title="Billing Submission Automation",
    page_icon="📊",
    layout="centered",
)

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
st.markdown("""
<style>
.block-container {max-width: 780px; padding-top: 2.2rem;}

.header-block {margin-bottom: 1.5rem;}
.header-block h1 {
    font-size: 1.35rem;
    font-weight: 650;
    letter-spacing: -0.01em;
    margin: 0 0 0.4rem 0;
    line-height: 1.3;
}
.header-block .subtitle {
    margin: 0;
    opacity: 0.7;
    font-size: 0.92rem;
    line-height: 1.5;
}

.step-label {font-weight: 600; font-size: 0.85rem; margin-bottom: 0.25rem;}
.step-hint {opacity: 0.65; font-size: 0.82rem; margin-top: -0.15rem; margin-bottom: 0.5rem;}
hr {margin: 1.5rem 0;}

/* Process button: red + ready when enabled, muted + clearly disabled otherwise */
div.stButton > button[kind="primary"] {
    background-color: #E5484D;
    border: 1px solid #E5484D;
    color: #ffffff;
    font-weight: 600;
    transition: background-color 0.15s ease, border-color 0.15s ease;
}
div.stButton > button[kind="primary"]:hover:not(:disabled) {
    background-color: #C93D42;
    border-color: #C93D42;
}
div.stButton > button[kind="primary"]:disabled {
    background-color: #262A31;
    border: 1px solid #363B44;
    color: rgba(255, 255, 255, 0.4);
    opacity: 1;
}
div.stButton > button[kind="primary"]:disabled:hover {
    cursor: not-allowed;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-block">
<h1>📊 Billing Submission Automation</h1>
<p class="subtitle">Upload your BDR, Case_AR, and Master files below, then hit process. The tool builds the Billing Submission Master File for you.</p>
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

st.divider()

# ---------------------------------------------------------------------------
# Uploads
# ---------------------------------------------------------------------------
st.markdown('<div class="step-label">1. BDR</div>', unsafe_allow_html=True)
st.markdown('<div class="step-hint">Billing Detail Report — CSV or TXT</div>', unsafe_allow_html=True)
bdr_file = st.file_uploader("BDR", type=["csv", "txt"], key="bdr", label_visibility="collapsed")

st.markdown('<div class="step-label">2. Case_AR</div>', unsafe_allow_html=True)
st.markdown('<div class="step-hint">Case AR workbook — Excel</div>', unsafe_allow_html=True)
case_file = st.file_uploader("Case_AR", type=["xlsx", "xls"], key="case_ar", label_visibility="collapsed")

st.markdown('<div class="step-label">3. Master</div>', unsafe_allow_html=True)
st.markdown('<div class="step-hint">Billing Submission Master — Excel</div>', unsafe_allow_html=True)
master_file = st.file_uploader("Master", type=["xlsx", "xls"], key="master", label_visibility="collapsed")

st.divider()

ready = bdr_file is not None and case_file is not None and master_file is not None

# Build a specific "what's missing" message for the hover tooltip
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

if st.button(
    "Process files",
    type="primary",
    disabled=not ready,
    use_container_width=True,
    help=button_help,
):
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
