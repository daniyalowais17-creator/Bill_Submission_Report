import streamlit as st
import pandas as pd
import tempfile
from pathlib import Path
import traceback
import importlib.util

st.set_page_config(
    page_title="Billing Submission Automation",
    page_icon="📊",
    layout="wide",
)

st.markdown("""
<style>
.block-container {max-width: 1150px; padding-top: 2rem;}
.hero {
    padding: 1.5rem 1.8rem;
    border-radius: 16px;
    border: 1px solid rgba(128,128,128,.25);
    margin-bottom: 1.2rem;
}
.card {
    padding: 1rem 1.2rem;
    border-radius: 14px;
    border: 1px solid rgba(128,128,128,.22);
    height: 100%;
}
.small {opacity:.75; font-size:.9rem;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
<h1>📊 Billing Submission Automation</h1>
<p>Upload BDR, Case_AR and Master files to generate a prepared Billing Submission Master File.</p>
</div>
""", unsafe_allow_html=True)

# Load the verified processing engine from web_app.py.
# This module is loaded only for its processing functions; Flask server startup is not executed.
ENGINE_PATH = Path(__file__).with_name("web_app.py")

if not ENGINE_PATH.exists():
    st.error("web_app.py was not found beside streamlit_app.py.")
    st.stop()

@st.cache_resource
def load_engine():
    spec = importlib.util.spec_from_file_location("billing_engine", ENGINE_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

try:
    engine = load_engine()
except Exception as e:
    st.error("Could not load the billing processing engine.")
    st.code(traceback.format_exc())
    st.stop()

st.markdown("### 📁 Upload Files")

c1, c2, c3 = st.columns(3)

with c1:
    st.markdown('<div class="card"><h4>1️⃣ BDR</h4><p class="small">Billing Detail Report</p></div>', unsafe_allow_html=True)
    bdr_file = st.file_uploader(
        "Upload BDR",
        type=["csv", "txt"],
        key="bdr",
        label_visibility="collapsed",
    )

with c2:
    st.markdown('<div class="card"><h4>2️⃣ Case_AR</h4><p class="small">Case AR workbook</p></div>', unsafe_allow_html=True)
    case_file = st.file_uploader(
        "Upload Case_AR",
        type=["xlsx", "xls"],
        key="case_ar",
        label_visibility="collapsed",
    )

with c3:
    st.markdown('<div class="card"><h4>3️⃣ Master</h4><p class="small">Billing Submission Master</p></div>', unsafe_allow_html=True)
    master_file = st.file_uploader(
        "Upload Master",
        type=["xlsx", "xls"],
        key="master",
        label_visibility="collapsed",
    )

st.divider()

ready = bdr_file is not None and case_file is not None and master_file is not None

if not ready:
    st.info("Upload all three files to enable processing.")

if st.button(
    "🚀 PROCESS FILES",
    type="primary",
    disabled=not ready,
    use_container_width=True,
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

        st.markdown("### ⚙️ Processing")

        with st.status("Processing billing submission...", expanded=True) as status:
            st.write(f"BDR: `{bdr_file.name}`")
            st.write(f"Case_AR: `{case_file.name}`")
            st.write(f"Master: `{master_file.name}`")

            # Reuse the existing verified engine.
            engine.process_one(
                bdr_path,
                case_path,
                master_path,
                output_path,
            )

            if not output_path.exists() or output_path.stat().st_size == 0:
                raise RuntimeError("Processing completed but no valid output file was created.")

            status.update(
                label="Processing completed successfully",
                state="complete",
                expanded=False,
            )

        st.success("Billing Submission Master File created successfully.")

        data = output_path.read_bytes()

        st.download_button(
            "⬇️ Download Billing_Submission_Master_File.xlsx",
            data=data,
            file_name="Billing_Submission_Master_File.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.caption(f"Output size: {len(data) / 1024 / 1024:.2f} MB")

    except Exception:
        st.error("Processing failed. No output file was offered for download.")
        with st.expander("Show technical error"):
            st.code(traceback.format_exc())
