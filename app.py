import streamlit as st
from cba_generator import generate_cba_from_uploaded_template
import os
from datetime import datetime

st.set_page_config(page_title="CBA Matrix Generator", layout="centered")
st.markdown(
    """
    <style>
      .cba-title { font-size: 32px; color: #0B5394; margin: 0 0 0.25rem 0; }
    </style>
    <div class="cba-title">TEG Choose-By-Advantage Matrix Formatter </div>
    """,
    unsafe_allow_html=True,
)
purpose_choice = st.selectbox(
    "Purpose of Choose-By-Advantage Matrix:",
    [
        "Deep Foundation System",
        "Support of Excavation Systems",
        "Underpinning",
        "Ground Improvement",
        "Other",
    ],
)

if purpose_choice == "Other":
    purpose_other = st.text_input("Enter purpose:", value="")
    purpose = purpose_other.strip()
else:
    purpose = purpose_choice.strip()

project_name = st.text_input("Project Name:", value="")
project_location = st.text_input("Project Location:", value="")

uploaded = st.file_uploader("Upload your XLSX file from Custom GPT", type=["xlsx", "xlsm"])

disabled = (
    uploaded is None
    or not purpose
    or not project_name.strip()
    or not project_location.strip()
)

# --- Generate automatically (no "Generate Excel" button) ---

ready = (
    uploaded is not None
    and purpose.strip()
    and project_name.strip()
    and project_location.strip()
)

# Initialize storage
if "xlsx_bytes" not in st.session_state:
    st.session_state.xlsx_bytes = None
if "out_name" not in st.session_state:
    st.session_state.out_name = None
if "last_inputs_sig" not in st.session_state:
    st.session_state.last_inputs_sig = None

# Create a signature so we only regenerate when inputs change
inputs_sig = None
if ready:
    inputs_sig = (
        purpose.strip(),
        project_name.strip(),
        project_location.strip(),
        uploaded.name,           # filename
        uploaded.size,           # size
    )

# Auto-generate once when ready AND inputs changed
if ready and inputs_sig != st.session_state.last_inputs_sig:
    st.session_state.xlsx_bytes, st.session_state.out_name = generate_cba_from_uploaded_template(
        uploaded_xlsx_bytes=uploaded.getvalue(),  # safe to call multiple times
        purpose=purpose.strip(),
        project_name=project_name.strip(),
        project_location=project_location.strip(),
        sheet_name=None,
    )
    st.session_state.last_inputs_sig = inputs_sig
    st.success("The formatted Excel file is generated. Click Download to download.")

# Show download button when we have output
if st.session_state.xlsx_bytes:
    st.download_button(
        "Download",
        data=st.session_state.xlsx_bytes,
        file_name=st.session_state.out_name or "CBA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload the Excel file from CBA MATRIX PRO GPT and fill in the fields to generate the formatted Excel file.")
    
# Footer: last updated based on app.py file modified time
_last_updated_ts = os.path.getmtime(__file__)
_last_updated = datetime.fromtimestamp(_last_updated_ts).strftime("%B %d, %Y")

st.markdown("---")
st.caption(f'Last updated on {_last_updated}.')
