import os
from datetime import datetime

import streamlit as st
from cba_generator import generate_cba_from_uploaded_template

st.set_page_config(page_title="CBA Matrix Generator", layout="centered")

st.markdown(
    """
    <style>
      .cba-title { font-size: 32px; color: #0B5394; margin: 0 0 0.25rem 0; }
    </style>
    <div class="cba-title">TEG Choose-By-Advantage Matrix Formatter</div>
    """,
    unsafe_allow_html=True,
)

# Session storage
if "xlsx_bytes" not in st.session_state:
    st.session_state.xlsx_bytes = None
if "out_name" not in st.session_state:
    st.session_state.out_name = None
if "last_inputs_sig" not in st.session_state:
    st.session_state.last_inputs_sig = None

# --- Form (prevents rerun on each keystroke) ---
with st.form("cba_form", clear_on_submit=False):
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

    # This is NOT "Generate Excel" and avoids regenerating while typing
    submitted = st.form_submit_button("Apply")

# Only run generation when user clicks Apply
if submitted:
    missing = []
    if not purpose:
        missing.append("Purpose")
    if not project_name.strip():
        missing.append("Project Name")
    if not project_location.strip():
        missing.append("Project Location")
    if uploaded is None:
        missing.append("Template XLSX")

    if missing:
        st.session_state.xlsx_bytes = None
        st.session_state.out_name = None
        st.error("Please fill in: " + ", ".join(missing))
    else:
        inputs_sig = (
            purpose,
            project_name.strip(),
            project_location.strip(),
            uploaded.name,
            uploaded.size,
        )

        # Only regenerate if something truly changed since last Apply
        if inputs_sig != st.session_state.last_inputs_sig:
            with st.spinner("Generating formatted workbook..."):
                st.session_state.xlsx_bytes, st.session_state.out_name = generate_cba_from_uploaded_template(
                    uploaded_xlsx_bytes=uploaded.getvalue(),
                    purpose=purpose,
                    project_name=project_name.strip(),
                    project_location=project_location.strip(),
                    sheet_name=None,
                )
            st.session_state.last_inputs_sig = inputs_sig

        st.success("The formatted Excel file is generated. Click Download to download.")

# Download always visible when output exists
if st.session_state.xlsx_bytes:
    st.download_button(
        "Download",
        data=st.session_state.xlsx_bytes,
        file_name=st.session_state.out_name or "CBA.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Fill in the fields, upload the file, then click Apply to generate the formatted Excel file.")

# Footer
_last_updated_ts = os.path.getmtime(__file__)
_last_updated = datetime.fromtimestamp(_last_updated_ts).strftime("%B %d, %Y")
st.markdown("---")
st.caption(f"Last updated on {_last_updated}. Version 0.1")
