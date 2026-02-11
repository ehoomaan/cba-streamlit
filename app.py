import streamlit as st
from cba_generator import generate_cba_from_uploaded_template

st.set_page_config(page_title="CBA Matrix Generator", layout="wide")
st.title("CBA Matrix Generator")

purpose = st.text_input(
    "Purpose of CBA (e.g., Deep Foundation System, Support of Excavation Systems, Underpinning)",
    value="",
)
project_name = st.text_input("Project Name", value="")
project_location = st.text_input("Project Location", value="")

uploaded = st.file_uploader("Upload your XLSX template (row labels in column A)", type=["xlsx", "xlsm"])

disabled = (
    uploaded is None
    or not purpose.strip()
    or not project_name.strip()
    or not project_location.strip()
)

if st.button("Generate Excel", disabled=disabled):
    xlsx_bytes, out_name = generate_cba_from_uploaded_template(
        uploaded_xlsx_bytes=uploaded.read(),
        purpose=purpose.strip(),
        project_name=project_name.strip(),
        project_location=project_location.strip(),
        sheet_name=None,
    )

    st.download_button(
        "Download generated workbook",
        data=xlsx_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
