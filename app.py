import streamlit as st
from cba_generator import generate_cba_from_uploaded_template

st.set_page_config(page_title="CBA Matrix Generator", layout="wide")
st.title("TEG Choose-By-Advantage Matrix Formatter v1.0")

purpose_choice = st.selectbox(
    "Purpose of CBA",
    [
        "Deep Foundation System",
        "Support of Excavation Systems",
        "Underpinning",
        "Ground Improvement",
        "Other",
    ],
)

if purpose_choice == "Other":
    purpose_other = st.text_input("Enter purpose", value="")
    purpose = purpose_other.strip()
else:
    purpose = purpose_choice.strip()

project_name = st.text_input("Project Name", value="")
project_location = st.text_input("Project Location", value="")

uploaded = st.file_uploader("Upload your XLSX template (row labels in column A)", type=["xlsx", "xlsm"])

disabled = (
    uploaded is None
    or not purpose
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
