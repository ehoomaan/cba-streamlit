import streamlit as st
from cba_generator import generate_cba_xlsx, default_filename

st.set_page_config(page_title="CBA Matrix Generator", layout="wide")
st.title("CBA Matrix Generator")

project_name = st.text_input("Project name", value="")

st.subheader("Options")
n_opts = st.number_input("Number of options", min_value=2, max_value=8, value=3, step=1)

options = []
cols = st.columns(int(n_opts))
for i in range(int(n_opts)):
    with cols[i]:
        options.append(st.text_input(f"Option {i+1}", value=f"Option {chr(65+i)}"))

options = [o.strip() for o in options if o.strip()]

st.subheader("Factors")
if "factors" not in st.session_state:
    st.session_state.factors = [{"name": "Factor 1"}]

c1, c2 = st.columns(2)
with c1:
    if st.button("Add factor"):
        st.session_state.factors.append({"name": f"Factor {len(st.session_state.factors)+1}"})
with c2:
    if st.button("Remove last factor") and len(st.session_state.factors) > 1:
        st.session_state.factors.pop()

for i, f in enumerate(st.session_state.factors):
    f["name"] = st.text_input(f"Factor {i+1} name", value=f.get("name", ""), key=f"factor_{i}")

data = {"project_name": project_name, "options": options, "factors": st.session_state.factors}

st.divider()
disabled = (not project_name.strip()) or (len(options) < 2)

if st.button("Generate Excel", disabled=disabled):
    xlsx = generate_cba_xlsx(data)
    st.download_button(
        "Download Excel workbook",
        data=xlsx,
        file_name=default_filename(project_name),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
