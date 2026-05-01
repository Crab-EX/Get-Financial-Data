import streamlit as st
from pathlib import Path

from functions import fill_es_template_informa, fill_fr_template_pappers

st.set_page_config(
    page_title="Get Financial Data",
    page_icon=":ledger:",
)

st.title("Get Financial Data")

st.text("This app is providing you the raw financial data from Excel files of Data Providers (such as Informa or Pappers).")


registration_number = st.text_input(
    "Company registration number",
    placeholder="Example: 12345678A",
    help="Enter the company registration number before uploading files.",
)

data_source = st.selectbox(
    "Data source",
    options=["Informa", "Pappers"],
    help="Select the provider for the uploaded financial data files.",
)

uploaded_files = st.file_uploader(
    "Upload company Excel files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
)

if registration_number and data_source and uploaded_files:
    st.success(
        f"Registration number {registration_number} received from {data_source}. "
        f"{len(uploaded_files)} XLSX file(s) uploaded."
    )

elif uploaded_files and not registration_number:
    st.warning("Please enter the company registration number to continue.")

generate_button = st.button("Generate output file")

if generate_button:
    if not registration_number:
        st.error("Please enter the company registration number.")
    elif not uploaded_files:
        st.error("Please upload at least one XLSX file.")
    else:
        if data_source == "Informa":
            template_path = Path("Templates") / "ES-Template.xlsx"
            fill_fn = fill_es_template_informa
            output_file_name = f"ES-Template-{registration_number.strip()}.xlsx"
            download_label = "Download filled ES template"
        else:  # Pappers
            template_path = Path("Templates") / "FR-Template.xlsx"
            fill_fn = fill_fr_template_pappers
            output_file_name = f"FR-Template-{registration_number.strip()}.xlsx"
            download_label = "Download filled FR template"

        if not template_path.exists():
            st.error(f"Template file not found: {template_path}")
        else:
            uploaded_files_by_name = {
                uploaded_file.name: uploaded_file.getvalue()
                for uploaded_file in uploaded_files
            }
            try:
                filled_template_bytes = fill_fn(
                    template_bytes=template_path.read_bytes(),
                    registration_number=registration_number.strip(),
                    uploaded_files_by_name=uploaded_files_by_name,
                )
            except Exception as ex:
                st.error(f"Error generating template: {ex}")
            else:
                st.success("Template generated successfully.")
                st.download_button(
                    download_label,
                    data=filled_template_bytes,
                    file_name=output_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

