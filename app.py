import streamlit as st
import tempfile
import pdfplumber
from processor import process_pdf

st.set_page_config(page_title="Cyber Fraud Analyzer", layout="centered")
st.title("Cyber Fraud PDF Analyzer")
st.write("Upload one or more cybercrime complaint PDFs and download structured Excel reports.")

uploaded_files = st.file_uploader("Upload PDF(s)", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.read())
            tmp_pdf_path = tmp_pdf.name

        output_path = tmp_pdf_path.replace(".pdf",".xlsx")

        with st.spinner(f"Processing {uploaded_file.name}..."):
            process_pdf(tmp_pdf_path, output_path)

        with open(output_path,"rb") as f:
            st.download_button(
                label=f"Download Excel for {uploaded_file.name}",
                data=f,
                file_name=uploaded_file.name.replace(".pdf",".xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
