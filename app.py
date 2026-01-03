import streamlit as st
import os
import uuid

# IMPORTANT: updated import name
from processor import process_pdf, save_consolidated_excel

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="NCRP Document Analyzer",
    layout="centered"
)

st.title("NCRP Document Analyzer")
st.write("Upload cybercrime complaint PDFs and download a structured Excel report.")

# =====================================================
# BACKEND STORAGE
# =====================================================
BACKEND_FOLDER = "uploaded_pdfs"
os.makedirs(BACKEND_FOLDER, exist_ok=True)

# =====================================================
# FILE UPLOAD
# =====================================================
uploaded_files = st.file_uploader(
    "Upload PDF(s)",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"{len(uploaded_files)} PDF(s) uploaded.")

    if st.button("Start Processing"):
        all_records = []

        with st.spinner("Processing PDFs..."):
            for uploaded_file in uploaded_files:
                try:
                    # ---------------------------------------------
                    # SAVE PDF SAFELY
                    # ---------------------------------------------
                    unique_name = f"{uuid.uuid4().hex}_{uploaded_file.name}"
                    backend_path = os.path.join(BACKEND_FOLDER, unique_name)

                    with open(backend_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())

                    # ---------------------------------------------
                    # PROCESS PDF (RETURNS LIST OF COMPLAINTS)
                    # ---------------------------------------------
                    records = process_pdf(backend_path)

                    if records:
                        all_records.extend(records)

                except Exception as e:
                    st.error(f"Failed to process {uploaded_file.name}: {e}")

        if not all_records:
            st.error("No valid complaints were extracted from the PDFs.")
            st.stop()

        # =================================================
        # SAVE EXCEL
        # =================================================
        output_excel_path = os.path.join(
            BACKEND_FOLDER,
            "Consolidated_Report.xlsx"
        )

        save_consolidated_excel(all_records, output_excel_path)

        st.success(f"Processing complete. {len(all_records)} unique complaint(s) extracted.")

        # =================================================
        # DOWNLOAD BUTTON
        # =================================================
        with open(output_excel_path, "rb") as f:
            st.download_button(
                label="Download Consolidated Excel",
                data=f,
                file_name="Consolidated_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
