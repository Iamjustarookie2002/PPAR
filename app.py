from turtle import color
import streamlit as st
import pandas as pd
import tempfile
import os
import io

# Import our custom modules
from excel_processor import process_excel_report
from pdf_processor import process_pdf_report

def main():
    st.set_page_config(page_title="PVM gait lab report", page_icon="🏥", layout="centered")
    
    st.title("🏥 PVM gait lab report")
    st.write("Upload the raw excel file from the pressure platform and click 'Generate Reports' to produce excel report.")
    
    # Initialize session state for storing processed data
    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = None
    if 'pdf_data' not in st.session_state:
        st.session_state.pdf_data = None
    if 'excel_filename' not in st.session_state:
        st.session_state.excel_filename = None
    if 'pdf_filename' not in st.session_state:
        st.session_state.pdf_filename = None
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    
    # File uploader - only for Excel file now
    uploaded_file = st.file_uploader("Choose the raw-data excel file", type=['xlsx', 'xls'], help="Upload Excel file with FILES_DAT and VISITS sheets")
    
    # Optional patient information input section
    st.subheader("📋 Optional Patient Information")
    st.write("Fill in the patient details below (optional):")
    
    # Create two columns for better layout
    col1, col2 = st.columns(2)
    
    with col1:
        species = st.text_input("Species")
        breed = st.text_input("Breed")
        color = st.text_input("Color")
    
    with col2:
        purdue_id = st.text_input("Purdue_ID")
        primary_dvm = st.text_input("Primary DVM")
    
    # Add a button to generate reports
    if uploaded_file is not None:
        # Generate Reports button
        if st.button("Generate Report", type="secondary", use_container_width=True):
            try:
                with st.spinner("Processing your file and generating report..."):
                    # Read the Excel file - both sheets
                    df_files_dat = pd.read_excel(uploaded_file, sheet_name="FILES_DAT")
                    df_visits = pd.read_excel(uploaded_file, sheet_name="VISITS")
                    
                    # Create output filenames
                    base_name = uploaded_file.name.replace('.xlsx', '').replace('.xls', '')
                    excel_filename = f"processed_{base_name}.xlsx"
                    pdf_filename = f"report_{base_name}.pdf"
                    
                    # Create Excel output with colors and formatting
                    excel_output = io.BytesIO()
                    temp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                    temp_excel.close()
                    
                    # Prepare manual patient data
                    manual_patient_data = {
                        'species': species,
                        'breed': breed,
                        'color': color,
                        'purdue_id': purdue_id,
                        'primary_dvm': primary_dvm
                    }
                    
                    # Process Excel with patient data from VISITS sheet and manual inputs
                    process_excel_report(df_files_dat, temp_excel.name, df_visits, manual_patient_data)
                    
                    # Create PDF report from Sheet1 of the processed Excel file
                    # pdf_data = process_pdf_report(temp_excel.name, uploaded_file.name)
                    
                    # Read the processed Excel file into memory
                    with open(temp_excel.name, "rb") as f:
                        excel_output.write(f.read())
                    excel_output.seek(0)
                    
                    # Clean up temporary Excel file AFTER both Excel and PDF processing
                    os.unlink(temp_excel.name)
                    
                    # Store data in session state for persistent downloads
                    st.session_state.excel_data = excel_output.getvalue()
                    # st.session_state.pdf_data = pdf_data
                    st.session_state.excel_filename = excel_filename
                    # st.session_state.pdf_filename = pdf_filename
                    st.session_state.processing_complete = True
                
                # st.success("✅ Reports generated successfully!")
                st.rerun()  # Rerun to show download buttons
            
            except Exception as e:
                st.error(f"❌ Error processing file: {str(e)}")
    
    # Show download buttons if processing is complete
    if st.session_state.processing_complete and st.session_state.excel_data:
        # st.success("✅ Reports generated successfully! Download your files below.")
        
        # col1, col2 = st.columns(2)
        # with col1:
        st.download_button("📥 Download Processed Excel", 
                         data=st.session_state.excel_data,
                         file_name=st.session_state.excel_filename,
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        # with col2:
        #     st.download_button("📄 Download PDF Report", 
        #                      data=st.session_state.pdf_data,
        #                      file_name=st.session_state.pdf_filename, 
        #                      mime="application/pdf")
        
        # Add a button to clear session state and start over
        if st.button("🔄 Process New File", use_container_width=True):
            st.session_state.excel_data = None
            # st.session_state.pdf_data = None
            st.session_state.excel_filename = None
            # st.session_state.pdf_filename = None
            st.session_state.processing_complete = False
            st.rerun()
    
    elif not st.session_state.processing_complete:
        st.info("Please upload the raw-data excel file to get started!")

if __name__ == "__main__":
    main()
