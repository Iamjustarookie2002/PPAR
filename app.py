import streamlit as st
import pandas as pd
import tempfile
import os
import io

# Import our custom modules
from excel_processor import process_excel_report
from pdf_processor import process_pdf_report

def main():
    st.set_page_config(page_title="Excel Processor", page_icon="📊", layout="centered")
    
    st.title("📊 Excel File Processor")
    st.write("Upload an Excel file, process it, and download both Excel and PDF reports.")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'], help="Upload Excel file")
    
    if uploaded_file is not None:
        try:
            with st.spinner("Processing your file and generating reports..."):
                # Read the Excel file
                df = pd.read_excel(uploaded_file, sheet_name="FILES_DAT")
                
                # Create output filenames
                base_name = uploaded_file.name.replace('.xlsx', '').replace('.xls', '')
                excel_filename = f"processed_{base_name}.xlsx"
                pdf_filename = f"report_{base_name}.pdf"
                
                # Create Excel output with colors and formatting
                excel_output = io.BytesIO()
                temp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                temp_excel.close()
                
                # Call process_excel_report directly
                process_excel_report(df, temp_excel.name)
                
                with open(temp_excel.name, "rb") as f:
                    excel_output.write(f.read())
                excel_output.seek(0)
                
                # Clean up temporary Excel file
                os.unlink(temp_excel.name)
                
                # Create PDF report
                pdf_data = process_pdf_report(df, uploaded_file.name)
            
            st.success("✅ File processed successfully! Reports generated.")
            
            # Download buttons
            col1, col2 = st.columns(2)
            with col1:
                st.download_button("📥 Download Processed Excel", 
                                 data=excel_output.getvalue(),
                                 file_name=excel_filename,
                                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col2:
                st.download_button("📄 Download PDF Report", 
                                 data=pdf_data,
                                 file_name=pdf_filename, 
                                 mime="application/pdf")
        
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
    else:
        st.info("👆 Please upload an Excel file to get started!")

if __name__ == "__main__":
    main()
