import streamlit as st
from spire.pdf import PdfDocument, FileFormat
import tempfile
import os

#st.set_page_config(page_title="PDF to Excel Converter", page_icon="📄")

def app():
    main()

def main():

    st.title("📄 PDF to Excel Converter")
    st.write("Easily convert PDF files to Excel format using Spire.PDF")

    # File uploader
    uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

    if uploaded_file is not None:
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf.write(uploaded_file.read())
            input_path = temp_pdf.name

        output_path = os.path.join(tempfile.gettempdir(), "converted_output.xlsx")

        try:
            with st.spinner("Converting PDF to Excel... Please wait ⏳"):
                pdf = PdfDocument()
                pdf.LoadFromFile(input_path)
                pdf.SaveToFile(output_path, FileFormat.XLSX)
                pdf.Close()

            st.success("✅ PDF successfully converted to Excel!")

            # Provide download button
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Excel File",
                    data=f,
                    file_name="Converted_File.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"❌ Conversion failed: {e}")

        finally:
            # Clean up temporary files
            if os.path.exists(input_path):
                os.remove(input_path)
            if os.path.exists(output_path):
                os.remove(output_path)
