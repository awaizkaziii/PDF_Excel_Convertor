import streamlit as st
import PDIG
import Text_extractor


st.set_page_config(page_title="PDF To Excel", layout="wide")

def main():    
    st.sidebar.title("PDF To Excel Converter")
    app_choice = st.sidebar.radio("Choose an App", ["PDF To Excel", "OCR"])

    if app_choice == "PDF To Excel":
        PDIG.app()
    elif app_choice == "OCR":
        Text_extractor.app()



if __name__ == "__main__":
    main()
