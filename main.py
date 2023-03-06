import streamlit as st

from utils import process_input_file

# Create an API key input field:
api_key = st.text_input("Enter your Google API key", type="password")

# create a file upload button (disabled until an API key is entered)
uploaded_file = st.file_uploader("Upload an excel file", type=['xlsx'], disabled=not api_key)

# if the user uploaded a file
if uploaded_file is not None:
    if uploaded_file.name.split('.')[-1] != 'xlsx':
        st.error('Please upload an excel file')
    else:
        try:
            # process the file
            file, filename = process_input_file(uploaded_file, api_key=api_key, standalone_mode=False)

            # display a success message
            st.success('File processed successfully')

            # display a download button
            st.download_button("Download processed file", file_name=filename, data=file,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        except Exception as e:
            st.error(e)
