import streamlit as st

from utils import process_input_file

# create a file upload button
uploaded_file = st.file_uploader("Choose a file")

# if the user uploaded a file
if uploaded_file is not None:
    if uploaded_file.name.split('.')[-1] != 'xlsx':
        st.error('Please upload an excel file')
    else:
        # process the file
        file, filename = process_input_file(uploaded_file)

        # display a success message
        st.success('File processed successfully')

        # display a download button
        st.download_button("Download processed file", file_name=filename, data=file,
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
