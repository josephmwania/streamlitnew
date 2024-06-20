import streamlit as st
import pandas as pd
import re
import os

def extract_project_details(text):
    # Extract translator or proofreader
    translator = ""
    proofreader = ""
    linguist_name = text.split(' - ')[0].strip()
    if "translation" in text:
        translator = linguist_name
    elif "proofreading" in text or "review" in text:
        proofreader = linguist_name

    # Extract word count or hours
    word_count_match = re.search(r'(\d+) (words|hours)', text)
    word_count = word_count_match.group(1) + " " + word_count_match.group(2) if word_count_match else ""

    # Extract service name
    service_name_match = re.search(r'(\d+ (words|hours)) (.+?) due', text)
    service_name = service_name_match.group(3) if service_name_match else ""

    # Extract client name
    client_match = re.search(r'due [\d\w\s:.]+ ([\w/]+) #', text)
    client_name = client_match.group(1).strip() if client_match else ""

    # Extract PO number
    po_number_match = re.search(r'#(\w+)', text)
    po_number = "#" + po_number_match.group(1) if po_number_match else ""

    # Extract PO amount
    po_amount_match = re.search(r',\s*(USD|EUR|GBP)\s*(\d+)', text)
    po_amount = po_amount_match.group(1) + " " + po_amount_match.group(2) if po_amount_match else ""

    return client_name, service_name, po_amount, word_count, translator, proofreader, po_number

st.title('Project Details Extraction')
st.write('Upload an Excel file to extract project details.')

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("File uploaded successfully!")

    # Extract project details for each row
    extracted_details = []
    for index, row in df.iterrows():
        extracted_details.append(extract_project_details(row['Project Details']))

    # Create a new DataFrame with the extracted details
    extracted_df = pd.DataFrame(extracted_details, columns=['Client Name', 'Service Name', 'PO Amount', 'Word Count', 'Translator', 'Proofreader', 'PO Number'])

    st.write("Extracted Project Details:")
    st.dataframe(extracted_df)

    # Export the extracted details to a new Excel file
    output_file = 'new_projects.xlsx'
    extracted_df.to_excel(output_file, index=False)

    with open(output_file, "rb") as file:
        btn = st.download_button(
            label="Download Extracted Details",
            data=file,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
