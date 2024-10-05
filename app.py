
import streamlit as st
import pandas as pd
import requests
import os

def download_pdf(url, file_path):
    try:
        with requests.get(url) as r:
            r.raise_for_status()  # Check if the request was successful
            with open(file_path, 'wb') as f:
                f.write(r.content)  # Write the content to a file
        return True
    except requests.HTTPError as e:
        st.error(f'HTTP error occurred while downloading {url}: {e}')  # Display HTTP error
        return False
    except Exception as e:
        st.error(f'An error occurred: {e}')  # Display any other errors
        return False

def main():
    st.title("PDF Downloader from Excel")

    # File uploader for Excel file
    uploaded_file = st.file_uploader("Choose an Excel file", type='xlsx')

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)

        if 'InvoiceCopy' in df.columns and 'Invoice No.(s)' in df.columns:
            downloads_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
            download_count = 0

            # Iterate over the rows of the DataFrame
            for index, row in df.iterrows():
                invcopy = row['InvoiceCopy']
                invoicenumber = row['Invoice No.(s)']

                manifest_pdf = f'https://scope.sparecare.in/Upload/InvoiceCopySPM/{invcopy}'
                file_path = os.path.join(downloads_folder, f'{invoicenumber}_{invcopy}.pdf')  # Ensure .pdf extension

                if download_pdf(manifest_pdf, file_path):
                    download_count += 1

            st.success(f'Successfully downloaded {download_count} PDFs to your Downloads folder.')

        else:
            st.error("The uploaded Excel file must contain 'InvoiceCopy' and 'Invoice No.(s)' columns.")

if __name__ == '__main__':
    main()