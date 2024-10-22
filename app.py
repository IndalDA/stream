import streamlit as st
import pandas as pd
import requests
import io
import zipfile
import yt_dlp

def download_pdf(url):
    try:
        r = requests.get(url)
        r.raise_for_status()  # Check if the request was successful
        return r.content  # Return the content of the PDF file
    except requests.HTTPError as e:
        st.error(f'HTTP error occurred while downloading {url}: {e}')  # Display HTTP error
        return None
    except Exception as e:
        st.error(f'An error occurred: {e}')  # Display any other errors
        return None

def pdf_downloader():
    st.title("Spare Care PDF Downloader")

    # File uploader for Excel file
    uploaded_file = st.file_uploader("Choose an Excel file", type='xlsx')

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)

        if 'InvoiceCopy' in df.columns and 'Invoice No.(s)' in df.columns:
            download_count = 0
            pdf_files = {}  # Dictionary to store PDF content

            # Iterate over the rows of the DataFrame
            for index, row in df.iterrows():
                invcopy = row['InvoiceCopy']
                invoicenumber = row['Invoice No.(s)']

                manifest_pdf = f'https://scope.sparecare.in/Upload/InvoiceCopySPM/{invcopy}'
                pdf_content = download_pdf(manifest_pdf)

                if pdf_content:
                    pdf_files[f'{invoicenumber}_{invcopy}.pdf'] = pdf_content
                    download_count += 1

            if download_count > 0:
                # Create a ZIP file with all the PDFs
                zip_buffer = io.BytesIO()
                with st.spinner('Preparing download...'):
                    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
                        for filename, content in pdf_files.items():
                            zip_file.writestr(filename, content)

                # Provide download link for the ZIP file
                st.success(f'Successfully downloaded {download_count} PDFs.')
                st.download_button(
                    label="Download all PDFs as ZIP",
                    data=zip_buffer.getvalue(),
                    file_name="pdfs.zip",
                    mime="application/zip"
                )
        else:
            st.error("The uploaded Excel file must contain 'InvoiceCopy' and 'Invoice No.(s)' columns.")

def download_playlist(playlist_link):
    ydl_opts = {
        'format': 'best',  # Download the best quality available
        'outtmpl': '%(title)s.%(ext)s',  # Save with the video title as the filename
        'cookiefile': 'cookies.txt',  # Add your cookies file here
        'noplaylist': False,  # Enable playlist downloading
    }

    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([playlist_link])
        st.success("Download complete.")
    except yt_dlp.utils.DownloadError as e:
        st.error(f"An error occurred: {e}. Please check if the video requires login or is publicly accessible.")

def youtube_playlist_downloader():
    st.title("YouTube Playlist Downloader")
    
    playlist_link = st.text_input("Enter the YouTube playlist URL:")

    if st.button("Download Playlist"):
        if playlist_link:
            st.info(f"Downloading playlist: {playlist_link}")
            download_playlist(playlist_link)
        else:
            st.error("Please enter a valid playlist URL.")

# Main application
def main():
    st.sidebar.title("Menu")
    option = st.sidebar.selectbox("Choose an option", ["PDF Downloader", "YouTube Playlist Downloader"])

    if option == "PDF Downloader":
        pdf_downloader()
    elif option == "YouTube Playlist Downloader":
        youtube_playlist_downloader()

if __name__ == '__main__':
    main()
