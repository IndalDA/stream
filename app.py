import streamlit as st
import pandas as pd
import requests
import io
import zipfile  # Import the zipfile module

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

def main():
    st.title("Spare care PDF Downloader")

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

if st.button("Click here To Send Lr Alter Msg"):
    import os
    import time
    import pyodbc
    import selenium
    import pandas as pd
    import pyperclip as pc
    import streamlit as st
    from datetime import datetime
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
    from webdriver_manager.chrome import ChromeDriverManager
    
    # Set up the Streamlit UI
    st.title("WhatsApp Message Automation")
    st.write("This app automates sending WhatsApp messages.")
    
    # Button to trigger the automation
    if st.button("Start WhatsApp Automation"):
        # Set up Chrome options for headless mode
        chrome_options = Options()
        #chrome_options.add_argument('--headless')  # Enable headless mode
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
    
        # Initialize the Chrome WebDriver
        driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options)
    
        driver.maximize_window()
        wait = WebDriverWait(driver, 15)
    
        # Open WhatsApp Web
        driver.get('https://web.whatsapp.com/')
        st.write("Please scan the QR code to log in to WhatsApp Web.")
        time.sleep(30)  # Wait for user to scan the QR code
    
    # Database connection
    conn = pyodbc.connect(
            r'DRIVER={ODBC Driver 17 for SQL Server};'
            r'SERVER=4.240.64.61,1232;'
            r'DATABASE=z_scope;'
            r'UID=Utkrishtsa;'
            r'PWD=AsknSDV*3h9*RFhkR9j73;')
    cursor = conn.cursor() 
    # Execute stored procedure
    cursor.execute('Uad_Gainer_Lr_Alter_Msg_Automation_Details_sp')
    conn.commit()

    # Fetch SQL data
    sql_data = pd.read_sql_query('''
        SELECT *
        FROM Uad_Gainer_Lr_Alter_Msg_Automation_Details
        WHERE msg_status IS NULL 
          AND CAST(ManifestDate AS date) = CAST(GETDATE() AS date)
          AND FORMAT(GETDATE(), 'hh:mm') <= '12:15'
    ''', conn)

    # Iterate over the SQL data and send WhatsApp messages
    for Do, brand, dealer, location, buyerdetails, invoiceNo, InvoiceAmount, lsp, lrn, box, Group, conct, spm in zip(
        sql_data['DispatchOrderNo'],
        sql_data['Brand'],
        sql_data['SellingDealer'],
        sql_data['SellingLocation'],
        sql_data['BuyingDealer_Location'],
        sql_data['InvoiceNumber'],
        sql_data['InvoiceAmount'],
        sql_data['finallsp'],
        sql_data['LRNumber'],
        sql_data['Boxes'],
        sql_data['Whatsapp_SPM_Name'],
        sql_data['Contact_no'],
        sql_data['SPM_Name']
    ):
        st.write(f"Sending message to contact: {conct}")
        dt = datetime.now().strftime('%d-%b-%y on %H:%M')
        status = f"Msg sent on: {dt}"

        # Search for the contact
        new_chat = "//div[@title='New chat']"
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, new_chat))).click()

        new_search = "//div[@class='_ai07 _ai01 _akmh']//p[@class='selectable-text copyable-text x15bjb6t x1n2onr6']"
        search_box = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, new_search)))
        search_box.send_keys(Keys.CONTROL + "a")
        search_box.send_keys(Keys.BACKSPACE)

        group_name = conct
        search_box.send_keys(group_name)
        search_box.send_keys(Keys.ENTER)

        # Prepare and send the message
        seller = f'Seller - {dealer}_{location}'
        message = f'''*Dear Mr. {spm},*
        {seller}

        Please note LR is generated for the following order:

        *Buyer  : {buyerdetails}*
        *Invoice No :  {invoiceNo}*
        *Invoice Value : {InvoiceAmount}*
        *No of Box - {box}*
        *LR Number : {lrn}*
        *Courier Name : {lsp}*

        Kindly take printout of packing slip from Scope portal & paste on boxes.
        Thanks & Regards, Team Gainer
        '''

        pc.copy(message)

        # Paste and send the message
        message_box_xpath = '//div[@aria-placeholder="Type a message"]'
        message_box = wait.until(EC.visibility_of_element_located((By.XPATH, message_box_xpath)))
        message_box.click()

        ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        ActionChains(driver).send_keys(Keys.RETURN).perform()

        time.sleep(2)

        # Update the database with message status
        update_query = f"UPDATE Uad_Gainer_Lr_Alter_Msg_Automation_Details SET MSG_STATUS='{status}' WHERE DispatchOrderNo={Do}"
        cursor.execute(update_query)
        conn.commit()

        st.success(f"Message sent to {group_name} successfully.")

    # Close the database and browser
    cursor.close()
    conn.close()
    driver.quit()

st.write("WhatsApp Automation complete.")


if __name__ == '__main__':
    main()
