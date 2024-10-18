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
    from selenium import webdriver
    from datetime import datetime
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains

    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=chrome_options)
    
    # driver = webdriver.Chrome()
    # driver.maximize_window()
    # wait = WebDriverWait(driver,15)
    driver.get('https://web.whatsapp.com/')
    time.sleep(30)
    
    conn = pyodbc.connect(
        r'DRIVER={ODBC Driver 17 for SQL Server};'
        r'SERVER=4.240.64.61,1232;'
        r'DATABASE=z_scope;'
        r'UID=Utkrishtsa;'
        r'PWD=AsknSDV*3h9*RFhkR9j73;')
    cursor = conn.cursor()
    
    
    
    cursor.execute('Uad_Gainer_Lr_Alter_Msg_Automation_Details_sp')
    conn.commit() 
    
    
    sql_data = pd.read_sql_query('''SELECT  *
     FROM  Uad_Gainer_Lr_Alter_Msg_Automation_Details
    where msg_status is null and cast(ManifestDate AS date)= cast(getdate() as date) and format(GETDATE() ,'hh:mm') <='12:15'
     ''',conn)
    
    df = sql_data
    for Do,brand,dealer,location,buyerdetails,invoiceNo,InvoiceAmount,lsp,lrn,box,Group,conct,spm in zip( df['DispatchOrderNo'],
        df['Brand'],df['SellingDealer'],df['SellingLocation'], df['BuyingDealer_Location'],
                                df['InvoiceNumber'],df['InvoiceAmount'],df['finallsp'] ,df['LRNumber'],df['Boxes'],                                                             
                        df['Whatsapp_SPM_Name'],df['Contact_no'],df['SPM_Name']):    
        # after correctionn loginc
        
        print("Contact Searching For :", group_name)
        dt = datetime.now().strftime('%d-%b-%y on %H:%M')
        status = "Msg sent on : "+ dt
        new_chat= "//div[@title='New chat']"
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,new_chat))).click()
      
        new_Search ="//div[@class='_ai07 _ai01 _akmh']//p[@class='selectable-text copyable-text x15bjb6t x1n2onr6']"
     
        search_box = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,new_Search)))
        #.send_keys(Group)
        search_box.send_keys(Keys.CONTROL + "a")
        search_box.send_keys(Keys.BACKSPACE)
       
        group_name =conct
        search_box.send_keys(group_name)
        search_box.send_keys(Keys.ENTER) 
        contact_naem ='//div[@class="_ak8q"]//div[@class="x1c4vz4f x3nfvp2 xuce83p x1bft6iq x1i7k8ik xq9mrsl x6s0dn4"]'
        try:
            dr = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.XPATH,
                                     "//span[@class='_ao3e']")))
            error = WebDriverWait(driver, 4).until(EC.visibility_of_element_located((By.XPATH, "//span[contains(text(),'No results found for')]")))
            er_message = error.text
            if "No results found for" in er_message:
                print("contact number not found :", group_name)
                search_box.clear()
                continue
        except:
            pass
        
    
        
        #WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH,contact_naem))).click()
                
        dfs ='Seller - '+dealer+'_'+location
        
        #con = '@ '+str(conct).replace('(','').replace(')','')
        name = f'*Dear Mr. {spm},*'
        d = datetime.today()
        d = d.strftime('%d-%b-%Y')
        #sub = '*Parts Buying Opportunity for '+d+'*'
        sub = 'Please note LR is generated for following order'
        seller = f'*{dfs}*'
        
        N_p = f'*Buyer  : {buyerdetails}*'
        T_p = f'*Invoice No :  {invoiceNo}*'
        A_p = f'*Invoice Value : {InvoiceAmount}*'
        N_B = f'*No of Box - {box}*'
        LR = f'*LR Number : {lrn}*'
        courier_name = f'*Courier Name : {lsp}*'
        
        info = f'Kindly take printout of packing slip from Scope portal & paste on boxes.Also request to share the Image of Packed boxes (after pasting packing slip).'
    
        ft = 'This helps in faster pickup & early sales realisation.'
        ft2 = 'Thanks & Regards'
        ft3 = 'Team Gainer'
    
    
        msg_f= f'''{name}
    {seller}
    
    {sub}
    
    {N_p}
    {T_p}
    {A_p}
    {N_B}
    
    {LR}
    {courier_name}
    
    {info}
    
    {ft}
    
    {ft2}
    {ft3}
    '''
        pc.copy(msg_f)
        
    
        try:
            message_box_xpath = '//div[@aria-placeholder="Type a message"]'
            message_box = wait.until(EC.visibility_of_element_located((By.XPATH, message_box_xpath)))
            message_box.click()
    
            ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
            ActionChains(driver).send_keys(Keys.RETURN).perform()
        except:
            pass
        time.sleep(2)
        
        upd = f"UPDATE Uad_Gainer_Lr_Alter_Msg_Automation_Details SET MSG_STATUS='{status}' WHERE DispatchOrderNo={Do}"
        cursor.execute(upd)
        conn.commit() 
            
    print('Loop exits')
    
    cursor.close()
    conn.close()

if __name__ == '__main__':
    main()
