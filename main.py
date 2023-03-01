import integration
import send_mail
import os
from dotenv import load_dotenv

load_dotenv()

def main():
    try:
        try:
            if os.path.exists("/Users/godfather/Downloads/eren-franke-hafele.xlsx") == True:
                os.remove("/Users/godfather/Downloads/eren-franke-hafele.xlsx")

            if os.path.exists("/Users/godfather/Downloads/dionaks-franke-hafele.xlsx") == True:
                os.remove("/Users/godfather/Downloads/dionaks-franke-hafele.xlsx")
        except:
            pass
        driver = integration.driver_installer()
        try:
           integration.extract_excel_file_from_evan(driver)
        except:
            print("An error occured while extracting the .xlsx file from evan.com.tr")
            email_message = '''Subject: Update from evan.com.tr to dionaks.com has been failed.\n
            
            There was a problem while extracting the file from evan.com.tr. Check the code once more to see what caused the error.
            '''
        try:
            integration.extract_excel_file_from_dionaks(driver)
            print("Both files are successfully generated. New .xls file has been generating now.")
        except:
            print("An error occured while extracting the .xlsx file from dionaks.com")
            email_message = '''Subject: Update from evan.com.tr to dionaks.com has been failed.\n

                        There was a problem while extracting the file from dionaks.com. Check the code once more to see what caused the error.
                        '''
        evan_fpath = "/Users/godfather/Downloads/eren-franke-hafele.xlsx"
        dionaks_fpath = "/Users/godfather/Downloads/dionaks-franke-hafele.xlsx"
        excel_list = integration.mutual_list_returner(evan_fpath, dionaks_fpath, cname = "stockCode", price_column = "price1", rebate_column = "rebate", rebate_type_column = "rebateType", stock_column="stockAmount", stock_type_column="stockType")
        headers = ["stockCode", "price1", "rebate", "rebateType", "stockAmount", "stockType"]
        integration.create_dataframe_and_extract_to_xls(headers, excel_list)
        print("Excel file has been successfully generated. The submission process is starting.")
        integration.submit_file_to_dionaks(driver)
        print("All Done!")
        email_message_subject = '''Subject: The update from evan.com.tr to dionaks.com has been successfully completed.'''
        email_message_content = '''See the attachment .xls file to see the file which has been uploaded.\nDo not reply to this e-mail.'''

    except:
        email_message_subject = '''Subject: The update from evan.com.tr to dionaks.com has been failed.'''
        email_message_content = '''See the code for further details about the error.\nDo not reply to this e-mail.'''
    send_mail.send_mail_with_excel(email_message_subject, email_message_content)
if __name__ == "__main__":
    main()
