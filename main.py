import integration
import os
from dotenv import load_dotenv

load_dotenv()

def main():
    try:
        if os.path.exists("/Users/godfather/Downloads/evan-franke.xlsx") == True or os.path.exists("/Users/godfather/Downloads/dionaks-franke.xlsx") == True:
            os.remove("/Users/godfather/Downloads/evan-franke.xlsx")
            os.remove("/Users/godfather/Downloads/dionaks-franke.xlsx")
    except:
        pass
    driver = integration.driver_installer()
    try:
        integration.extract_excel_file_from_evan(driver)
    except:
        print("An error occured while extracting the .xlsx file from evan.com.tr")
    try:
        integration.extract_excel_file_from_dionaks(driver)
        print("Both files are successfully generated. New .xls file has been generating now.")
    except:
        print("An error occured while extracting the .xlsx file from dionaks.com")
    evan_fpath = "/Users/godfather/Downloads/evan-franke.xlsx"
    dionaks_fpath = "/Users/godfather/Downloads/dionaks-franke.xlsx"
    excel_list = integration.mutual_list_returner(evan_fpath, dionaks_fpath, cname = "stockCode", price_column = "price1", rebate_column = "rebate", rebate_type_column = "rebateType")
    headers = ["stockCode", "price1", "rebate", "rebateType"]
    integration.create_dataframe_and_extract_to_xls(headers, excel_list)
    print("Excel file has been successfully generated. The submission process is starting.")
    integration.submit_file_to_dionaks(driver)
    print("All Done!")
if __name__ == "__main__":
    main()