import os
import win32com.client
import subprocess
import pandas as pd
import datetime

# 1. Fetch new emails and update the Database (Report.py)
print("Step 1: Fetching new emails and updating database...")
report_script = r"C:\Users\sawak\OneDrive\デスクトップ\売上メール\Report.py"
try:
    subprocess.run(["python", report_script], check=True)
    print("Database updated successfully.")
except Exception as e:
    print(f"Error running Report.py: {e}")
    # Proceed anyway

# 2. Find the most recent date in the DB
print("Step 2: Finding the most recent date in the Database...")
db_path = r"C:\Users\sawak\OneDrive\デスクトップ\売上メール\データベース\売上管理表.xlsx"
try:
    df = pd.read_excel(db_path, sheet_name='Data')
    df['日付'] = pd.to_datetime(df['日付'])
    max_date = df['日付'].max()
    max_date_str = max_date.strftime("%Y/%m/%d")
    print(f"Found latest date in DB: {max_date_str}")
except Exception as e:
    print(f"Error reading DB: {e}")
    print("Fallback to today's date.")
    max_date_str = datetime.date.today().strftime("%Y/%m/%d")

# 3. Open Excel and trigger the VBA to build the Latest Tab
print("Step 3: Triggering VBA to generate the Latest view...")
excel_path = r"C:\Users\sawak\OneDrive\デスクトップ\売上メール\売上確認.xlsm"

xl = win32com.client.Dispatch("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False

try:
    wb = xl.Workbooks.Open(excel_path)
    
    # Check if '最新' tab already exists and delete it
    for sht in wb.Sheets:
        if sht.Name == "最新":
            sht.Delete()
            print("Deleted old '最新' tab.")
            break
            
    # Set search criteria
    ws_search = wb.Sheets("検索")
    ws_search.Range("D8").Value = "すべて"
    ws_search.Range("D9").Value = max_date_str
    ws_search.Range("D10").Value = max_date_str
    
    # Run the existing macro (which will generate 日別成績 for the max_date)
    xl.Application.Run("売上確認.xlsm!FetchSalesData")
    
    # Rename the newly generated "日別成績" to "最新"
    try:
        ws_day = wb.Sheets("日別成績")
        ws_day.Name = "最新"
        ws_day.Tab.Color = 255 # Red color to make it obvious
        print(f"Renamed '日別成績' to '最新' for date {max_date_str}")
        
    except Exception as e:
        print(f"Could not rename sheet: {e}")
    
    # Ensure Search is the active tab so they see the button when opening
    ws_search.Activate()
    
    wb.Save()
    print("Excel VBA Update completed! The '最新' tab is ready.")

except Exception as e:
    print(f"Error during Excel manipulation: {e}")

finally:
    # Ensure Excel is closed properly
    try:
        wb.Close(SaveChanges=True)
    except:
        pass
    xl.Quit()
    print("Finished auto_update_excel.py.")
