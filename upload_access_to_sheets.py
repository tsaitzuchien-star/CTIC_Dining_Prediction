import pandas as pd
import pyodbc
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import warnings

# 忽略 pandas 的 UserWarning 警告，保持畫面乾淨
warnings.filterwarnings('ignore', category=UserWarning)

# ==========================================
# 1. Google Sheets API 連線設定
# ==========================================
# 請確保 credentials.json 放在同一個資料夾
CREDENTIALS_FILE = 'credentials.json' 
# 你的 Google 試算表名稱
SPREADSHEET_NAME = 'CTIC_Access_Log'

def setup_gspread():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    # 🎯 確保上傳到第一個分頁 (今日預測大腦)
    return client.open(SPREADSHEET_NAME).sheet1

# ==========================================
# 2. 門禁 SQL 連線設定 (本機免密碼 Windows 驗證)
# ==========================================
pums_conn_str = (
    r'DRIVER={SQL Server};'
    r'SERVER=localhost;'    
    r'DATABASE=PUMS;'
    r'Trusted_Connection=yes;' 
)

def main():
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 開始執行資料拋轉作業...")

    # 自動取得今天日期
    target_date = datetime.now().strftime('%Y-%m-%d') 

    # 1. 撈取內網門禁資料
    try:
        conn = pyodbc.connect(pums_conn_str)
        
        # 🌟 SQL 修改：新增 CardID 欄位，並將 Date 命名為 EntryTime
        query = f"""
            SELECT EmpID, CardID, Name, DoorDsc, Date AS EntryTime
            FROM (
                SELECT EmpID, CardID, Name, DoorDsc, Date,
                       ROW_NUMBER() OVER(PARTITION BY EmpID ORDER BY Date ASC) as rn
                FROM Access
                WHERE CAST(Date AS DATE) = '{target_date}'
                  AND CAST(Date AS TIME) BETWEEN '06:00:00' AND '09:10:00'
                  AND in_out = 'in'
            ) as SubQuery
            WHERE rn = 1
        """
        df_access = pd.read_sql(query, conn)
        conn.close()
        
        if df_access.empty:
            print("目前時段尚無入園資料，結束本次作業。")
            return
            
        # 資料清理與格式化
        df_access['EmpID'] = df_access['EmpID'].astype(str).str.strip()
        # 🌟 強制確保 CardID 是乾淨的字串，解決網頁抓不到的問題
        df_access['CardID'] = df_access['CardID'].astype(str).str.strip()
        
        # 格式化時間，讓試算表跟網頁都好讀
        df_access['EntryTime'] = pd.to_datetime(df_access['EntryTime']).dt.strftime('%Y-%m-%d %H:%M:%S')
        
        print(f"成功撈取 {len(df_access)} 筆入園名單 (含卡號與時間)。")
            
    except Exception as e:
        print(f"❌ 門禁 SQL 連線錯誤: {e}")
        return

    # 2. 上傳至 Google Sheets
    try:
        sheet = setup_gspread()
        # 清除舊資料，保持大腦乾淨
        sheet.clear()
        # 上傳資料與標題
        data_to_upload = [df_access.columns.values.tolist()] + df_access.values.tolist()
        sheet.update(values=data_to_upload, range_name='A1')
        print("✅ 成功將名單、卡號、姓名、卡機與時間上傳至雲端！")
    except Exception as e:
        print(f"❌ Google Sheets 寫入錯誤: {e}")

if __name__ == "__main__":
    main()