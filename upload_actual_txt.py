import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os
import shutil  # 🌟 新增：用來移動檔案的模組
import warnings

warnings.filterwarnings('ignore', category=UserWarning)

# ==========================================
# 1. 設定與 Google Sheets 連線
# ==========================================
CREDENTIALS_FILE = 'credentials.json' 
SPREADSHEET_NAME = 'CTIC_Access_Log'
WORKSHEET_NAME = '實際用餐紀錄'  # 🎯 丟到我們剛建好的新分頁

def setup_gspread():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    return client.open(SPREADSHEET_NAME).worksheet(WORKSHEET_NAME)

def main():
    # 自動產生今天的檔名，例如 "20260430.TXT"
    today_str = datetime.now().strftime("%Y%m%d")
    txt_filename = f"{today_str}.TXT"
    
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 準備讀取實際用餐名單: {txt_filename}")

    # 檢查檔案是否存在
    if not os.path.exists(txt_filename):
        print(f"❌ 找不到今天的檔案 {txt_filename}，請確認卡機資料已放入資料夾！")
        return

    try:
        # 讀取沒有標題列的 TXT 檔，並自己給它標題
        df = pd.read_csv(txt_filename, header=None, names=['Date', 'Time', 'CardID', 'Status'])
        
        # 確保卡號是乾淨的字串
        df['CardID'] = df['CardID'].astype(str).str.strip()
        print(f"成功讀取 {len(df)} 筆實際刷卡紀錄。")
        
    except Exception as e:
        print(f"❌ 讀取 TXT 檔案失敗: {e}")
        return

    # 上傳到 Google Sheets
    try:
        sheet = setup_gspread()
        sheet.clear() # 清空昨天的舊資料
        
        # 轉成清單格式上傳
        data_to_upload = [df.columns.values.tolist()] + df.values.tolist()
        sheet.update(values=data_to_upload, range_name='A1')
        print(f"✅ 成功將 {txt_filename} 的資料拋轉至 Google 試算表「{WORKSHEET_NAME}」！")
        
        # ==========================================
        # 🌟 自動歸檔機制 (移至 OK 資料夾)
        # ==========================================
        ok_folder = "OK"
        # 1. 如果沒有 OK 資料夾，就自動建立一個
        if not os.path.exists(ok_folder):
            os.makedirs(ok_folder)
            
        # 2. 設定目的地路徑
        destination = os.path.join(ok_folder, txt_filename)
        
        # 3. 如果 OK 資料夾裡面已經有同名檔案(例如今天重複測試)，先刪除舊的
        if os.path.exists(destination):
            os.remove(destination)
            
        # 4. 把剛剛上傳完的 TXT 移進去
        shutil.move(txt_filename, destination)
        print(f"📂 檔案已成功移動至 {ok_folder} 資料夾歸檔備查！")

    except Exception as e:
        print(f"❌ 寫入試算表或移動檔案時發生錯誤: {e}")

if __name__ == "__main__":
    main()