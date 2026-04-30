import pandas as pd
import pyodbc
from datetime import datetime

# ==========================================
# 1. 資料庫連線設定區
# ==========================================
# [A] 門禁主機 (PUMS) 連線設定
# 注意：若不在同一台電腦，通常需要輸入帳號密碼 (SQL Server 驗證)
pums_conn_str = (
    r'DRIVER={SQL Server};'
    r'SERVER=請填入門禁主機的IP或名稱;'  # 例如: 192.168.1.100
    r'DATABASE=PUMS;'
    r'UID=請填入帳號;'                 # 例如: sa
    r'PWD=請填入密碼'                  
)

# [B] 本機餐廳 (Dining) 連線設定 (沿用我們之前成功的設定)
dining_conn_str = (
    r'DRIVER={SQL Server};'
    r'SERVER=14-0A00035-93\SQLEXPRESS;'
    r'DATABASE=Dining;'
    r'Trusted_Connection=yes;'
)

def get_today_prediction():
    print("🔄 正在連線撈取最新資料...\n")
    
    # 取得今天的日期 (為了方便你用 4/29 的資料測試，我先寫死，上線時再改成今天)
    # 正式上線時請改成: target_date = datetime.now().strftime('%Y-%m-%d')
    target_date = '2026-04-29' 

    # ==========================================
    # 2. 撈取門禁系統：今天 9:00 前入園的名單
    # ==========================================
    try:
        pums_conn = pyodbc.connect(pums_conn_str)
        # SQL 語法：抓取特定日期、9:00前、且動作為 in 的不重複 EmpID
        access_query = f"""
            SELECT DISTINCT EmpID
            FROM Access
            WHERE CAST(Date AS DATE) = '{target_date}'
              AND CAST(Date AS TIME) <= '09:00:00'
              AND in_out = 'in'
        """
        df_access = pd.read_sql(access_query, pums_conn)
        pums_conn.close()
        
        # 清除可能多餘的空白字元
        df_access['EmpID'] = df_access['EmpID'].astype(str).str.strip()
        
    except Exception as e:
        print(f"❌ 門禁主機 (PUMS) 連線失敗，請檢查 IP 或帳密: {e}")
        return

    # ==========================================
    # 3. 撈取本機系統：有效的常客名單
    # ==========================================
    try:
        dining_conn = pyodbc.connect(dining_conn_str)
        # 只要抓 IsActive = 1 的有效常客
        regular_query = "SELECT EmployeeID, Name, CompanyType FROM RegularDiners WHERE IsActive = 1"
        df_regulars = pd.read_sql(regular_query, dining_conn)
        dining_conn.close()
        
        df_regulars['EmployeeID'] = df_regulars['EmployeeID'].astype(str).str.strip()

    except Exception as e:
        print(f"❌ 本機 (Dining) 連線失敗: {e}")
        return

    # ==========================================
    # 4. 進行交集比對，算出預測人數！
    # ==========================================
    # 將兩張表依照員工編號 (EmpID = EmployeeID) 結合
    df_predicted = pd.merge(df_access, df_regulars, left_on='EmpID', right_on='EmployeeID', how='inner')

    # 計算總人數
    predicted_count = len(df_predicted)

    # ==========================================
    # 5. 輸出儀表板需要的結果
    # ==========================================
    print("=" * 40)
    print(f"🍽️ 【戰情室即時預測】 統計日期：{target_date}")
    print("=" * 40)
    print(f"總計入園常客數 ➡️  {predicted_count} 人")
    print("-" * 40)
    
    if predicted_count > 0:
        print("預計用餐名單預覽：")
        print(df_predicted[['CompanyType', 'Name', 'EmployeeID']].head(15).to_string(index=False))
    else:
        print("目前尚無常客入園。")

if __name__ == "__main__":
    get_today_prediction()