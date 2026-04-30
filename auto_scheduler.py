import time
from datetime import datetime
import subprocess
import sys

# 🎯 設定要呼叫的目標程式 (你原本寫好的完美上傳程式)
TARGET_SCRIPT = "upload_access_to_sheets.py"

print("=====================================================")
print("  🤖 中創園區 - 早盤自動上傳排程引擎已啟動  ")
print("=====================================================")
print("  ⏰ 監控區間：每天早上 06:00 到 09:30")
print("  ⏳ 執行頻率：每 5 分鐘自動拋轉一次")
print("  ⚠️ 注意：請保持此黑色視窗開啟，不要關閉喔！")
print("=====================================================\n")

while True:
    now = datetime.now()
    
    # 定義今天的開始與結束時間
    start_time = now.replace(hour=6, minute=0, second=0, microsecond=0)
    end_time = now.replace(hour=9, minute=30, second=0, microsecond=0)

    # 判斷現在是否在「06:00 ~ 09:30」的區間內
    if start_time <= now <= end_time:
        print(f"[{now.strftime('%H:%M:%S')}] 🟢 時間觸發！開始啟動資料拋轉...")
        
        try:
            # 呼叫你原本的程式執行 (使用 sys.executable 確保用同一個 Python 環境)
            subprocess.run([sys.executable, TARGET_SCRIPT], check=True)
            print(f"[{datetime.now().strftime('%H:%M:%S')}] ✅ 本次任務完成，進入 5 分鐘休眠...\n")
        except Exception as e:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] ❌ 執行目標程式時發生錯誤: {e}\n")
            
        # 休息 5 分鐘 (300秒) 後再進入下一次迴圈
        time.sleep(300)
        
    else:
        # 如果不在設定的時間內 (例如下午或半夜)，就什麼都不做
        # 讓程式輕度休眠 1 分鐘後再醒來檢查一次手錶，完全不吃電腦效能
        time.sleep(60)