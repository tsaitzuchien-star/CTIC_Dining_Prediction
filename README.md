# CTIC_Dining_Prediction
中創園區 - 用餐人數自動預測與戰情室系統

## 預測模型全自動校準（2026-11-18 起）

- `scripts/calibrate_models.py` 由 GitHub Actions（`.github/workflows/calibrate.yml`）平日 14:45、16:30 自動執行。
- 逐日只用以前的資料回測各方法（近 8 次同星期中位數、08:20 已到常客模型、兩者混合、09:20 進園推估、App 建議），
  結果在 `data/model_daily.csv`；目前採用的方法與係數在 `data/model_params.json`，戰情室與訂餐 App 都讀這個檔。
- 第一版方法：挑戰者近 10 個工作日平均誤差少 1 份以上才自動切換，反過來也會自動切回；出席率每天重校；
  混合比重與 App 進園比重看最近 20 個工作日；連假前一天係數滿 3 次且確實降低誤差才啟用（國定假日抓政府行事曆）。
- 每週一或切換時寫 `data/model_report.json`，筆電排程 `CTIC\ModelNotify`（`model_notify.py`）讀到新報告就 LINE 通知。
