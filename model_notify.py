# -*- coding: utf-8 -*-
r"""
預測模型週報／切換通知 → LINE（2026-11-18）

GitHub 每天下午自動校準（scripts/calibrate_models.py）後，若有新報告（每週一的週報，或第一版方法自動切換時），
寫在 data/model_report.json。這支程式由筆電排程 CTIC\ModelNotify（平日 15:00、17:00）執行：
讀 GitHub 上最新的報告，還沒發過就用「中創餐廳通知」LINE 群發，發過的 id 記在 logs\model_notify_sent.txt。

用法：python model_notify.py            # 有新報告才發
      python model_notify.py --dry-run  # 只印出，不發
"""
import json
import os
import sys
import urllib.request

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
AUTO_DIR = r'C:\Users\User\CTIC_Dining_Auto'
REPORT_URL = 'https://raw.githubusercontent.com/tsaitzuchien-star/CTIC_Dining_Prediction/main/data/model_report.json'
SENT = os.path.join(BASE_DIR, 'logs', 'model_notify_sent.txt')


def main():
    dry = '--dry-run' in sys.argv
    try:
        import ssl
        ctx = None
        try:
            import truststore  # 院內網路會攔截 HTTPS：有裝 truststore 就用 Windows 憑證
            ctx = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        except Exception:
            pass
        with urllib.request.urlopen(REPORT_URL + '?t=' + str(os.getpid()), timeout=30, context=ctx) as r:
            rep = json.load(r)
    except Exception as e:
        print(f'讀不到模型報告：{e}')
        return 1
    rid = str(rep.get('id', ''))
    os.makedirs(os.path.dirname(SENT), exist_ok=True)
    sent = open(SENT, encoding='utf-8').read().split() if os.path.exists(SENT) else []
    if not rid or rid in sent:
        print(f'沒有新報告（最新 {rid}）')
        return 0
    text = rep.get('text', '')
    if dry:
        print(text)
        return 0
    sys.path.insert(0, AUTO_DIR)
    from line_push import send
    if send(text):
        with open(SENT, 'a', encoding='utf-8') as f:
            f.write(rid + '\n')
        return 0
    return 1


if __name__ == '__main__':
    sys.exit(main())
