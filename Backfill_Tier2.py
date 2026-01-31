# -*- coding: utf-8 -*-
"""
Backfill_Tier2.py
משיכת נתונים יומיים מה-1 לחודש ועד אתמול ושמירתם בארכיון TIER2.
"""
import requests
import pandas as pd
from datetime import datetime, timedelta
import os
import time

# הגדרות (מועתק מ-TIER2.PY)
API_KEY = "a0bb0de4-2193-41c6-bff6-2f87344953ea"
API_SECRET = "ZWHRKYQNdHsX3HuoK27Xk6omQchnieko28iadd3qxTyxAVKMu1K54jLVsFNoa3nsJC1Ea4ajfg6zsAcIbQOit36B2urQCpGd4K6nkPeJmtixYSoP6ZMwTmCgWgQiVnLt"
EMAIL_FROM = "MosheI1@VERIFONE.com"
TIER2_DIR = os.path.join(os.getcwd(), "TIER2")

if not os.path.exists(TIER2_DIR):
    os.makedirs(TIER2_DIR)

def get_access_token():
    url = "https://verifone.glassix.com/api/v1.2/token/get"
    payload = {"apiKey": API_KEY, "apiSecret": API_SECRET, "userName": EMAIL_FROM}
    response = requests.post(url, json=payload, timeout=90)
    response.raise_for_status()
    return response.json().get("access_token")

def get_glassix_tickets(token, since, until):
    headers = {"Authorization": f"Bearer {token}"}
    tickets_all = []
    url = f"https://verifone.glassix.com/api/v1.2/tickets/list?since={since}&until={until}"
    
    while url:
        try:
            response = requests.get(url, headers=headers, timeout=90)
            if response.status_code == 429:
                print("Too many requests, waiting 60s...")
                time.sleep(60)
                continue
            response.raise_for_status()
            data = response.json()
            tickets = data.get("tickets", [])
            tickets_all.extend(tickets)
            paging = data.get("paging")
            url = paging.get("next") if paging and "next" in paging else None
        except Exception as e:
            print(f"Error: {e}")
            break
            
    return tickets_all

def main():
    print("="*60)
    print("   TIER 2 DATA BACKFILL TOOL")
    print("="*60)

    now = datetime.now()
    start_date = now.replace(day=1)
    # אתמול
    end_date = now - timedelta(days=1)

    if start_date > end_date:
        print("[!] היום ה-1 לחודש, אין נתונים קודמים לחודש זה למשיכה.")
        return

    print(f"[*] מתחיל משיכה מה-{start_date.strftime('%d/%m/%Y')} ועד ה-{end_date.strftime('%d/%m/%Y')}...")
    
    token = get_access_token()
    current_date = start_date
    
    while current_date <= end_date:
        date_str = current_date.strftime('%d/%m/%Y')
        safe_date = current_date.strftime('%d_%m_%Y')
        filename = f"Tickets_דוח_יומי_{safe_date}.xlsx"
        filepath = os.path.join(TIER2_DIR, filename)
        
        if os.path.exists(filepath):
            print(f"   [-] {date_str}: קובץ כבר קיים, מדלג.")
        else:
            print(f"   [*] {date_str}: שולף נתונים...")
            since = current_date.strftime("%d/%m/%Y 00:00:00:00")
            until = current_date.strftime("%d/%m/%Y 23:59:59:00")
            
            tickets = get_glassix_tickets(token, since, until)
            if tickets:
                pd.DataFrame(tickets).to_excel(filepath, index=False)
                print(f"   [V] {date_str}: נשמרו {len(tickets)} פניות.")
            else:
                print(f"   [!] {date_str}: לא נמצאו נתונים.")
        
        current_date += timedelta(days=1)

    print("\n[SUCCESS] תהליך ה-Backfill הסתיים!")

if __name__ == "__main__":
    main()
