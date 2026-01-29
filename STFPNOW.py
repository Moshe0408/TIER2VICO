import sys, io
sys.stdout.reconfigure(encoding='utf-8')
import threading
import requests
import json
import time
import re
import os
from datetime import datetime, timedelta
from tkinter import messagebox, Tk, Text, END, Scrollbar, VERTICAL, RIGHT, Y, LEFT, BOTH, Frame, Label
from tkinter import ttk
import win32com.client as win32  # לשליחת מייל אאוטלוק
import pythoncom  # <<< חשוב לתיקון שגיאת CoInitialize

# ----------------- פרטי התחברות -----------------
USER = "MosheI1"
PASSWORD = "Verifone!!2026"

ENTITY_IDS = [
    "PISR00011462","PISR00011083","PISR00011075","PISR00020836","PISR00026034",
    "PISR00021515","PISR00021614","PISR00021700","PISR00021867","PISR00021868",
    "PISR00022400","PISR00022028","PISR00021515","PISR00022227","PISR00022226",
    "PISR00022228","PISR00022229","PISR00022384","PISR00022958","PISR00022230",
    "PISR00022083","PISR00022501","PISR00022401","PISR00022374","PISR00022916",
    "PISR00022957","PISR00023049","PISR00023310","PISR00023491","PISR00023569",
    "PISR00023821","PISR00023973","PISR00024065","PISR00024208","PISR00024807",
    "PISR00024771","PISR00024293","PISR00024459","PISR00025087","PISR00024929",
    "PISR00022082","PISR00024846","PISR00025037","PISR00025568","PISR00025567",
    "PISR00025951","PISR00026034","PISR00026586","PISR00026585","PISR00026687",
    "PISR00027068","PISR00026920","PISR00026919","PISR00027129","PISR00027173",
    "PISR00027374","PISR00027452","PISR00027953","PISR00028363","PISR00028426",
    "PISR00028539","PISR00028540","PISR00028776","PISR00028796","PISR00028940",
    "PISR00024930","PISR00029166","PISR00029232","PISR00029306","PISR00029509",
    "PISR00029510","PISR00029511","PISR00029512","PISR00015162","PISR00012062",
    "PISR00028200",
]

# ----------------- הכנה לשמירת לוגים -----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(BASE_DIR, "logs_stf")
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)
log_filename = os.path.join(LOG_DIR, f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
log_file = open(log_filename, "a", encoding="utf-8")

def log_write(text_widget, message):
    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S] ")
    full_msg = timestamp + message
    text_widget.insert(END, full_msg + "\n")
    text_widget.see(END)
    log_file.write(full_msg + "\n")
    log_file.flush()
    os.fsync(log_file.fileno())

# ----------------- שליחת מייל מקצועי עם עיצוב מודרני -----------------
def send_outlook_email_modern(subject, records, to_recipients, cc_recipients=None, text_widget=None):
    """שליחת מייל - ה-CoInitialize כבר בוצע ב-Thread הראשי"""
    try:
        log_write(text_widget, "📧 מתחיל תהליך שליחת מייל...")
        outlook = win32.Dispatch('outlook.application')
        namespace = outlook.GetNamespace("MAPI")
        namespace.Logon()
        
        mail = outlook.CreateItem(0)
        mail.Subject = subject

        # הוספת נמענים בצורה בטוחה
        to_list = to_recipients if isinstance(to_recipients, list) else [to_recipients]
        for addr in to_list:
            if addr:
                recipient = mail.Recipients.Add(addr)
                recipient.Type = 1 # olTo

        if cc_recipients:
            cc_list = cc_recipients if isinstance(cc_recipients, list) else [cc_recipients]
            for addr in cc_list:
                if addr:
                    recipient = mail.Recipients.Add(addr)
                    recipient.Type = 2 # olCC

        mail.Recipients.ResolveAll()

        # ----- HTML מודרני ונקי עם מספרים מימין לשמאל -----
        if records:
            total_sum = sum(rec.get('totalValue', 0) for rec in records)
            table_rows = ""
            for rec in records:
                branch_name = rec.get("branch_name", "Unknown")
                total_value = f"₪{rec.get('totalValue',0):,.2f}"
                table_rows += f"""
                <tr>
                    <td style="padding:6px 12px;border:1px solid #ccc;">{rec.get('site')}</td>
                    <td style="padding:6px 12px;border:1px solid #ccc;">{rec.get('account')}</td>
                    <td style="padding:6px 12px;border:1px solid #ccc;">{branch_name}</td>
                    <td style="padding:6px 12px;border:1px solid #ccc;">{rec.get('shiftId',0)}</td>
                    <td style="padding:6px 12px;border:1px solid #ccc;text-align:right; direction:ltr;">{total_value}</td>
                </tr>
                """
            body_html = f"""
            <html>
            <body style="direction: rtl; text-align: right; font-family: 'Segoe UI', Arial, sans-serif; font-size: 14px; color: #2c2c2c; background-color: #ffffff;">
                <div style="background-color: #0066cc; padding: 16px 22px; color: white; font-size: 18px; font-weight: 600; border-radius: 6px;">
                    דו"ח סניפים עם STFP נכשל
                </div>
                <div style="height: 20px;"></div>
                <p>שלום,</p>
                <p>להלן סניפים בהם העסקאות נכשלו (STFP):</p>
                <table style="border-collapse: collapse; width: 100%; margin-top: 10px;">
                    <thead style="background-color: #f0f0f0;">
                        <tr>
                            <th style="padding:8px 12px;border:1px solid #ccc;">SITE</th>
                            <th style="padding:8px 12px;border:1px solid #ccc;">ACCOUNT</th>
                            <th style="padding:8px 12px;border:1px solid #ccc;">Branch</th>
                            <th style="padding:8px 12px;border:1px solid #ccc;">Shift ID</th>
                            <th style="padding:8px 12px;border:1px solid #ccc;">Total</th>
                        </tr>
                    </thead>
                    <tbody>
                        {table_rows}
                        <tr style="font-weight:bold;">
                            <td colspan="4" style="padding:8px 12px;border:1px solid #ccc;">סכום כולל</td>
                            <td style="padding:8px 12px;border:1px solid #ccc;text-align:right; direction:ltr;">₪{total_sum:,.2f}</td>
                        </tr>
                    </tbody>
                </table>
                <p style="margin-top: 28px; line-height: 1.6;">
                    בברכה,<br>
                    <span style="font-weight: 600; color: #0066cc;">Verifone Israel</span>
                </p>
            </body>
            </html>
            """
        else:
            # במקרה שאין רשומות
            body_html = """
            <html>
            <body style="direction: rtl; text-align: right; font-family: 'Segoe UI', Arial, sans-serif; font-size: 14px; color: #2c2c2c; background-color: #ffffff;">
                <div style="background-color: #0066cc; padding: 16px 22px; color: white; font-size: 18px; font-weight: 600; border-radius: 6px;">
                    דו"ח סניפים עם STFP נכשל
                </div>
                <div style="height: 20px;"></div>
                <p>שלום,</p>
                <p>כל הסניפים תקינים, אין עסקאות שנכשלו.</p>
                <p style="margin-top: 28px; line-height: 1.6;">
                    בברכה,<br>
                    <span style="font-weight: 600; color: #0066cc;">Verifone Israel</span>
                </p>
            </body>
            </html>
            """

        mail.HTMLBody = body_html
        try:
            mail.Save()
            mail.Send()
            log_write(text_widget, "✅ המייל נשלח בהצלחה!")
        except Exception as e:
            log_write(text_widget, f"⚠️ שגיאה בשלב שליחת המייל: {e}")
        
        # ניסיון "לדחוף" את אאוטלוק לשלוח מה-Outbox
        try:
            for sync in namespace.SyncObjects:
                sync.Start()
        except:
            pass
            
        log_write(text_widget, "⌛ ממתין 15 שניות לסיום שידור...")
        time.sleep(15)
            
    except Exception as e:
        log_write(text_widget, f"❌ שגיאה בבניית המייל: {e}")
        raise e



# ----------------- פונקציות API -----------------
def get_token(user_id, password, text_widget):
    url = "https://us.vfmerchantportal.com/gmp-web/rest/public/security/login"
    headers = {"Content-Type": "application/json"}
    payload = {"userId": user_id, "password": password}
    resp = requests.post(url, headers=headers, json=payload)
    if resp.status_code != 200:
        raise Exception(f"❌ Login failed: {resp.status_code}")
    data = resp.json()
    log_write(text_widget, "✅ Token received")
    return data["bearerToken"], resp.cookies.get_dict()

def get_failed_batches(entity_id, token, cookies, text_widget):
    url = "https://us.vfmerchantportal.com/gmp-web/rest/settlement/execution"
    headers = {"Authorization": f"Bearer {token}"}
    today = datetime.today()
    yesterday = today - timedelta(days=1)
    params = {
        "dateFrom": yesterday.strftime("%Y%m%d"),
        "dateTo": today.strftime("%Y%m%d"),
        "entityId": entity_id,
        "entityType": "MERCHANT",
        "submissionFlag": "STFP_FAILED",
        "pageSize": 100
    }
    resp = requests.get(url, headers=headers, params=params, cookies=cookies)
    log_write(text_widget, f"[GET BATCHES] {resp.status_code} - {entity_id}")
    if resp.status_code != 200:
        return []
    batches = resp.json().get("list") or resp.json().get("results") or []
    result = []
    for b in batches:
        result.append({
            "site": b.get("entityId", "Unknown"),
            "account": entity_id,
            "totalValue": b.get("totalValue", 0),
            "shiftId": b.get("shiftId", 0),
            "branch_name": b.get("branchName", "Unknown")
        })
    return result

def get_terminals(branch_id, token, cookies, text_widget):
    url = "https://us.vfmerchantportal.com/gmp-web/rest/terminal/query"
    headers = {"Authorization": f"Bearer {token}"}
    terminals = []
    for devtype in ["backofficeserver", "P400 Plus"]:
        params = {
            "branchId": branch_id,
            "deviceTypeId": devtype,
            "firstRow": "1",
            "pageSize": "4",
            "appType": "UNDEFINED"
        }
        resp = requests.get(url, headers=headers, params=params, cookies=cookies)
        log_write(text_widget, f"[GET TERMINALS] {resp.status_code} - Branch: {branch_id} - Devtype: {devtype}")
        if resp.status_code != 200:
            continue
        data = resp.json().get("list") or resp.json().get("results") or []
        for t in data:
            terminals.append({
                "site": branch_id,
                "term": t["terminalId"],
                "devtype": devtype
            })
    return terminals

# ----------------- שליחת שידור -----------------
def send_settlement(row, text_widget):
    url = "https://us.vpaasgateway.com/weaver-fim-ugp-ndata/UGPHttpServlet"
    headers = {"Content-Type":"application/xml","ugp-version":"3.14","Host":"us.vpaasgateway.com"}
    log_write(text_widget, f"\n🔄 Sending settlement on {row['term']} (DEVTYPE={row['devtype']})")

    # SETUP
    setup_body = f"""
<UGPREQUEST>
<FUNCTION_TYPE>ADMIN</FUNCTION_TYPE>
<COMMAND>SETUP</COMMAND>
<DEVTYPE>{row['devtype']}</DEVTYPE>
<TERM>{row['term']}</TERM>
<SERIAL_NUM>111-222-333</SERIAL_NUM>
</UGPREQUEST>
"""
    resp1 = requests.post(url, headers=headers, data=setup_body)
    log_write(text_widget, f"📤 Setup Response: {resp1.status_code}")
    match = re.search(r"<DEVICEKEY>(.*?)</DEVICEKEY>", resp1.text)
    if not match:
        log_write(text_widget, "❌ Could not find DEVICEKEY")
        return resp1.status_code
    device_key = match.group(1)
    log_write(text_widget, f"🔑 DEVICEKEY: {device_key}")

    # SETTLE
    settle_body = f"""
<UGPREQUEST VER='1.0'>
<FUNCTION_TYPE>BATCH</FUNCTION_TYPE>
<COMMAND>SETTLE</COMMAND>
<ACCOUNT>{row['account']}</ACCOUNT>
<SITE>{row['site']}</SITE>
<TERM>{row['term']}</TERM>
<SERIAL_NUM>111-222-333</SERIAL_NUM>
<DEVTYPE>{row['devtype']}</DEVTYPE>
<DEVICEKEY>{device_key}</DEVICEKEY>
<PROCESSOR_ID>ISRAEL-ABS</PROCESSOR_ID>
<SCH_SYNC_FLAG>TRUE</SCH_SYNC_FLAG>
<SETTLEMENT_LEVEL>BRANCH</SETTLEMENT_LEVEL>
<SHIFT_ID>0</SHIFT_ID>
</UGPREQUEST>
"""
    resp2 = requests.post(url, headers=headers, data=settle_body)
    log_write(text_widget, f"✅ Settlement Response: {resp2.status_code}")
    return resp2.status_code

# ----------------- ביצוע השידור לכל הרשומות -----------------
# ----------------- ביצוע השידור עם עדכון כמות וסכום -----------------
def perform_settlement(records, text_widget, progressbar, percent_label, count_label, sum_label):
    total = len(records)
    sent_count = 0
    total_sum = 0

    for i, row in enumerate(records, start=1):
        try:
            status_code = send_settlement(row, text_widget)
            if status_code == 200:
                sent_count += 1
                total_sum += row.get('totalValue', 0)
        except Exception as e:
            log_write(text_widget, f"❌ Error sending settlement: {str(e)}")

        # עדכון פרוגרס בר
        progress = int((i / total) * 100)
        progressbar['value'] = progress
        percent_label.config(text=f"{progress}%")

        # עדכון כמות וסכום עסקאות
        count_label.config(text=f"עסקאות שנשלחו: {sent_count}/{total}")
        sum_label.config(text=f"סכום עסקאות: ₪{total_sum:,.2f}")

        progressbar.update_idletasks()

    percent_label.config(text="100%")
    progressbar['value'] = 100
    count_label.config(text=f"עסקאות שנשלחו: {sent_count}/{total}")
    sum_label.config(text=f"סכום עסקאות: ₪{total_sum:,.2f}")
    progressbar.update_idletasks()

# ----------------- MAIN -----------------
def main(text_widget, progressbar, percent_label, count_label, sum_label):
    try:
        token, cookies = get_token(USER, PASSWORD, text_widget)
        all_records = []
        seen_pairs = set()
        for account in ENTITY_IDS:
            batches = get_failed_batches(account, token, cookies, text_widget)
            for batch in batches:
                branch_id = batch.get("site")
                terminals = get_terminals(branch_id, token, cookies, text_widget)
                for term in terminals:
                    key = (term["site"], account)
                    if key not in seen_pairs:
                        seen_pairs.add(key)
                        all_records.append({
                            "site": term["site"],
                            "term": term["term"],
                            "account": account,
                            "devtype": term["devtype"],
                            "totalValue": batch.get("totalValue",0),
                            "shiftId": batch.get("shiftId",0),
                            "branch_name": batch.get("branch_name","Unknown")
                        })
        if all_records:
            log_write(text_widget, f"🔍 Found {len(all_records)} terminals for settlement")
            perform_settlement(all_records, text_widget, progressbar, percent_label, count_label, sum_label)
            send_outlook_email_modern(
                subject="דו\"ח סניפים עם STFP נכשל",
                records=all_records,
                to_recipients=["i.verticals1@verifone.com"],
                cc_recipients=["moshe.isakov@verifone.com", "nadav.lieber@verifone.com"],
                text_widget=text_widget
            )
            messagebox.showinfo("Finished", f"✅ Settlement completed.\nEmail sent.\nLog saved to:\n{log_filename}")
        else:
            send_outlook_email_modern(
                subject="דו\"ח סניפים עם STFP נכשל - אין בעיות",
                records=[],
                to_recipients=["moshe.isakov@verifone.com"],
                cc_recipients=["i.verticals1@verifone.com", "nadav.lieber@verifone.com"],
                text_widget=text_widget
            )
            messagebox.showinfo("No Data", "⚠️ No terminals found. Email sent.")
    except Exception as e:
        log_write(text_widget, f"❌ Fatal error: {str(e)}")
        messagebox.showerror("Error", str(e))

# ----------------- GUI -----------------
def run_gui():
    root = Tk()
    root.title("Vfmerchantportal Settlement")
    root.geometry("800x650")
    root.configure(bg="#2c3e50")

    # כותרת עליונה
    Label(root, text="Moshe Systems", font=("Arial",30,"bold"), bg="#2c3e50", fg="white").pack(pady=15)

    # מסגרת עם טקסט ולוג
    frame = Frame(root)
    frame.pack(fill=BOTH, expand=True, padx=10, pady=5)
    scrollbar = Scrollbar(frame, orient=VERTICAL)
    scrollbar.pack(side=RIGHT, fill=Y)
    text_widget = Text(frame, yscrollcommand=scrollbar.set, wrap="word", height=20, width=90, bg="#34495e", fg="white", font=("Consolas",10))
    text_widget.pack(side=LEFT, fill=BOTH, expand=True)
    scrollbar.config(command=text_widget.yview)

    # פרוגרס בר
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("green.Horizontal.TProgressbar", foreground='#27ae60', background='#27ae60')
    progressbar = ttk.Progressbar(root, style="green.Horizontal.TProgressbar", orient="horizontal", length=700, mode="determinate", maximum=100)
    progressbar.pack(pady=10)
    percent_label = Label(root, text="0%", font=("Arial",14), bg="#2c3e50", fg="white")
    percent_label.pack()

    # Labels לכמות עסקאות וסכום
    count_label = Label(root, text="עסקאות שנשלחו: 0/0", font=("Arial",12), bg="#2c3e50", fg="white")
    count_label.pack()
    sum_label = Label(root, text="סכום עסקאות: ₪0.00", font=("Arial",12), bg="#2c3e50", fg="white")
    sum_label.pack(pady=(0,10))

    # הפעלת השידור ב-thread נפרד
    def start_thread():
        threading.Thread(target=threaded_main, args=(text_widget, progressbar, percent_label, count_label, sum_label), daemon=True).start()

    def threaded_main(text_widget, progressbar, percent_label, count_label, sum_label):
        pythoncom.CoInitialize()
        try:
            main(text_widget, progressbar, percent_label, count_label, sum_label)
        finally:
            pythoncom.CoUninitialize()

    root.after(1000, start_thread)
    root.mainloop()

if __name__ == "__main__":
    run_gui()
