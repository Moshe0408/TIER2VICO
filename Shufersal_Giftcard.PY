# -*- coding: utf-8 -*-
import requests
import base64
import os
import pandas as pd
from datetime import datetime, timedelta
import win32com.client as win32

# ---------------- CONFIG ----------------
ADIR_USER_ID = "9bc2c357-6261-48e3-8562-75674c2b1071"
ADIR_API_KEY = "itVBRDSsaPVYFDlvhKKDnNKRCeZxoOSPvKbO"

BASE_URL = "https://emea.gsc.verifone.cloud/oidc/report-engine/api/v1"

SAVE_DIRS = [
    r"K:\moshe VICO\MosheTools\Shufersal GBX - Giftcard",
    r"J:\VIRD\Shufersal GBX - Giftcard",
    r"C:\Users\Moshei1\OneDrive - Verifone\Desktop\TIP\Shufersal_Giftcard"
]

# ספרייה מרכזית לאיחוד
MAIN_DIR = SAVE_DIRS[-1]

# עמודות להסרה בקובץ החודשי (ללא Transaction_type / Transaction_state)
COLUMNS_TO_DROP_IN_MONTHLY = ["Transaction_time", "Created_in_UTC", "OwnerName", "DistributionDate"]

# ---------------- AUTH ----------------
auth_str = f"{ADIR_USER_ID}:{ADIR_API_KEY}"
b64_auth = base64.b64encode(auth_str.encode()).decode()
headers = {"Authorization": f"Basic {b64_auth}"}

# ---------------- HELPERS / MERGE LOGIC ----------------
def merge_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    מיזוג שורות לפי Merchant_reference + Masked_card_number + Current_amount
    או לפי שמות העמודות בעברית במקרה שהקבצים כבר עברו rename.
    """

    if df is None or df.empty:
        return df

    df = df.fillna('').astype(str).copy()
    
    # ניקוי שורות רווח או שורות שאין בהן מידע מינימלי
    # (שורות שבהן כל המפתחות ריקים הן "זבל")
    eng_keys = ['Merchant_reference', 'Masked_card_number', 'Current_amount']
    heb_keys = [
        COLUMNS_HEBREW_MAP.get('Merchant_reference', 'קוד סוחר'),
        COLUMNS_HEBREW_MAP.get('Masked_card_number', 'מספר כרטיס מוסתר'),
        COLUMNS_HEBREW_MAP.get('Current_amount', 'סכום נוכחי')
    ]

    if all(k in df.columns for k in eng_keys):
        keys = eng_keys
    elif all(k in df.columns for k in heb_keys):
        keys = heb_keys
    else:
        # אם אין עמודות מפתח, פשוט נחזיר את ה-df המקורי ללא שינוי (או נזרוק שגיאה)
        return df

    # נסנן שורות שבהן כל המפתחות ריקים - זה ה"זבל" הראשוני
    df = df[~(df[keys].apply(lambda x: x.str.strip()).eq('').all(axis=1))]

    df['_orig_idx'] = range(len(df))
    merged_rows = []

    # מיזוג קבוצות
    for _, group in df.groupby(keys, sort=False):
        # אם המפתח ריק (כלומר קיבלנו קבוצה של שורות ריקות למחצה שנותרו), 
        # אנחנו לא רוצים לאחד אותן לאחת אלא לשמור אותן או להשמיט
        # כאן נחליט להשמיט אותן אם אין להן Transaction_ID
        if str(group[keys[0]].iloc[0]).strip() == "" and \
           str(group[keys[1]].iloc[0]).strip() == "" and \
           str(group[keys[2]].iloc[0]).strip() == "":
            # השמטת שורות שכל המפתחות שלהן ריקים
            continue

        combined = group.sort_values('_orig_idx').iloc[0].copy()

        for _, r in group.sort_values('_orig_idx').iterrows():
            for col in df.columns:
                if col == '_orig_idx':
                    continue
                if str(combined.get(col, '')).strip() == "" and str(r.get(col, '')).strip() != "":
                    combined[col] = r[col]

        merged_rows.append(combined.drop(labels=['_orig_idx'], errors='ignore'))

    df_merged = pd.DataFrame(merged_rows)
    orig_cols = [c for c in df.columns if c != '_orig_idx']
    df_merged = df_merged.reindex(columns=orig_cols, fill_value='')

    return df_merged


# ---------------- FUNCTIONS ----------------
def last_report_id():
    url = f"{BASE_URL}/reports"
    params = {"pageNumber": 1, "pageSize": 1, "orderBy": "DESC", "orderCriteria": "createdOn"}
    resp = requests.get(url, headers=headers, params=params)
    resp.raise_for_status()
    data = resp.json()
    report_id = data["reports"][0]["reportUid"]
    print("Last Report ID:", report_id)
    return report_id

def download_report(report_id):
    url = f"{BASE_URL}/reports/{report_id}"
    params = {"download": "true"}
    resp = requests.get(url, headers=headers, params=params)
    resp.raise_for_status()

    today_str = datetime.today().strftime("%Y-%m-%d")
    file_name = f"Shufersal_Giftcard_Settled__{today_str}.csv"

    saved_paths = []
    for save_dir in SAVE_DIRS:
        try:
            if not os.path.exists(os.path.splitdrive(save_dir)[0]):
                print(f"⚠️ כונן לא זמין: {save_dir}, מדלג...")
                continue

            os.makedirs(save_dir, exist_ok=True)
            full_path = os.path.join(save_dir, file_name)
            with open(full_path, "wb") as f:
                f.write(resp.content)
            print(f"הדו\"ח היומי נשמר: {full_path}")
            saved_paths.append(full_path)
        except PermissionError:
            print(f"⚠️ אין הרשאה לכתוב לתיקייה: {save_dir}, מדלג...")

    if not saved_paths:
        raise Exception("❌ לא הצלחתי לשמור את הדו\"ח היומי באף תיקייה")

    return saved_paths[-1]

# סדר עמודות קבוע
COLUMNS_ORDER = [
    "Timezone", "Transaction_ID", "Product", "Current_amount", "Current_currency",
    "Transaction_type", "Transaction_state", "Channel", "Masked_card_number",
    "Settlement_date", "Settlement_ID", "Source_PIM", "Merchant_reference",
    "Merchant_ID", "CAVV_result", "3DS_XID", "3DS_authentication_value",
    "phone_number", "Transaction_time"
]

# מיפוי שמות עמודות לעברית בלבד
COLUMNS_HEBREW_MAP = {
    "Organization_name": "שם הארגון",
    "Created_datetime": "תאריך יצירה",
    "Timezone": "אזור זמן",
    "Transaction_ID": "מזהה עסקה",
    "Product": "מוצר",
    "Current_amount": "סכום נוכחי",
    "Current_currency": "מטבע נוכחי",
    "Transaction_type": "סוג עסקה",
    "Transaction_state": "סטטוס עסקה",
    "Channel": "ערוץ",
    "Masked_card_number": "מספר כרטיס מוסתר",
    "Settlement_date": "תאריך שידור",
    "Settlement_ID": "מזהה שידור",
    "Source_PIM": "מקור PIM",
    "Merchant_reference": "קוד סוחר",
    "Merchant_ID": "מזהה סוחר",
    "CAVV_result": "תוצאת CAVV",
    "3DS_XID": "3DS XID",
    "3DS_authentication_value": "ערך אימות 3DS",
    "phone_number": "מספר טלפון",
    "Transaction_time": "זמן עסקה"
}


def merge_report(csv_file):
    """מיזוג שורות לדו"ח היומי + יצירת phone_number ו-Transaction_time (+3 שעות)"""
    try:
        df = pd.read_csv(csv_file, dtype=str, encoding='utf-8-sig').fillna('')
    except Exception:
        df = pd.read_csv(csv_file, dtype=str, encoding='utf-8').fillna('')

    # ניקוי רווחים בעמודות קריטיות
    for col in ['Merchant_reference', 'Masked_card_number', 'Current_amount']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    merged_rows = []
    used_indices = set()

    for idx, row in df.iterrows():
        if idx in used_indices:
            continue

        mask = pd.Series([True] * len(df))
        if 'Merchant_reference' in df.columns:
            mask &= (df['Merchant_reference'] == row.get('Merchant_reference', ''))
        if 'Masked_card_number' in df.columns:
            mask &= (df['Masked_card_number'] == row.get('Masked_card_number', ''))
        if 'Current_amount' in df.columns:
            mask &= (df['Current_amount'] == row.get('Current_amount', ''))

        group = df[mask]

        if 'Transaction_type' in df.columns and 'Transaction_state' in df.columns:
            capture_settled = group[
                (group['Transaction_type'] == 'CAPTURE') &
                (group['Transaction_state'] == 'SETTLED')
            ]
        else:
            capture_settled = pd.DataFrame()

        combined_row = capture_settled.iloc[0].copy() if not capture_settled.empty else row.copy()

        for _, r in group.iterrows():
            for col in df.columns:
                if str(combined_row.get(col, '')).strip() == "" and str(r.get(col, '')).strip() != "":
                    combined_row[col] = r[col]

        merged_rows.append(combined_row)
        used_indices.update(group.index.tolist())

    df_merged = pd.DataFrame(merged_rows, columns=df.columns)
    
    # פילטור "זבל": נשמור רק שורות שיש להן מזהה עסקה וסכום
    if 'Transaction_ID' in df_merged.columns:
        df_merged = df_merged[df_merged['Transaction_ID'].str.strip() != '']
    if 'Current_amount' in df_merged.columns:
        df_merged = df_merged[df_merged['Current_amount'].str.strip() != '']

    if 'Merchant_reference' in df_merged.columns:
        df_merged['phone_number'] = df_merged['Merchant_reference'].str[:10]
    else:
        df_merged['phone_number'] = ""

    if 'Created_datetime' in df_merged.columns:
        dt_series = pd.to_datetime(df_merged['Created_datetime'], errors='coerce') + timedelta(hours=3)
        df_merged['Transaction_time'] = dt_series.dt.strftime('%Y-%m-%dT%H:%M:%S.%f').str[:-3] + 'Z'
    elif 'Created_in_UTC' in df_merged.columns:
        dt_series = pd.to_datetime(df_merged['Created_in_UTC'], errors='coerce') + timedelta(hours=3)
        df_merged['Transaction_time'] = dt_series.dt.strftime('%Y-%m-%dT%H:%M:%S.%f').str[:-3] + 'Z'
    else:
        df_merged['Transaction_time'] = "1970-01-01T00:00:00.000Z"

    df_merged['Transaction_type'] = 'CAPTURE'
    df_merged['Transaction_state'] = 'SETTLED'

    # שינוי שמות העמודות לעברית בלבד
    df_merged.rename(columns=COLUMNS_HEBREW_MAP, inplace=True)
    
    df_merged.to_csv(csv_file, index=False, encoding='utf-8-sig')
    print(f"הדו\"ח היומי המאוחד נשמר עם שמות עמודות בעברית בלבד: {csv_file}")
    return csv_file

def merge_all_reports_in_dir(directory, output_file):
    """מיזוג כל הדוחות היומיים של החודש הנוכחי בלבד לתוך קובץ חודשי מצטבר"""
    current_year = datetime.today().year
    current_month = datetime.today().month

    all_files = []
    for f in os.listdir(directory):
        if f.endswith(".csv") and f.startswith("Shufersal_Giftcard_Settled__"):
            try:
                date_str = f.split("__")[1].replace(".csv", "")
                file_date = datetime.strptime(date_str, "%Y-%m-%d")
                if file_date.year == current_year and file_date.month == current_month:
                    all_files.append(os.path.join(directory, f))
            except Exception:
                continue

    if not all_files:
        raise Exception(f"❌ לא נמצאו קבצים יומיים עבור {current_month}/{current_year} בתיקייה: {directory}")

    df_list = []
    for f in all_files:
        try:
            # ננסה לקרוא עם utf-8-sig כדי לטפל ב-BOM
            df = pd.read_csv(f, dtype=str, encoding='utf-8-sig').fillna('')
        except Exception:
            # fallback ל-windows-1255 אם נדרש
            df = pd.read_csv(f, dtype=str, encoding='windows-1255').fillna('')
        df_list.append(df)
        
    df_all = pd.concat(df_list, ignore_index=True)

    # נרמול שמות עמודות: אם הקובץ כבר עם עברית, נחזיר לאנגלית זמנית לצורך המיזוג
    # אנחנו צריכים שהמיזוג יעבוד לפי מפתח אחיד.
    
    # מיפוי הפוך: מעברית לאנגלית
    REVERSE_HEBREW_MAP = {v: k for k, v in COLUMNS_HEBREW_MAP.items()}
    
    # ננסה להמיר שמות עמודות באנגלית אם הם בעברית
    new_columns = []
    for col in df_all.columns:
        if col in REVERSE_HEBREW_MAP:
            new_columns.append(REVERSE_HEBREW_MAP[col])
        else:
            new_columns.append(col)
    
    df_all.columns = new_columns
    
    # כעת הפונקציה merge_dataframe תעבוד רגיל כי העמודות באנגלית
    df_merged = merge_dataframe(df_all)

    if 'Merchant_reference' in df_merged.columns:
        df_merged['phone_number'] = df_merged['Merchant_reference'].str[:10]
    else:
        df_merged['phone_number'] = ""

    if 'Created_datetime' in df_merged.columns:
        def fill_transaction_time(row):
            if not row.get('Transaction_time') or row['Transaction_time'].strip() == '':
                dt = pd.to_datetime(row['Created_datetime'], errors='coerce')
                if pd.notnull(dt):
                    return (dt + timedelta(hours=3)).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'
            return row.get('Transaction_time', '')
        df_merged['Transaction_time'] = df_merged.apply(fill_transaction_time, axis=1)

    elif 'Created_in_UTC' in df_merged.columns:
        def fill_transaction_time(row):
            if not row.get('Transaction_time') or row['Transaction_time'].strip() == '':
                dt = pd.to_datetime(row['Created_in_UTC'], errors='coerce')
                if pd.notnull(dt):
                    return (dt + timedelta(hours=3)).strftime('%Y-%m-%dT%H:%M:%S.%f')[:-3] + 'Z'
            return row.get('Transaction_time', '')
        df_merged['Transaction_time'] = df_merged.apply(fill_transaction_time, axis=1)
    else:
        df_merged['Transaction_time'] = "1970-01-01T00:00:00.000Z"

    df_merged['Transaction_type'] = 'CAPTURE'
    df_merged['Transaction_state'] = 'SETTLED'

    # השמטת עמודות שביקשת להסיר (אם הן קיימות)
    if COLUMNS_TO_DROP_IN_MONTHLY:
        df_merged.drop(columns=[c for c in COLUMNS_TO_DROP_IN_MONTHLY if c in df_merged.columns and c != 'Transaction_time'], inplace=True, errors='ignore')

    # פילטור "זבל" סופי לפני השמירה: רק שורות עם מזהה עסקה וסכום, וללא תאריך 1970
    if 'Transaction_ID' in df_merged.columns:
        df_merged = df_merged[df_merged['Transaction_ID'].str.strip() != '']
    if 'Current_amount' in df_merged.columns:
        df_merged = df_merged[df_merged['Current_amount'].str.strip() != '']
    if 'Transaction_time' in df_merged.columns:
        df_merged = df_merged[df_merged['Transaction_time'] != "1970-01-01T00:00:00.000Z"]

    df_merged = df_merged.reindex(columns=[c for c in COLUMNS_ORDER if c in df_merged.columns], fill_value='')

    # שינוי שמות העמודות לעברית בלבד
    df_merged.rename(columns=COLUMNS_HEBREW_MAP, inplace=True)

    df_merged.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(f"📂 קובץ חודשי מצטבר נשמר עם שמות עמודות בעברית בלבד: {output_file}  (שורות: {len(df_merged)})")
    return output_file


def send_email_with_attachments(to_recipients, subject, body, attachment_paths, cc_recipients=None):
    """שליחת מייל מקצועי ומודרני עם קבצים מצורפים"""
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = "; ".join(to_recipients)
    if cc_recipients:
        mail.CC = "; ".join(cc_recipients)
    mail.Subject = subject

    # ----- HTML מודרני ונקי -----
    body_html = f"""
    <html>
    <body style="
        direction: rtl; 
        text-align: right; 
        font-family: 'Segoe UI', Arial, sans-serif; 
        font-size: 15px; 
        color: #2c2c2c; 
        background-color: #ffffff;
    ">

        <!-- כותרת עליונה מודרנית -->
        <div style="
            background-color: #0066cc; 
            padding: 16px 22px; 
            color: white; 
            font-size: 18px; 
            font-weight: 600; 
            border-radius: 6px;
        ">
            דו"ח גיפטקארד שופרסל
        </div>

        <!-- ריווח -->
        <div style="height: 20px;"></div>

        <!-- תוכן -->
        <p style="margin: 0 0 14px 0;">
            שלום,
        </p>

        <p style="margin: 0 0 22px 0; line-height: 1.6;">
            מצורף דוח עסקאות גיפטקארד שופרסל – יומי ומצטבר חודשי.
        </p>

        <!-- חתימה -->
        <p style="margin-top: 28px; line-height: 1.6;">
            בברכה,<br>
            <span style="font-weight: 600; color: #0066cc;">Verifone Israel</span>
        </p>

    </body>
    </html>
    """

    mail.HTMLBody = body_html

    # הוספת קבצים
    for path in attachment_paths:
        if path and os.path.exists(path):
            mail.Attachments.Add(path)
            print(f"📎 מצורף למייל: {path}")
        else:
            print(f"⚠️ קובץ לא נמצא: {path}")

    mail.Send()
    print("📧 המייל נשלח בהצלחה.")


# ---------------- MAIN ----------------
if __name__ == "__main__":
    report_id = last_report_id()

    # הורדת הדו"ח היומי ושמירתו בתיקיות
    daily_file = download_report(report_id)

    # נריץ את מיזוג/ניקוי על הדו"ח היומי
    merge_report(daily_file)

    # ניצור/נעדכן את הקובץ החודשי המצטבר
    month_name = datetime.today().strftime("%B")  # למשל "September"
    monthly_file = os.path.join(MAIN_DIR, f"Shufersal 3DS - Cumulative_{month_name}.csv")
    merged_file = merge_all_reports_in_dir(MAIN_DIR, monthly_file)

    recipients = ["verifone3dslog@shufersal.co.il"]
    cc_recipients = [
        "nadav.lieber@verifone.com",
        "eli.nachmani@verifone.com",
        "moshei1@verifone.com",
        "shufersal@verifone.com"
    ]

    send_email_with_attachments(
        recipients,
        subject=f'דו\"ח גיפטקארד שופרסל {datetime.today().strftime("%Y-%m-%d")}',
        body="שלום ,\n\nמצורף דוח עסקאות גיפטקארד שופרסל - יומי ומצטבר חודשי.\n\nבברכה,\nוריפון",
        attachment_paths=[daily_file, merged_file],
        cc_recipients=cc_recipients
    )

