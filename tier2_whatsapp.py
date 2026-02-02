# -*- coding: utf-8 -*-
import os
import sys
import time
import datetime
import traceback
import requests
import pandas as pd
import pytz
import json
import pyperclip
from collections import defaultdict

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# הגדרת קידוד לעברית בטרמינל
if sys.platform == "win32":
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# ----------------- CONFIG -----------------
API_KEY = "a0bb0de4-2193-41c6-bff6-2f87344953ea"
API_SECRET = "ZWHRKYQNdHsX3HuoK27Xk6omQchnieko28iadd3qxTyxAVKMu1K54jLVsFNoa3nsJC1Ea4ajfg6zsAcIbQOit36B2urQCpGd4K6nkPeJmtixYSoP6ZMwTmCgWgQiVnLt"
EMAIL_FROM = "MosheI1@VERIFONE.com"

REPORTS_GROUP = "דוחות Tier2"
BOT_GROUP = "בוט טייר 2"

TIER2_MAP = {
    "niv.arieli": "NivA2",
    "danv1": "DanV1",
    "liorb5": "LiorB5",
    "moshei1": "MosheI1",
    "Isakov, Moshe": "MosheI1",
    "Vaysman, Dan": "DanV1",
    "Benjamin, Lior": "LiorB5",
    "Arieli, Niv": "NivA2"
}

SLA_LIMIT_HOURS = 1.5
CHECK_INTERVAL_SECONDS = 90      # מרווח בדיקה
POST_ALERT_SLEEP = 90            # המתנה אחרי התראת SLA

# עיכובים בין ניסיונות שליחה (בשניות)
RETRY_DELAYS = [10, 20, 60, 70, 90]

# זיכרון לקריאות שכבר קיבלו התראת SLA
alerted_ticket_ids = set()

# שימוש בפרופיל נפרד לחלוטין עבור הוואטסאפ כדי שלא יתנגש עם Verint
CHROME_USER_DATA_DIR = os.path.join(os.getcwd(), "AutomationProfile_WA")
REAL_PROFILE_PATH = CHROME_USER_DATA_DIR
VERINT_REPORTS_DIR = os.path.join(os.getcwd(), "Verint_Reports")
WA_DEBUG_PORT = 9223 # פורט נפרד עבור וואטסאפ



DESKTOP = os.path.join(os.environ.get('USERPROFILE', '.'), 'Desktop')
LOG_DIR = os.path.join(DESKTOP, 'tier2_logs')
os.makedirs(LOG_DIR, exist_ok=True)
LOG_FILE = os.path.join(LOG_DIR, "tier2_monitor.log")

TOKEN_CACHE_FILE = os.path.join(LOG_DIR, "glassix_token_cache.json")
TOKEN_EXPIRY_SECONDS = 3600  # 1 שעה

global_token = None

# ---------- LOGGING ----------
def log(msg, also_print=True):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        # אל תעצור לולאה בגלל רישום לוג
        pass
    if also_print:
        print(line)

def log_exc(context=""):
    etype, evalue, etb = sys.exc_info()
    tb_str = "".join(traceback.format_exception(etype, evalue, etb))
    log(f"שגיאה: {context}\n{tb_str}", also_print=True)

# ---------- TOKEN HANDLING ----------
def save_token(token):
    try:
        data = {"token": token, "timestamp": time.time()}
        with open(TOKEN_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        log_exc("שמירת טוקן")

def load_token():
    try:
        if not os.path.exists(TOKEN_CACHE_FILE):
            return None
        with open(TOKEN_CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        token = data.get("token")
        timestamp = data.get("timestamp", 0)
        if token and (time.time() - timestamp) < TOKEN_EXPIRY_SECONDS:
            return token
    except Exception:
        log_exc("טעינת טוקן")
    return None

def get_access_token():
    url = "https://verifone.glassix.com/api/v1.2/token/get"
    payload = {"apiKey": API_KEY, "apiSecret": API_SECRET, "userName": EMAIL_FROM}
    try:
        r = requests.post(url, json=payload, timeout=30)
        r.raise_for_status()
        data = r.json()
        token = data.get("access_token")
        if not token:
            log(f"תגובה חסרה access_token: {data}")
            raise RuntimeError("לא התקבל access_token מה־Glassix")
        log("Access token התקבל בהצלחה.")
        return token
    except requests.exceptions.HTTPError as e:
        text = ""
        try:
            text = r.text
        except Exception:
            pass
        log(f"HTTPError {getattr(r, 'status_code', '?')}: {text}")
        raise
    except Exception:
        log_exc("קבלת טוקן")
        raise

def get_token():
    token = load_token()
    if token:
        return token
    token = get_access_token()
    save_token(token)
    return token
    
# ---------- TOKEN REFRESH LOOP ----------
import threading

# משתנה גלובלי שישמור תמיד את הטוקן האחרון
global_token = None  

def refresh_token_loop():
    global global_token
    while True:
        try:
            global_token = get_token()
            log("🔑 טוקן חודש בהצלחה")
        except Exception:
            log_exc("שגיאה בחידוש טוקן")
        time.sleep(3 * 60 * 60)  # ריענון כל 3 שעות

# ---------- TIME HELPERS ----------
def ensure_utc(dt: datetime.datetime):
    try:
        if dt is None:
            return None
        if isinstance(dt, pd.Timestamp):
            dt = dt.to_pydatetime()
        if dt.tzinfo is None:
            local_tz = pytz.timezone("Asia/Jerusalem")
            dt_localized = local_tz.localize(dt)
            dt_utc = dt_localized.astimezone(pytz.utc)
            return dt_utc.replace(tzinfo=None)
        else:
            return dt.astimezone(pytz.utc).replace(tzinfo=None)
    except Exception:
        return None

def to_utc_dt(value):
    """המרת ערך לתאריך/שעה ב-UTC כ- datetime (עם tzinfo=UTC)."""
    try:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None
        if isinstance(value, datetime.datetime):
            if value.tzinfo is None:
                return value.replace(tzinfo=pytz.UTC)
            return value.astimezone(pytz.UTC)
        # החלפת Z ל-UTC
        if isinstance(value, str):
            s = value.strip().replace("Z", "+00:00")
            try:
                dt = datetime.datetime.fromisoformat(s)
            except Exception:
                dt = pd.to_datetime(s, utc=True).to_pydatetime()
        else:
            dt = pd.to_datetime(value, utc=True).to_pydatetime()
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=pytz.UTC)
        return dt.astimezone(pytz.UTC)
    except Exception:
        return None
# ---------- API ----------
def get_tickets(token, since=None, until=None, states=None, retries=3, max_wait_time=600):
    """
    שולף כרטיסים מ-Glassix לפי טווח זמנים ומצב (ticketState).
    מטפל ב-paging עד שכל הכרטיסים יובאו.
    """
    headers = {"Authorization": f"Bearer {token}"}
    now = datetime.datetime.now(datetime.timezone.utc)


    # הכנה של טווח הזמנים (UTC)
    since_dt = ensure_utc(since or (now - datetime.timedelta(days=7)))
    until_dt = ensure_utc(until or now)
    if since_dt and until_dt and since_dt > until_dt:
        since_dt, until_dt = until_dt, since_dt

    since_str = (since_dt or now - datetime.timedelta(days=7)).strftime("%d/%m/%Y %H:%M:%S") + ":00"
    until_str = (until_dt or now).strftime("%d/%m/%Y %H:%M:%S") + ":00"

    base_url = "https://verifone.glassix.com/api/v1.2/tickets/list"

    params = {
        "since": since_str,
        "until": until_str
    }

    # שימוש בפרמטר הנכון לפי הדוקומנטציה: ticketState
    if states:
        try:
            params["ticketState"] = ",".join(states)
        except Exception:
            log_exc("בעיה עם states")

    tickets = []
    url = base_url
    attempt = 0
    wait_time = 60

    while True:
        attempt += 1
        try:
            r = requests.get(url, headers=headers, params=params, timeout=60)
            r.raise_for_status()
            data = r.json() if r.content else {}

            page_tickets = data.get("tickets", [])
            if isinstance(page_tickets, dict):
                page_tickets = [page_tickets]
            if not isinstance(page_tickets, list):
                page_tickets = []

            tickets.extend(page_tickets)
            log(f"✅ דף {attempt}: נטענו {len(page_tickets)} כרטיסים (סה\"כ {len(tickets)})")

            paging = data.get("paging") or {}
            next_url = paging.get("next")
            if next_url:
                # אם יש next – ממשיכים עם ה-URL שנתנו (בלי פרמטרים נוספים!)
                url = next_url
                params = None
                time.sleep(1)
                continue
            break

        except requests.exceptions.HTTPError:
            log_exc("HTTPError בשליפת כרטיסים")
            break
        except requests.exceptions.RequestException:
            log_exc("RequestException בשליפת כרטיסים")
            time.sleep(wait_time)
            wait_time = min(wait_time * 2, max_wait_time)
            continue
        except Exception:
            log_exc("שגיאה כללית בשליפת כרטיסים")
            break

    return tickets


# ---------- DATA FRAME BUILD ----------
def safe_get_owner_name(owner):
    try:
        if isinstance(owner, dict):
            raw = owner.get("UserName") or owner.get("userName") or owner.get("User") or owner.get("email") or ""
            if raw:
                user_id = str(raw).split('@')[0].lower()
                return TIER2_MAP.get(user_id, user_id.capitalize())
        elif owner:
            raw_str = str(owner).lower()
            return TIER2_MAP.get(raw_str, raw_str.capitalize())
    except Exception:
        pass
    return ""

def pick_first_available(t: dict, keys: list):
    for k in keys:
        if k in t and t[k]:
            return t[k]
    return None

def normalize_tags(tag_value):
    try:
        if tag_value is None:
            return ""
        if isinstance(tag_value, list):
            items = []
            for v in tag_value:
                if isinstance(v, dict):
                    items.append(v.get("name") or v.get("Name") or v.get("value") or v.get("Value") or str(v))
                else:
                    items.append(str(v))
            return ", ".join([s for s in items if s])
        if isinstance(tag_value, dict):
            return tag_value.get("name") or tag_value.get("Name") or tag_value.get("value") or tag_value.get("Value") or json.dumps(tag_value, ensure_ascii=False)
        return str(tag_value)
    except Exception:
        return ""

def build_open_calls_df(tickets):
    rows = []
    for t in tickets if isinstance(tickets, list) else []:
        try:
            state = pick_first_available(t, ["state", "State", "ticketState", "TicketState"])
            state = str(state).strip().lower() if state else ""
            owner = t.get("owner") or t.get("Owner") or {}
            agent = safe_get_owner_name(owner)
            first_cust = pick_first_available(t, [
                "firstCustomerMessageDateTime", "firstCustomerMessageDate", "firstCustomer",
                "lastCustomerMessageDateTime"
            ])
            first_agent = pick_first_available(t, [
                "firstAgentMessageDateTime", "firstAgentMessageDate", "lastAgentMessageDateTime"
            ])
            created = pick_first_available(t, ["createdAt", "creationDate", "createdDate", "open"])
            closed = pick_first_available(t, ["close", "closedAt", "closedDate", "closeDate"])
            subject = pick_first_available(t, ["subject", "title"]) or ""
            ticket_id = pick_first_available(t, ["ticketId", "id", "_id", "TicketId", "TicketID", "Id"])
            tags_raw = pick_first_available(t, ["tags", "Tags"])
            tags_str = normalize_tags(tags_raw)

            # הוספת field1 כעמודה חדשה (נושא אמיתי מה-API)
            field1 = t.get("field1", "")

            rows.append({
                "id": ticket_id,
                "state": state,
                "agent": agent,
                "firstCustomerMessageDateTime": first_cust,
                "firstAgentMessageDateTime": first_agent,
                "createdAt": created,
                "closedAt": closed,
                "subject": subject,
                "field1": field1,  # ← חדש
                "tags": tags_str,
                "raw": t
            })
        except Exception:
            log_exc("שגיאה בעיבוד כרטיס יחיד (התעלמות והמשך)")
            continue

    df = pd.DataFrame(rows)
    for col in ["id", "state", "agent", "firstCustomerMessageDateTime",
                "firstAgentMessageDateTime", "createdAt", "closedAt", "subject", "field1", "tags"]:
        if col not in df.columns:
            df[col] = None
    try:
        df['firstCustomer_dt'] = df['firstCustomerMessageDateTime'].apply(to_utc_dt)
    except Exception:
        df['firstCustomer_dt'] = None
    try:
        df['firstAgent_dt'] = df['firstAgentMessageDateTime'].apply(to_utc_dt)
    except Exception:
        df['firstAgent_dt'] = None
    try:
        df['created_dt'] = df['createdAt'].apply(to_utc_dt)
    except Exception:
        df['created_dt'] = None
    try:
        df['closed_dt'] = df['closedAt'].apply(to_utc_dt)
    except Exception:
        df['closed_dt'] = None
    df['state'] = df['state'].fillna("").astype(str).str.lower()
    df['agent'] = df['agent'].fillna("").astype(str)
    df['tags'] = df['tags'].fillna("").astype(str)
    df['field1'] = df['field1'].fillna("").astype(str)  # ← נוספה גם כאן
    return df


# ---------- METRICS ----------
def compute_metrics(df):
    try:
        if df is None or df.empty:
            return {
                "total_open": 0,
                "per_agent_open": {},
                "sla_violations_df": pd.DataFrame(columns=df.columns if df is not None else []),
                "open_df": pd.DataFrame(columns=df.columns if df is not None else [])
            }
        open_mask = df['state'].fillna("").astype(str).str.lower().eq('open')
        open_df = df[open_mask].copy()
        total_open = len(open_df)
        try:
            per_agent_open = open_df.groupby(open_df['agent'].fillna('Unassigned')).size().to_dict()
        except Exception:
            per_agent_open = {}
        now = datetime.datetime.now(datetime.timezone.utc)
        def hours_since(dt):
            try:
                if dt is None or pd.isna(dt):
                    return None
                if isinstance(dt, pd.Timestamp):
                    dt = dt.to_pydatetime()
                if dt.tzinfo is None:
                    dt = dt.replace(tzinfo=datetime.timezone.utc)
                delta = now - dt
                return max(0.0, delta.total_seconds() / 3600.0)
            except Exception:
                return None
        for col in ['firstCustomer_dt', 'firstAgent_dt', 'created_dt']:
            if col not in open_df.columns:
                open_df[col] = None
        try:
            open_df['hours_since_first_customer'] = open_df['firstCustomer_dt'].apply(hours_since)
        except Exception:
            open_df['hours_since_first_customer'] = None
        try:
            open_df['hours_since_created'] = open_df['created_dt'].apply(hours_since)
        except Exception:
            open_df['hours_since_created'] = None
        open_df['first_agent_delay_hours'] = None
        try:
            for idx, row in open_df.iterrows():
                fc = row.get('firstCustomer_dt')
                fa = row.get('firstAgent_dt')
                if fc is not None and fa is not None:
                    try:
                        if isinstance(fc, pd.Timestamp):
                            fc = fc.to_pydatetime()
                        if isinstance(fa, pd.Timestamp):
                            fa = fa.to_pydatetime()
                        if fc.tzinfo is None:
                            fc = fc.replace(tzinfo=datetime.timezone.utc)
                        if fa.tzinfo is None:
                            fa = fa.replace(tzinfo=datetime.timezone.utc)
                        open_df.at[idx, 'first_agent_delay_hours'] = (fa - fc).total_seconds() / 3600.0
                    except Exception:
                        open_df.at[idx, 'first_agent_delay_hours'] = None
        except Exception:
            pass
        try:
            open_df['hours_since_first_customer'] = pd.to_numeric(open_df['hours_since_first_customer'], errors='coerce')
        except Exception:
            open_df['hours_since_first_customer'] = None
        try:
            open_df['first_agent_delay_hours'] = pd.to_numeric(open_df['first_agent_delay_hours'], errors='coerce')
        except Exception:
            open_df['first_agent_delay_hours'] = None
        try:
            cond1 = open_df['firstAgent_dt'].isna() & (open_df['hours_since_first_customer'] > SLA_LIMIT_HOURS)
        except Exception:
            cond1 = pd.Series([False] * len(open_df))
        try:
            cond2 = open_df['first_agent_delay_hours'].notna() & (open_df['first_agent_delay_hours'] > SLA_LIMIT_HOURS)
        except Exception:
            cond2 = pd.Series([False] * len(open_df))
        try:
            sla_violations_df = open_df[cond1 | cond2].copy()
        except Exception:
            sla_violations_df = open_df.iloc[0:0].copy()
        return {
            "total_open": total_open,
            "per_agent_open": per_agent_open,
            "sla_violations_df": sla_violations_df,
            "open_df": open_df
        }
    except Exception:
        log_exc("compute_metrics")
        return {
            "total_open": 0,
            "per_agent_open": {},
            "sla_violations_df": pd.DataFrame(),
            "open_df": pd.DataFrame()
        }

def compute_snoozed_metrics(df):
    try:
        if df is None or df.empty:
            return {
                "total_snoozed": 0,
                "per_agent_snoozed": {},
                "snoozed_df": pd.DataFrame(columns=df.columns if df is not None else [])
            }
        mask = df['state'].fillna("").astype(str).str.lower().eq('snoozed')
        snoozed_df = df[mask].copy()
        total_snoozed = len(snoozed_df)
        try:
            per_agent_snoozed = snoozed_df.groupby(snoozed_df['agent'].fillna('לא משויך')).size().to_dict()
        except Exception:
            per_agent_snoozed = {}
        return {
            "total_snoozed": total_snoozed,
            "per_agent_snoozed": per_agent_snoozed,
            "snoozed_df": snoozed_df
        }
    except Exception:
        log_exc("compute_snoozed_metrics")
        return {
            "total_snoozed": 0,
            "per_agent_snoozed": {},
            "snoozed_df": pd.DataFrame()
        }



# תשתית לבוט אינטראקטיבי
LAST_PROCESSED_IDS = {} # מילון לפי שם קבוצה

def get_whatsapp_driver():
    """מאתחל דרייבר של כרום בפורט 9223 (נפרד מה-Verint)"""
    try:
        from selenium.webdriver.chrome.options import Options
        import subprocess

        import socket
        def is_port_open(port):
            try:
                with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                    s.settimeout(1)
                    return s.connect_ex(('127.0.0.1', port)) == 0
            except: return False

        options = Options()
        
        if is_port_open(WA_DEBUG_PORT):
            log(f"[*] מזהה שפורט {WA_DEBUG_PORT} פתוח. מתחבר לוואטסאפ...")
            options.add_experimental_option("debuggerAddress", f"127.0.0.1:{WA_DEBUG_PORT}")
        else:
            log(f"[*] פותח חלון כרום נפרד עבור וואטסאפ (פורט {WA_DEBUG_PORT})...")
            chrome_exe = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            if not os.path.exists(chrome_exe):
                chrome_exe = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            
            # פרופיל נפרד למניעת נזק/התנגשות
            cmd = f'start "" "{chrome_exe}" --remote-debugging-port={WA_DEBUG_PORT} --user-data-dir="{CHROME_USER_DATA_DIR}" --profile-directory="Default" --start-maximized "https://web.whatsapp.com/"'
            subprocess.Popen(cmd, shell=True)
            time.sleep(10)
            options.add_experimental_option("debuggerAddress", f"127.0.0.1:{WA_DEBUG_PORT}")

        driver = webdriver.Chrome(options=options)
        
        if "web.whatsapp.com" not in driver.current_url:
            driver.get("https://web.whatsapp.com/")
            
        log("ממתין לטעינת WhatsApp Web...")
        wait = WebDriverWait(driver, 40)
        wait.until(EC.presence_of_element_located((By.ID, "pane-side")))
        log("✅ WhatsApp Web מוכן ובחלון נפרד.")
        return driver
            
    except Exception as e:
        log(f"❌ אתחול הדרייבר נכשל: {str(e)}")
        log("וודא שסגרת חלונות כרום אחרים שאינם במצב דיבאג (או תן לסקריפט לפתוח עבורך)")
        return None



def clean_text_for_comparison(text):
    """מנקה טקסט מתווים בלתי נראים, סימני פיסוק ורווחים כפולים"""
    if not text: return ""
    chars_to_remove = ["\u200f", "\u200e", "\u202a", "\u202b", "\u202c", "\u202d", "\u202e", "\u200b"]
    for c in chars_to_remove:
        text = text.replace(c, "")
    text = text.lower().replace("!", "").replace(".", "").replace("?", "").replace(",", "").strip()
    return text

def listen_for_commands(driver, group_name, token):
    """מאזין להודעות בקבוצה - מניעת כפילויות וזיהוי מדויק"""
    global LAST_PROCESSED_IDS
    try:
        # 1. וודא שאנחנו בצ'אט הנכון (קבוצת הבוט החדשה)
        try:
            # חיפוש חכם בכותרת
            header = driver.find_element(By.XPATH, '//*[@id="main"]//header')
            if group_name not in header.text:
                log(f"🔄 עובר לצ'אט של הבוט: {group_name}")
                group_xpath = f'//span[@title="{group_name}"]'
                driver.find_element(By.XPATH, group_xpath).click()
                time.sleep(1)
        except: 
            # אם לא גלוי או לא נמצא, ננסה לחפש ברשימת הצ'אטים
            try:
                group_xpath = f'//span[@title="{group_name}"]'
                el = driver.find_element(By.XPATH, group_xpath)
                el.click()
                time.sleep(1)
            except: pass

        # 2. חיפוש הודעות חדשות (ID לכל קבוצה בנפרד)
        last_id = LAST_PROCESSED_IDS.get(group_name)

        msg_xpath = '//div[@role="row"] | //div[contains(@class, "message-in")] | //div[contains(@class, "message-out")]'
        messages = driver.find_elements(By.XPATH, msg_xpath)
        
        if not messages: return

        latest_el = messages[-1]
        latest_id = latest_el.get_attribute("data-id") or str(hash(latest_el.text))
        
        if last_id is None:
            LAST_PROCESSED_IDS[group_name] = latest_id
            return

        if latest_id == last_id:
            return

        # 3. סריקת הודעות חדשות
        new_commands = []
        # בודקים רק נתח קטן אחרון
        for msg_el in reversed(messages[-8:]):
            m_id = msg_el.get_attribute("data-id") or str(hash(msg_el.text))
            if m_id == last_id: break
            
            cls = msg_el.get_attribute("class") or ""
            if "message-in" in cls:
                raw_text = msg_el.text.strip()
                if not raw_text: continue
                
                lines = [line.strip() for line in raw_text.split("\n")]
                commands_list = ["עזרה", "help", "סטטוס", "status", "סלא", "sla", "יומי", "daily", "בדיקה", "test"]
                
                for line in lines:
                    clean_line = clean_text_for_comparison(line)
                    # תומך גם בפקודות עם ! וגם בלי
                    is_cmd = any(clean_line == c or clean_line == f"!{c}" for c in commands_list)
                    if is_cmd or line.startswith("!"):
                        log(f"🎯 פקודה ב-{group_name}: '{clean_line}'")
                        new_commands.append(line)
                        break

        # מעדכנים את הסימניה
        LAST_PROCESSED_IDS[group_name] = latest_id

        # ביצוע
        for cmd in reversed(new_commands):
            handle_command(driver, cmd, token, group_name)

    except Exception as e:
        if "stale element" not in str(e).lower():
            log(f"⚠️ שגיאה במאזין: {str(e)}")

def handle_command(driver, text, token, group_name):
    """מפענח ומבצע פקודות - תגובה תמיד לאותה קבוצה ממנה הגיעה הפקודה"""
    clean_full = clean_text_for_comparison(text)
    if not clean_full: return
    
    clean_cmd = clean_full.split()[0]
    log(f"⚙️ מעבד פקודה '{clean_cmd}' עבור קבוצת '{group_name}'")
    
    response = None
    if clean_cmd in ["עזרה", "help"]:
        response = "🤖 *Tier 2 Bot - Commands:*\n\n▫️ *status* - Open/Snoozed report\n▫️ *sla* - Immediate SLA check\n▫️ *daily* - Daily closures summary\n▫️ *test* - Test connection"
    elif clean_cmd in ["סטטוס", "status"]:
        send_hourly_report(token, group_name, driver=driver)
        return
    elif clean_cmd in ["סלא", "sla"]:
        check_sla_and_alert(token, group_name, driver=driver)
        return
    elif clean_cmd in ["יומי", "daily"]:
        send_current_daily_summary(token, group_name, driver=driver)
        return
    elif clean_cmd in ["בדיקה", "test"]:
        response = "👋 Bot is online and listening!"
    

    if response:
        send_whatsapp_message_direct(driver, group_name, response)




def send_whatsapp_message_direct(driver, group_name, message):
    """שולח הודעה ישירות דרך הדרייבר הפתוח - גרסה חסינה במיוחד"""
    try:
        wait = WebDriverWait(driver, 15)
        
        # 0. ניקוי תיבת חיפוש אם נשארו בה שאריות (בלי לזרוק שגיאה אם לא נמצא)
        try:
            search_box = driver.find_element(By.XPATH, '//div[@role="textbox" and @data-tab="3"] | //div[@contenteditable="true"][@data-tab="3"]')
            if search_box.text:
                search_box.click()
                search_box.send_keys(Keys.CONTROL, "a")
                search_box.send_keys(Keys.BACKSPACE)
        except:
            pass

        # 1. בדיקה מדויקת אם הקבוצה הנכונה פתוחה
        is_open = False
        try:
            header = driver.find_element(By.XPATH, '//*[@id="main"]//header')
            if group_name in header.text:
                is_open = True
        except:
            pass

        if not is_open:
            log(f"🔍 מחפש את הקבוצה '{group_name}'...")
            
            # 2. ניסיון למצוא את הקבוצה ברשימת הצ'אטים (חיפוש גמיש)
            group_xpath = '//span[contains(@title, "דוחות") and contains(@title, "Tier2")]'
            try:
                # גלילה קלה למעלה כדי לוודא שהרשימה מעודכנת
                side_pane = driver.find_element(By.ID, "pane-side")
                driver.execute_script("arguments[0].scrollTop = 0;", side_pane)
                
                group_el = wait.until(EC.element_to_be_clickable((By.XPATH, group_xpath)))
                group_el.click()
            except:
                log(f"⚠️ קבוצה לא נמצאה ברשימה, מנסה לבצע חיפוש אקטיבי...")
                # 3. חיפוש דרך תיבת החיפוש
                search_xpaths = [
                    '//div[@role="textbox" and @data-tab="3"]',
                    '//div[@contenteditable="true"][@data-tab="3"]//p',
                    '//div[@contenteditable="true"][@data-tab="3"]',
                    '//div[@title="חיפוש או התחלת צ\'אט חדש"]'
                ]
                search_box = None
                for sx in search_xpaths:
                    try:
                        search_box = driver.find_element(By.XPATH, sx)
                        break
                    except: continue
                
                if not search_box:
                    raise Exception("לא הצלחתי לאתר את תיבת החיפוש בוואטסאפ")
                
                search_box.click()
                search_box.send_keys(group_name)
                time.sleep(2)
                
                group_el = wait.until(EC.element_to_be_clickable((By.XPATH, group_xpath)))
                group_el.click()

        # 4. מציאת תיבת ההקלדה והדבקה ( Paste)
        time.sleep(1)
        input_box_xpaths = [
            '//*[@id="main"]//footer//div[@contenteditable="true"][@data-tab="10"]',
            '//*[@id="main"]//footer//div[@contenteditable="true"]',
            '//div[@id="main"]//div[@title="הקלדת הודעה"]',
            '//footer//div[@role="textbox"]'
        ]
        
        input_box = None
        for ix in input_box_xpaths:
            try:
                input_box = driver.find_element(By.XPATH, ix)
                if input_box.is_displayed(): break
            except: continue
            
        if not input_box:
            raise Exception("לא נמצאה תיבת הקלדה (Input Box)")

        import pyperclip
        pyperclip.copy(message)
        input_box.click()
        time.sleep(0.3)
        input_box.send_keys(Keys.CONTROL, "v")
        time.sleep(0.5)
        input_box.send_keys(Keys.ENTER)
        log("✅ הודעה נשלחה בהצלחה.")
        return True
        
    except Exception as e:
        log(f"❌ שגיאה בשליחה ישירה: {str(e)}")
        return False

def send_whatsapp_group_instant(group_id_or_name, message):
    """שליחה 'מהירה' על ידי פתיחה וסגירה של דפדפן (לגיבוי בלבד)"""
    driver = None
    try:
        options = Options()
        options.add_argument(f"--user-data-dir={CHROME_USER_DATA_DIR}")
        options.add_argument("--profile-directory=Default")
        driver = webdriver.Chrome(options=options)
        driver.get("https://web.whatsapp.com/")
        wait = WebDriverWait(driver, 60)
        wait.until(EC.presence_of_element_located((By.XPATH, f'//span[@title="{group_id_or_name}"]'))).click()
        input_box = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')))
        import pyperclip
        pyperclip.copy(message)
        input_box.send_keys(Keys.CONTROL, "v")
        time.sleep(0.5)
        input_box.send_keys(Keys.ENTER)
        time.sleep(1)
        driver.quit()
        return True, None
    except Exception as e:
        if driver: driver.quit()
        return False, str(e)


def send_with_retries(group_name, message, driver=None):
    """שולח הודעה עם העדפה לדרייבר פתוח"""
    if driver:
        if send_whatsapp_message_direct(driver, group_name, message):
            return True
    
    for i, delay in enumerate(RETRY_DELAYS, start=1):
        success, _ = send_whatsapp_group_instant(group_name, message)
        if success: return True
        time.sleep(delay)
    return False

# ---------- UPDATED SLA & REPORTS ----------

def check_sla_and_alert(token, group_name, driver=None):
    """בודק SLA ושולח התראה מרוכזת"""
    try:
        tickets = get_tickets(token, states=["open"])
        df = build_open_calls_df(tickets)
        metrics = compute_metrics(df)
        sla_df = metrics["sla_violations_df"]
        global alerted_ticket_ids
        new_violations = []
        for _, row in sla_df.iterrows():
            tid = row.get("id")
            if tid not in alerted_ticket_ids:
                new_violations.append({
                    "id": tid,
                    "agent": row.get("agent", "לא משויך"),
                    "hours": row.get("hours_since_first_customer", 0.0),
                    "subject": row.get("field1", "ללא")
                })
        if not new_violations: return
        msg = ["🚨 *ריכוז חריגות SLA — Tier 2*", "════════════════════"]
        for v in new_violations:
            msg.extend([f"🎫 *קריאה:* {v['id']}", f"👤 *נציג:* {v['agent']}", f"⏱ *זמן:* {v['hours']:.1f} שעות", "-------------------"])
            alerted_ticket_ids.add(v['id'])
        msg.append("\n🤖 _נשלח אוטומטית_")
        send_with_retries(group_name, "\n".join(msg), driver=driver)
    except: log_exc("check_sla")


def send_hourly_report(token, group_name, driver=None):
    try:
        tickets = get_tickets(token)
        df = build_open_calls_df(tickets)
        now_local = datetime.datetime.now(pytz.timezone("Asia/Jerusalem"))
        total_open = len(df[df['state'] == "open"])
        total_snoozed = len(df[df['state'] == "snoozed"])
        
        msg = [
            "📊 *סטטוס שעתי — Tier 2*",
            "════════════════════",
            f"⏰ שעה: *{now_local.strftime('%H:%M')}*",
            f"\n📌 *קריאות פתוחות:* {total_open}",
        ]
        open_agents = df[df['state'] == "open"]['agent'].value_counts()
        if not open_agents.empty:
            for ag, count in open_agents.items():
                msg.append(f"   ▫️ {ag}: {count}")
        else:
            msg.append("   ▫️ אין קריאות פתוחות")
            
        msg.append(f"\n💤 *בממתינה (Snoozed):* {total_snoozed}")
        snoozed_agents = df[df['state'] == "snoozed"]['agent'].value_counts()
        if not snoozed_agents.empty:
            for ag, count in snoozed_agents.items():
                msg.append(f"   ▫️ {ag}: {count}")

        msg.append("\n════════════════════")
        msg.append("🤖 _עדכון שעתי אוטומטי_")
        
        send_with_retries(group_name, "\n".join(msg), driver=driver)
        log(f"דו\"ח שעתי נשלח ({now_local.strftime('%H:%M')})")
    except Exception:
        log_exc("send_hourly_report")

def send_current_daily_summary(token, group_name, driver=None):
    """סיכום שוטף עבור פקודת 'יומי' - כולל Glassix ו-Verint"""
    try:
        tz = pytz.timezone("Asia/Jerusalem")
        now_local = datetime.datetime.now(tz)
        start_day = now_local.replace(hour=0, minute=0, second=0, microsecond=0)

        # 1. נתוני Glassix
        tickets = get_tickets(token, since=start_day, until=now_local, states=("Closed",))
        df_closed = build_open_calls_df(tickets)
        total_closed = len(df_closed)
        
        msg = [
            f"📈 *סטטוס שוטף — {now_local.strftime('%d/%m/%Y %H:%M')}*",
            "════════════════════",
            "✅ *סגירות (Glassix)*",
            f"סה\"כ קריאות שנסגרו היום: *{total_closed}*",
            ""
        ]
        
        if not df_closed.empty:
            for ag, count in df_closed['agent'].value_counts().items():
                msg.append(f"   👤 {ag}: {count}")
        else:
            msg.append("   ▫️ טרם נסגרו קריאות היום")

        # 2. נתוני Verint
        vstats = analyze_verint_daily(now_local)
        if vstats:
            msg.extend([
                "",
                "📞 *שיחות (Verint)*",
                f"סה\"כ שיחות: *{vstats['total_calls']}*",
                f"פילוח: 🏰 Vico:{vstats['vico']} | 📦 Shuf:{vstats['shuf']} | 💎 Tier1:{vstats['tier1']}",
                ""
            ])
            for ag, count in sorted(vstats['agent_stats'].items(), key=lambda x: x[1], reverse=True):
                msg.append(f"   ▫️ {ag}: {count}")

        msg.append("\n🤖 _עדכון שוטף לבקשת משתמש_")
        send_with_retries(group_name, "\n".join(msg), driver=driver)
        log("סיכום יומי שוטף נשלח.")
    except Exception:
        log_exc("send_current_daily_summary")

def send_daily_report(token, group_name, driver=None):
    """דוח סוף יום משולב: Glassix + Verint"""
    try:
        tz = pytz.timezone("Asia/Jerusalem")
        now_local = datetime.datetime.now(tz)
        start_day = now_local.replace(hour=0, minute=0, second=0, microsecond=0)

        # 1. נתוני Glassix
        tickets = get_tickets(token, since=start_day, until=now_local, states=("Closed",))
        df_closed = build_open_calls_df(tickets)
        total_closed = len(df_closed)
        
        msg = [
            f"🏁 *סיכום סוף יום — {now_local.strftime('%d/%m/%Y')}*",
            "════════════════════",
            "✅ *סגירות (Glassix)*",
            f"סה\"כ קריאות שנסגרו היום: *{total_closed}*",
            ""
        ]
        
        if not df_closed.empty:
            for ag, count in df_closed['agent'].value_counts().items():
                msg.append(f"   👤 {ag}: {count}")
        else:
            msg.append("   ▫️ לא נסגרו קריאות היום")

        # 2. נתוני Verint
        vstats = analyze_verint_daily(now_local)
        if vstats:
            msg.extend([
                "",
                "📞 *שיחות (Verint)*",
                f"סה\"כ שיחות: *{vstats['total_calls']}*",
                f"פילוח שירותים:",
                f"🏰 וייקו: {vstats['vico']}",
                f"📦 שופרסל: {vstats['shuf']}",
                f"💎 טייר 1: {vstats['tier1']}",
                f"🚀 ורטיקלים: {vstats['vert']}",
                "",
                "📊 *ביצועי נציגים (שיחות):*"
            ])
            for ag, count in sorted(vstats['agent_stats'].items(), key=lambda x: x[1], reverse=True):
                msg.append(f"   ▫️ {ag}: {count}")

        msg.extend([
            "\n════════════════════",
            "🌟 עבודה טובה צוות!",
            "🤖 _סיכום יום אוטומטי_"
        ])

        send_with_retries(group_name, "\n".join(msg), driver=driver)
        log("דו\"ח סוף יום משולב נשלח בהצלחה.")
    except Exception:
        log_exc("send_daily_report")

def send_weekly_report(token, group_name, driver=None):
    """דוח שבועי מעוצב ומקצועי (א-ו)"""
    try:
        tz = pytz.timezone("Asia/Jerusalem")
        now = datetime.datetime.now(tz)
        
        # תחילת שבוע: ראשון 00:00
        days_since_sunday = (now.weekday() + 1) % 7
        start_of_week = (now - datetime.timedelta(days=days_since_sunday)).replace(hour=0, minute=0, second=0, microsecond=0)
        
        tickets = get_tickets(token, since=start_of_week, until=now, states=("Closed",))
        df = build_open_calls_df(tickets)
        if 'closed_dt' in df.columns:
            df = df[df['closed_dt'].notna()]
            
        total_closed = len(df)
        if total_closed == 0: return
        
        sla_met = df[df['sla_met'] == True].shape[0] if 'sla_met' in df.columns else 0
        sla_pct = round((sla_met / total_closed) * 100, 1)
        fcr_met = df[df['first_contact_resolved'] == True].shape[0] if 'first_contact_resolved' in df.columns else 0
        fcr_pct = round((fcr_met / total_closed) * 100, 1)
        
        msg = [
            f"📅 *סיכום שבועי — {start_of_week.strftime('%d/%m')} עד {now.strftime('%d/%m')}*",
            "══════════════════════",
            "📊 *KPIs שבועיים*",
            f"• סה\"כ סגירות: *{total_closed}*",
            f"• עמידה ב־SLA: *{sla_pct}%*",
            f"• פתרון שלב א' (FCR): *{fcr_pct}%*",
            "",
            "👥 *ביצועי נציגים*",
        ]
        
        for ag, count in df['agent'].value_counts().items():
            msg.append(f"   👤 {ag}: {count} סגירות")
            
        msg.extend([
            "",
            "══════════════════════",
            "🚀 שבוע מצוין לכולם!",
            "🤖 _עדכון שבועי אוטומטי_"
        ])
        
        send_with_retries(group_name, "\n".join(msg), driver=driver)
    except Exception:
        log_exc("send_weekly_report")

def send_monthly_report(token, group_name, start_date, end_date, driver=None):
    """דוח חודשי מעוצב"""
    try:
        tickets = get_tickets(token, since=start_date, until=end_date, states=("Closed",))
        df = build_open_calls_df(tickets)
        total = len(df)
        if total == 0: return

        msg = [
            f"📑 *סיכום חודשי — {start_date.strftime('%m/%Y')}*",
            "════════════════════",
            f"סה\"כ סגירות בחודש: *{total}*",
            ""
        ]
        
        top_agents = df['agent'].value_counts().head(3)
        msg.append("🏆 *מובילים החודש:*")
        for ag, count in top_agents.items():
            msg.append(f"   🥇 {ag}: {count} סגירות")
            
        msg.extend([
            "\n════════════════════",
            "🤖 _עדכון חודשי אוטומטי_"
        ])
        send_with_retries(group_name, "\n".join(msg), driver=driver)
    except Exception:
        log_exc("send_monthly_report")

def is_weekend_block_time():
    tz = pytz.timezone("Asia/Jerusalem")
    now = datetime.datetime.now(tz)
    wd = now.weekday()
    if wd == 4 and now.hour >= 15: return True # שישי אחרי 15:00
    if wd == 5: return True # שבת
    if wd == 6 and now.hour < 8: return True # ראשון לפני 08:00
    return False

# ---------- VERINT HELPERS ----------
def process_duration(dur_str):
    try:
        if pd.isna(dur_str): return 0
        parts = str(dur_str).split(':')
        if len(parts) == 3: return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        if len(parts) == 2: return int(parts[0])*60 + int(parts[1])
        return 0
    except: return 0

def analyze_verint_daily(date_obj):
    csv_filename = f"Verint_Today_{date_obj.strftime('%Y%m%d')}.csv"
    csv_path = os.path.join(VERINT_REPORTS_DIR, csv_filename)
    if not os.path.exists(csv_path):
        log(f"⚠️ קובץ Verint לא נמצא: {csv_filename}")
        return None
    try:
        df = pd.read_csv(csv_path)
        # ניקוי מספרי טלפון כפי שנעשה ב-Combined_Reporter
        df['ani_clean'] = df['Dialed From (ANI)'].astype(str).str.replace('+', '', regex=False)
        df['dnis_clean'] = df['Dialed To (DNIS)'].astype(str).str.replace('+', '', regex=False)
        
        df['Start Time'] = pd.to_datetime(df['Start Time'], dayfirst=True, errors='coerce')
        df = df[df['Start Time'].dt.date == date_obj.date()].dropna(subset=['Start Time'])
        if df.empty: return None
        
        # פילוח יעד (DNIS/ANI)
        vico = len(df[df['ani_clean'] == '97239029740'])
        tier1 = len(df[df['dnis_clean'] == '97239029740'])
        vert = len(df[df['dnis_clean'] == '972732069574'])
        shuf = len(df[df['dnis_clean'] == '972732069576'])

        def clean_name(n):
            raw = str(n).strip()
            # שימוש במיפוי הקיים
            if raw in TIER2_MAP: return TIER2_MAP[raw]
            if ',' in raw:
                p = raw.split(',')
                # הופכת 'Isakov, Moshe' ל-'Moshe Isakov' ואז בודקת שוב במיפוי
                swapped = f"{p[1].strip()} {p[0].strip()}"
                return TIER2_MAP.get(swapped, swapped)
            return TIER2_MAP.get(raw, raw)

        agent_stats = {clean_name(k): v for k, v in df.groupby('Employee').size().to_dict().items()}
        return {
            "total_calls": len(df), 
            "agent_stats": agent_stats,
            "vico": vico,
            "tier1": tier1,
            "vert": vert,
            "shuf": shuf
        }
    except Exception as e:
        log(f"❌ שגיאה בניתוח CSV: {str(e)}")
        return None


def monitor_loop(group_name):
    global global_token
    tz = pytz.timezone("Asia/Jerusalem")
    last_hourly_report_hour = None
    last_daily_report_date = None
    last_weekly_report_date = None
    last_monthly_report_month = None
    last_scheduled_minute = None

    # אתחול ראשוני
    driver = get_whatsapp_driver()
    log("מתחיל monitor_loop עם מנגנון התאוששות מקריסות.")

    while True:
        try:
            # ===== בדיקת תקינות הדרייבר =====
            driver_alive = False
            if driver:
                try:
                    # בדיקה פשוטה אם הדפדפן מגיב
                    _ = driver.current_url
                    driver_alive = True
                except:
                    log("⚠️ זוהתה קריסה/סגירה של הדפדפן. מנסה לאתחל מחדש...")
                    try: driver.quit()
                    except: pass
                    driver = None

            if not driver_alive:
                driver = get_whatsapp_driver()
                if not driver:
                    log("❌ נכשל באתחול דפדפן, מנסה שוב בעוד 30 שניות...")
                    time.sleep(30)
                    continue

            token = global_token or get_token()
            now = datetime.datetime.now(tz)
            hour = now.hour
            minute = now.minute
            weekday = now.weekday()
            today = now.date()
            current_time_key = (today, hour, minute)

            # ===== האזנה לפקודות (קבוצת הבוט הייעודית) =====
            listen_for_commands(driver, BOT_GROUP, token)

            # מניעת הרצת דוחות מרובת פעמים באותה דקה
            is_new_minute = (current_time_key != last_scheduled_minute)

            # ===== SLA (בדיקה כל 5 דקות - קבוצת הדוחות) =====
            if minute % 5 == 0 and is_new_minute: 
                if 8 <= hour < 19 and not is_weekend_block_time():
                    check_sla_and_alert(token, REPORTS_GROUP, driver=driver)
                    last_scheduled_minute = current_time_key

            # ===== דו"ח שעתי (קבוצת הדוחות) =====
            if not is_weekend_block_time() and is_new_minute:
                if last_hourly_report_hour is None and 8 <= hour <= 19:
                    log(f"שולח דו\"ח שעתי ראשון ({hour:02d}:{minute:02d})...")
                    send_hourly_report(token, REPORTS_GROUP, driver=driver)
                    last_hourly_report_hour = hour
                    last_scheduled_minute = current_time_key
                elif hour == 19 and minute == 0 and last_hourly_report_hour != 19:
                    log("שעה 19:00 — שולח דו\"ח שעתי האחרון להיום.")
                    send_hourly_report(token, REPORTS_GROUP, driver=driver)
                    last_hourly_report_hour = 19
                    last_scheduled_minute = current_time_key
                elif 8 <= hour < 19 and minute == 0 and last_hourly_report_hour != hour:
                    log(f"שעה עגולה {hour:02d}:00 — שולח דו\"ח שעתי.")
                    send_hourly_report(token, REPORTS_GROUP, driver=driver)
                    last_hourly_report_hour = hour
                    last_scheduled_minute = current_time_key

            # ===== דו"ח יומי (סוף יום - קבוצת הדוחות) =====
            if hour == 19 and minute == 0 and last_daily_report_date != today and is_new_minute:
                log("שעה 19:00 — שולח דו\"ח סוף יום משולב.")
                send_daily_report(token, REPORTS_GROUP, driver=driver)
                last_daily_report_date = today
                last_scheduled_minute = current_time_key

            # ===== דו"ח שבועי (קבוצת הדוחות) =====
            if weekday == 4 and hour == 14 and minute == 0 and last_weekly_report_date != today and is_new_minute:
                log("שולח דו\"ח שבועי (שישי 14:00).")
                send_weekly_report(token, REPORTS_GROUP, driver=driver)
                last_weekly_report_date = today
                last_scheduled_minute = current_time_key

            # ===== דו"ח חודשי (קבוצת הדוחות) =====
            if today.day == 1 and hour == 8 and minute == 15 and last_monthly_report_month != now.month and is_new_minute:
                log("שולח דו\"ח חודשי (1 לחודש ב-08:15).")
                prev_month_last_day = today.replace(day=1) - datetime.timedelta(days=1)
                prev_month_first_day = prev_month_last_day.replace(day=1)
                send_monthly_report(token, REPORTS_GROUP, prev_month_first_day, prev_month_last_day, driver=driver)
                last_monthly_report_month = now.month
                last_scheduled_minute = current_time_key

        except Exception as e:
            msg = str(e)
            if "target window already closed" in msg or "invalid session id" in msg:
                log("⚠️ חלון הדפדפן נסגר, יאותחל באיטרציה הבאה.")
                driver = None
            else:
                log_exc("שגיאה בלולאת המעקב הראשית")

        time.sleep(10) # בדיקה כל 10 שניות לטובת אינטראקטיביות


# ---------- ENTRY ----------
if __name__ == "__main__":
    try:
        # מפעיל ריענון טוקן ברקע
        threading.Thread(target=refresh_token_loop, daemon=True).start()

        # מפעיל לולאת המוניטור
        monitor_loop(REPORTS_GROUP)

    except KeyboardInterrupt:
        log("הפסקת הריצה על ידי המשתמש. ביי.")
    except Exception:
        log_exc("שגיאה קריטית ראשית")
