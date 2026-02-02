# -*- coding: utf-8 -*-
"""
דוח משולב - Tier 2 Tickets (Glassix) + Verint Calls (Outlook)
גרסה: Outlook Automation Edition - STABLE
"""
import requests
import pandas as pd
from datetime import datetime, timedelta
import os
import time
import matplotlib.pyplot as plt
import win32com.client
from collections import defaultdict
import arabic_reshaper
from bidi.algorithm import get_display
import base64
from io import BytesIO
import shutil
import tempfile

# ===== הגדרות Glassix (Tickets) =====
API_KEY = "a0bb0de4-2193-41c6-bff6-2f87344953ea"
API_SECRET = "ZWHRKYQNdHsX3HuoK27Xk6omQchnieko28iadd3qxTyxAVKMu1K54jLVsFNoa3nsJC1Ea4ajfg6zsAcIbQOit36B2urQCpGd4K6nkPeJmtixYSoP6ZMwTmCgWgQiVnLt"
EMAIL_FROM = "MosheI1@VERIFONE.com"
EMAIL_TO = "MosheI1@VERIFONE.com"
EMAIL_CC = "MosheI1@VERIFONE.com"

TIER2_MAP = {
    "niv.arieli": "ניב אריאלי",
    "din.weissman": "דין וייסמן",
    "lior.burstein": "ליאור בורשטיין",
    "avivs": "אביב סולר",
    "ebrahimf": "אברהים פריג",
    "orenw1": "אורן וייס",
    "ahmado": "אחמד עודה",
    "almancha": "אלמנך עלמיה",
    "zahiyas1": "זהייה אבו שמאלה",
    "tals": "טל שוקר",
    "yuvala1": "יובל אגרון",
    "yuliano": "יוליאן אולרסקו",
    "yoadc": "יועד כחלון",
    "nuphars": "נוּפר שלום",
    "idoh": "עידו הרמל",
    "aviele": "אביאל אלשוילי",
    "avivk": "אביב כץ",
    "bari": "בר ישראלי",
    "veral2": "ורה ליברמן",
    "danv1": "דן וייסמן",
    "liorb5": "ליאור בורשטיין",
    "lior.burstein": "ליאור בורשטיין",
    "niva2": "ניב אריאלי",
    "niv.arieli": "ניב אריאלי",
    "nadavl1": "נדב",
    "paulp": "פאול",
    "din.weissman": "דין וייסמן",
    "moshei1": "משה איסקוב",
    "nadav.lieber": "נדב",
    "erezm1": "ארז",
    "almanch.alme": "אלמנך עלמיה"
}

# Mapping for English names in Verint Excel to Hebrew names
VERINT_NAME_MAP = {
    "Dan Vaysman": "דן וייסמן",
    "Niv Arieli": "ניב אריאלי",
    "Moshe Isakov": "משה איסקוב",
    "Lior Braunstein": "ליאור בורשטיין",
    "Lior Burstein": "ליאור בורשטיין",
    "Erez Mizrahi": "ארז",
    "Nadav Lieber": "נדב",
    "Almanch Alme": "אלמנך עלמיה",
    "Aviv Solar": "אביב סולר",
    "Niv Areli": "ניב אריאלי",
    "Dan Weissman": "דן וייסמן",
    "Din Weissman": "דן וייסמן",
    "Tal Shoker": "טל שוקר",
    "Yuval Agron": "יובל אגרון",
    "Yulian Olersku": "יוליאן אולרסקו",
    "Ido Harmel": "עידו הרמל"
}

DELAY_SECONDS = 120

# ===== הגדרות כלליות =====
DOWNLOAD_DIR = os.path.join(os.getcwd(), "Verint_Reports")
TIER2_DIR = os.path.join(os.getcwd(), "TIER2")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)
if not os.path.exists(TIER2_DIR):
    os.makedirs(TIER2_DIR)

plt.rcParams['axes.unicode_minus'] = False

# ===== עזר לעברית בגרפים =====
def reshape_hebrew(text):
    try:
        if not isinstance(text, str): return str(text)
        reshaped_text = arabic_reshaper.reshape(text)
        return get_display(reshaped_text)
    except Exception:
        return text

# ===== Glassix API =====
def get_access_token():
    url = "https://verifone.glassix.com/api/v1.2/token/get"
    payload = {"apiKey": API_KEY, "apiSecret": API_SECRET, "userName": EMAIL_FROM}
    response = requests.post(url, json=payload, timeout=90)
    response.raise_for_status()
    return response.json().get("access_token")

def safe_get(url, headers):
    while True:
        try:
            response = requests.get(url, headers=headers, timeout=90)
            if response.status_code == 429:
                print("Too many requests, מחכה קצת...")
                time.sleep(DELAY_SECONDS)
                continue
            response.raise_for_status()
            return response.json()
        except Exception as e:
            print(f"Error in API request: {e}")
            return {}

def get_glassix_tickets(token, since=None, until=None):
    headers = {"Authorization": f"Bearer {token}"}
    tickets_all = []
    url = f"https://verifone.glassix.com/api/v1.2/tickets/list?since={since}&until={until}"
    
    while url:
        data = safe_get(url, headers)
        tickets = data.get("tickets", [])
        tickets_all.extend(tickets)
        paging = data.get("paging")
        url = paging.get("next") if paging and "next" in paging else None
        
    return tickets_all

def parse_tickets(tickets, work_days=1, hours_per_day=8):
    agents = {}
    tags = {}
    first_response_times = defaultdict(list)
    first_response_times_tags = defaultdict(list)
    total_hours = (work_days * hours_per_day) if work_days > 0 else 8
    
    # Advanced Tracking
    hourly_volume = {h: 0 for h in range(24)}
    hourly_closed = {h: 0 for h in range(24)}
    hourly_first_response = {h: [] for h in range(24)}
    queue_dist = {"0-1m": 0, "1-3m": 0, "3-5m": 0, "5m+": 0}
    reopen_count = 0
    old_tickets_closed = 0
    
    # NEW: AHT Tracking
    total_handle_time_sec = 0
    closed_with_aht_count = 0

    for t in tickets:
        state = (t.get("state") or "").lower()
        
        # 0. Reopen detection
        if t.get("isReopened") or t.get("reopened"): reopen_count += 1

        # 1. Hourly Metrics
        created_str = t.get("firstCustomerMessageDateTime") or t.get("open")
        closed_str = t.get("close")
        
        if created_str:
            try:
                dt_created = datetime.fromisoformat(str(created_str).replace("Z", "+00:00"))
                hourly_volume[dt_created.hour] += 1
                
                if state in ["closed", "resolved", "סגור"] and closed_str:
                    dt_closed = datetime.fromisoformat(str(closed_str).replace("Z", "+00:00"))
                    # Same day filter
                    if dt_closed.date() == dt_created.date():
                        hourly_closed[dt_closed.hour] += 1
                    else:
                        old_tickets_closed += 1
            except: pass

        # 2. Handle Time (AHT) - NEW
        if state in ["closed", "resolved", "סגור"]:
            try:
                dur_net = t.get("durationNet")
                if dur_net:
                    parts = [int(p) for p in str(dur_net).split(':')]
                    if len(parts) == 3: dur = parts[0]*3600 + parts[1]*60 + parts[2]
                    elif len(parts) == 2: dur = parts[0]*60 + parts[1]
                    else: dur = 0
                    if dur > 0:
                        total_handle_time_sec += dur
                        closed_with_aht_count += 1
            except: pass

        # 3. Queue Distribution
        q_wait_net = t.get("queueTimeNet") or t.get("queueTimeGross")
        if q_wait_net:
            try:
                parts = str(q_wait_net).split(':')
                sec = int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
                mins = sec / 60
                if mins <= 1: queue_dist["0-1m"] += 1
                elif mins <= 3: queue_dist["1-3m"] += 1
                elif mins <= 5: queue_dist["3-5m"] += 1
                else: queue_dist["5m+"] += 1
            except: pass

        # 4. Agent & Tag tracking
        owner = t.get("owner", {})
        if isinstance(owner, dict):
            raw_name = owner.get("UserName") or owner.get("userName") or ""
            username_local = (raw_name.split('@')[0] or "").lower()
            agent_key = username_local
            agent_display = TIER2_MAP.get(username_local, username_local.capitalize())
        else:
            agent_key = str(owner).lower()
            agent_display = str(owner)
        
        if not agent_key or agent_key == 'none': continue
        if len(agent_key) > 30 or "bot" in agent_key: continue

        tags_list = t.get("tags", [])
        if isinstance(tags_list, str): tags_list = [x.strip() for x in tags_list.split(",") if x.strip()]
        elif not isinstance(tags_list, list): tags_list = []

        if agent_key not in agents:
            agents[agent_key] = {
                "AgentKey": agent_key, "Agent": agent_display,
                "Open": 0, "Closed": 0, "Snoozed": 0, "Other": 0, "Total": 0,
                "SLA": 0, "FCR": 0, "AvgFirstResponseHours": 0, "AvgCallsPerHour": 0,
                "TotalHandleTimeSec": 0, "ClosedWithAHT": 0
            }
        
        agents[agent_key]["Total"] += 1
        if state == "open": agents[agent_key]["Open"] += 1
        elif state == "closed": agents[agent_key]["Closed"] += 1
        elif state == "snoozed": agents[agent_key]["Snoozed"] += 1
        else: agents[agent_key]["Other"] += 1

        # Track Agent AHT
        if state in ["closed", "resolved", "סגור"]:
            try:
                dur_net = t.get("durationNet")
                if dur_net:
                    parts = [int(p) for p in str(dur_net).split(':')]
                    if len(parts) == 3: dur = parts[0]*3600 + parts[1]*60 + parts[2]
                    elif len(parts) == 2: dur = parts[0]*60 + parts[1]
                    else: dur = 0
                    if dur > 0:
                        agents[agent_key]["TotalHandleTimeSec"] += dur
                        agents[agent_key]["ClosedWithAHT"] += 1
            except: pass

        for tag in tags_list:
            if tag not in tags:
                tags[tag] = {"Tag": tag, "Open": 0, "Closed": 0, "Snoozed": 0, "Other": 0, "AvgFirstResponseHours": 0}
            if state == "open": tags[tag]["Open"] += 1
            elif state == "closed": tags[tag]["Closed"] += 1
            elif state == "snoozed": tags[tag]["Snoozed"] += 1
            else: tags[tag]["Other"] += 1

        first_customer = t.get("firstCustomerMessageDateTime")
        first_agent = t.get("firstAgentMessageDateTime")
        if first_customer and first_agent:
            try:
                dt_cust = datetime.fromisoformat(str(first_customer).replace("Z", "+00:00"))
                dt_agent = datetime.fromisoformat(str(first_agent).replace("Z", "+00:00"))
                effective_start = dt_cust
                is_weekend = dt_cust.weekday() in [4, 5]
                is_outside_hours = dt_cust.hour < 8 or dt_cust.hour >= 17
                if is_weekend or is_outside_hours or dt_cust.date() < dt_agent.date():
                    effective_start = dt_agent.replace(hour=8, minute=0, second=0, microsecond=0)
                diff_hours = (dt_agent - effective_start).total_seconds() / 3600
                if diff_hours < 0: diff_hours = 0
                if 0 <= diff_hours <= 1000:
                    first_response_times[agent_key].append(diff_hours)
                    hourly_first_response[dt_cust.hour].append(diff_hours)
                    for tag in tags_list:
                        first_response_times_tags[tag].append(diff_hours)
            except: pass
        else:
            first_response_times[agent_key].append(0.0028)
            for tag in tags_list:
                first_response_times_tags[tag].append(0.0028)

    total_tickets = len(tickets)
    for agent_key, stat in agents.items():
        total = stat["Total"]
        closed = stat["Closed"]
        snoozed = stat["Snoozed"]
        responses = first_response_times[agent_key]
        fast_responses = [r for r in responses if r <= 1.5]
        stat["SLA"] = round((len(fast_responses) / len(responses)) * 100, 1) if responses else 0
        stat["FCR"] = round(((closed + snoozed) / total) * 100, 1) if total else 0
        stat["AvgFirstResponseHours"] = round((sum(responses) / len(responses)) if responses else 0, 2)
        avg_calls = (closed / total_hours) if closed > 0 else 0
        stat["AvgCallsPerHour"] = round(avg_calls, 2)
        # Agent AHT Min
        stat["AvgAHTMin"] = round((stat["TotalHandleTimeSec"] / stat["ClosedWithAHT"] / 60), 2) if stat["ClosedWithAHT"] > 0 else 0

    for tag_name, stat in tags.items():
        stat["Total"] = stat["Open"] + stat["Closed"] + stat["Snoozed"] + stat["Other"]
        stat["Share"] = round((stat["Total"] / total_tickets) * 100, 1) if total_tickets > 0 else 0
        times = first_response_times_tags.get(tag_name, [])
        stat["AvgFirstResponseHours"] = round(sum(times) / len(times), 2) if times else 0
    
    # Hourly response average
    hourly_resp_avg = {h: (round(sum(v)/len(v), 2) if v else 0) for h, v in hourly_first_response.items()}
    
    avg_tickets_aht_min = round((total_handle_time_sec / closed_with_aht_count / 60), 2) if closed_with_aht_count > 0 else 0
    
    return {
        "agents": list(agents.values()),
        "tags": list(tags.values()),
        "total_count": total_tickets,
        "hourly_volume": hourly_volume,
        "hourly_closed": hourly_closed,
        "hourly_resp_avg": hourly_resp_avg,
        "queue_dist": queue_dist,
        "reopen_count": reopen_count,
        "old_tickets_closed": old_tickets_closed,
        "avg_aht_min": avg_tickets_aht_min
    }

# ===== Outlook Fetching Integration =====
def fetch_from_outlook(target_date, subfolder_name, file_prefix):
    """משיכת דוח מצורף מאאוטלוק - עם חיפוש חכם לתיקיות"""
    print(f"\n[*] מתחבר לאאוטלוק למשיכת דוח מתיקייה: {subfolder_name}...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        
        # 1. ניווט לתיקת דוחות
        reports_folder = None
        for f in inbox.Folders:
            if f.Name.strip() == "דוחות":
                reports_folder = f
                break
        
        if not reports_folder:
            print("   [!] לא נמצאה תיקיית 'דוחות' ב-Inbox")
            return None
            
        # 2. דוחות -> Five9
        five9_folder = None
        for f in reports_folder.Folders:
            if f.Name.lower().strip() == "five9":
                five9_folder = f
                break
        
        if not five9_folder:
            print("   [!] לא נמצאה תיקיית 'Five9' בתוך 'דוחות'")
            return None
            
        # 3. Five9 -> subfolder (חיפוש גמיש)
        target_folder = None
        for f in five9_folder.Folders:
            if f.Name.lower().strip() == subfolder_name.lower().strip():
                target_folder = f
                break
        
        if not target_folder:
            print(f"   [!] התיקייה '{subfolder_name}' לא נמצאה בתוך Five9.")
            return None
        
        print(f"   [V] אותר תיקיית היעד: {target_folder.Name}")

        items = target_folder.Items
        items.Sort("[ReceivedTime]", True) 
        
        for item in items:
            try:
                if item.Attachments.Count > 0:
                    for att in item.Attachments:
                        fname_lower = att.FileName.lower()
                        if fname_lower.endswith(('.csv', '.xlsx', '.xls')):
                            ext = os.path.splitext(att.FileName)[1]
                            save_path = os.path.join(DOWNLOAD_DIR, f"{file_prefix}_{target_date.strftime('%Y%m%d')}{ext}")
                            
                            try:
                                att.SaveAsFile(save_path)
                                print(f"   [V] נשמר: {save_path}")
                            except Exception as e:
                                if os.path.exists(save_path):
                                    print(f"   [!] הקובץ נעול (פתוח ב-Excel?), משתמש בעותק הקיים...")
                                else:
                                    print(f"   [!] שגיאה בשמירת הקובץ: {e}")
                                    continue
                            return save_path
            except: continue
        

        print(f"   [!] לא נמצא קובץ בתיקייה {subfolder_name}.")
        return None
    except Exception as e:
        print(f"   [!] תקלה באאוטלוק: {e}")
        return None

# ===== Verint Processing =====
def process_duration(dur_str):
    try:
        if pd.isna(dur_str): return 0
        dur_str = str(dur_str)
        parts = dur_str.split(':')
        if len(parts) == 3: return int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        elif len(parts) == 2: return int(parts[0])*60 + int(parts[1])
        return 0
    except: return 0

def analyze_verint_csv(file_path, start_date, is_survey=False, end_date=None):
    print(f"\n[*] מנתח קובץ {'סקרים' if is_survey else 'שיחות'}: {os.path.basename(file_path)}")
    temp_file = None
    try:
        # Create a temporary copy to avoid "Permission Denied" if the file is open
        fd, temp_file = tempfile.mkstemp(suffix=os.path.splitext(file_path)[1])
        os.close(fd)
        shutil.copy2(file_path, temp_file)
        
        if temp_file.lower().endswith(('.xlsx', '.xls')):
            raw_df = pd.read_excel(temp_file, header=None)
            header_row = 0
            found = False
            for i, row in raw_df.iterrows():
                row_str = " ".join([str(val).upper() for val in row if pd.notna(val)])
                if 'AGENT NAME' in row_str or 'TALK TIME' in row_str or 'HANDLE TIME' in row_str:
                    header_row = i
                    found = True
                    break
            df = pd.read_excel(temp_file, skiprows=header_row) if found else raw_df
        else:
            try: df = pd.read_csv(temp_file)
            except: df = pd.read_csv(temp_file, encoding='cp1255')
    except Exception as e:
        print(f"   [!] שגיאה בקריאת הקובץ: {e}")
        return None
    finally:
        if temp_file and os.path.exists(temp_file):
            try: os.remove(temp_file)
            except: pass

    if df.empty: return None
    df = df.dropna(thresh=2)
    
    # Normalize Columns
    df.columns = [str(c).strip().upper() for c in df.columns]

    # Flexible column identification
    emp_col = None
    for c in df.columns:
        if c in ['EMPLOYEE', 'AGENT NAME', 'AGENT', 'נציג']:
            emp_col = c; break
    
    if not emp_col:
        print(f"   [!] לא נמצאה עמודת נציג. עמודות: {list(df.columns)}")
        return None

    # Normalizing names
    def clean_emp_name(name):
        n = str(name).strip()
        if ',' in n:
            parts = n.split(',')
            return f"{parts[1].strip()} {parts[0].strip()}"
        return n

    df[emp_col] = df[emp_col].apply(clean_emp_name)
    df = df[~df[emp_col].astype(str).str.contains('Cnt:|Avg:|Total|Report', na=False, case=False)]

    if is_survey:
        survey_stats = {}
        score_col = None
        for c in df.columns:
            if any(x in str(c).upper() for x in ['SCORE', 'Q1', 'SURVEY', 'ציון']):
                score_col = c; break
        
        if score_col:
            df[score_col] = pd.to_numeric(df[score_col], errors='coerce')
            df = df.dropna(subset=[score_col])
            for name, group in df.groupby(emp_col):
                survey_stats[name] = {
                    'avg': round(group[score_col].mean(), 1),
                    'count': len(group)
                }
        return survey_stats

    # Extracting Duration
    duration_col = None
    for c in ['TALK TIME', 'TALK_TIME', 'HANDLE TIME', 'Interaction Duration', 'Duration']:
        if c in df.columns: duration_col = c; break
    
    if duration_col:
        df['Duration_Sec'] = df[duration_col].apply(process_duration)
    else:
        df['Duration_Sec'] = 0

    # Date Filtering (Robust) - Range Support
    time_col = None
    for col in ['Start Time', 'DATE', 'Date', 'TIMESTAMP MILLISECOND', 'TIME']:
        if col in df.columns: time_col = col; break
    
    if time_col:
        try:
            df[time_col] = pd.to_datetime(df[time_col], errors='coerce')
            
            if end_date:
                # Filter by range
                df_filtered = df[(df[time_col].dt.date >= start_date.date()) & (df[time_col].dt.date <= end_date.date())]
            else:
                # Filter by single date
                df_filtered = df[df[time_col].dt.date == start_date.date()]
                
            if not df_filtered.empty:
                df = df_filtered
                print(f"   [V] סונן לתאריכים: {start_date.date()} - {end_date.date() if end_date else start_date.date()}. שורות: {len(df)}")
            else:
                print(f"   [!] אזהרה: לא נמצאו שורות בטווח התאריכים, משתמש בכל הדוח ({len(df)} שורות).")
        except Exception as ex: 
            print(f"   [!] Date filter error: {ex}")

    # Categorization based on ANI/DNIS (from analyze_calls.py parameters)
    ani_col = next((c for c in df.columns if 'ANI' in c or 'DIALED FROM' in c), None)
    dnis_col = next((c for c in df.columns if 'DNIS' in c or 'DIALED TO' in c), None)
    
    vico = 0; vert = 0; shuf = 0; tier1 = 0
    
    if ani_col and dnis_col:
        # Normalize numbers (remove non-digits and handle .0 floats)
        import re
        def clean_phone(p):
            s = str(p).strip()
            if s.endswith('.0'): s = s[:-2]
            return re.sub(r'\D', '', s)
            
        df[ani_col] = df[ani_col].apply(clean_phone)
        df[dnis_col] = df[dnis_col].apply(clean_phone)
        
        vico = len(df[df[ani_col].str.endswith('39029740', na=False)])
        tier1 = len(df[df[dnis_col].str.endswith('39029740', na=False)])
        vert = len(df[df[dnis_col].str.endswith('732069574', na=False)])
        shuf = len(df[df[dnis_col].str.endswith('732069576', na=False)])
        
        # New Categories Coverage
        vico_out = len(df[df[dnis_col].astype(str).str.contains('97235264646', na=False)])
        vico += vico_out
        
    else:
        # Fallback to campaign mapping
        campaign_col = next((c for c in df.columns if 'CAMPAIGN' in c), None)
        if campaign_col:
            vico = len(df[df[campaign_col].astype(str).str.contains('VICO', na=False, case=False)])
            
            # Analiza -> Tier 1
            tier1_analiza = len(df[df[campaign_col].astype(str).str.contains('Analiza', na=False, case=False)])
            
            vert = len(df[df[campaign_col].astype(str).str.contains('Verticals', na=False, case=False)])
            shuf = len(df[df[campaign_col].astype(str).str.contains('Shufersal', na=False, case=False)])
            
            # Simple Tier 1 deduction
            tier1 = len(df) - (vico + vert + shuf) 
            # If Analiza was not counted in vico/vert/shuf, it falls into remaining, 
            # but let's be explicit if needed. Since Tier 1 is "Rest", Analiza is automatically in Tier 1.
            # No double counting adjustment needed unless Analiza matches VICO as well.
            # VICO filter matches "VICO", so "IL_TLV_VICO_Analiza" MIGHT be counted as VICO.
            # Fix: exclude Analiza from VICO
            
            # Re-calculating strict
            is_vico = df[campaign_col].astype(str).str.contains('VICO', na=False, case=False)
            is_analiza = df[campaign_col].astype(str).str.contains('Analiza', na=False, case=False)
            is_shuf = df[campaign_col].astype(str).str.contains('Shufersal', na=False, case=False)
            is_vert = df[campaign_col].astype(str).str.contains('Verticals', na=False, case=False)
            
            vico = len(df[is_vico & ~is_analiza]) # VICO excluding Analiza
            shuf = len(df[is_shuf])
            vert = len(df[is_vert])
            tier1 = len(df) - (vico + shuf + vert) # Everything else (inc. Analiza)


    # Hourly Distribution
    hourly_calls = {h: 0 for h in range(24)}
    if time_col:
        try:
            for _, row in df.iterrows():
                if pd.notna(row[time_col]):
                    hourly_calls[row[time_col].hour] += 1
        except: pass

    employee_stats = {}
    for name, group in df.groupby(emp_col):
        avg_dur = round(group['Duration_Sec'].mean() / 60, 2)
        employee_stats[name] = {'count': len(group), 'avg_duration': avg_dur}

    return {
        'total_calls': len(df),
        'avg_duration_min': round(df['Duration_Sec'].mean()/60, 2) if len(df)>0 else 0,
        'vico_count': vico, 'tier1_count': tier1, 'vert_count': vert, 'shuf_count': shuf,
        'hourly_calls': hourly_calls,
        'employee_counts': {k: v['count'] for k, v in employee_stats.items()},
        'employee_stats': employee_stats
    }


# ===== גרפים ו-HTML =====
def setup_plt_dark_style():
    plt.style.use('dark_background')
    plt.rcParams.update({'figure.facecolor': '#0f172a', 'axes.facecolor': '#0f172a', 'axes.grid': False})

def _save_fig_to_b64(fig):
    buf = BytesIO()
    fig.savefig(buf, dpi=200, bbox_inches='tight', transparent=True)
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode()

def plot_efficiency_b64(hourly_opened, hourly_closed):
    setup_plt_dark_style()
    hours = sorted(hourly_opened.keys())
    opened = [hourly_opened[h] for h in hours]
    closed = [hourly_closed[h] for h in hours]
    
    fig, ax = plt.subplots(figsize=(10, 4))
    x = range(len(hours))
    ax.bar([i - 0.2 for i in x], opened, width=0.4, label=reshape_hebrew('נפתחו'), color='#3b82f6', align='center')
    ax.bar([i + 0.2 for i in x], closed, width=0.4, label=reshape_hebrew('נסגרו (אותו יום)'), color='#10b981', align='center')
    
    ax.set_xticks(x)
    ax.set_xticklabels(hours)
    ax.legend(facecolor='#1e293b', edgecolor='#334155', labelcolor='white')
    ax.set_title(reshape_hebrew("יעילות טיפול לפי שעה (פתיחה vs סגירה)"), color='white', fontweight='bold')
    plt.tight_layout()
    return _save_fig_to_b64(fig)

def plot_response_trend_b64(hourly_resp):
    setup_plt_dark_style()
    hours = sorted(hourly_resp.keys())
    resp = [hourly_resp[h] for h in hours]
    
    fig, ax = plt.subplots(figsize=(10, 3))
    ax.plot(hours, resp, marker='o', linestyle='-', color='#6366f1', linewidth=3, markersize=8)
    ax.fill_between(hours, resp, color='#6366f1', alpha=0.2)
    
    ax.set_xticks(hours)
    ax.set_title(reshape_hebrew("זמן תגובה ממוצע לפי שעת פתיחה (שעות)"), color='white', fontweight='bold')
    ax.grid(axis='y', linestyle='--', alpha=0.3)
    plt.tight_layout()
    return _save_fig_to_b64(fig)


def plot_agents_bar_b64(agents):
    setup_plt_dark_style()
    agents_sorted = sorted(agents, key=lambda x: x["Total"], reverse=True)
    names = [reshape_hebrew(a["Agent"]) for a in agents_sorted]
    totals = [a["Total"] for a in agents_sorted]
    if not totals: return ""
    
    fig, ax = plt.subplots(figsize=(9, 7))
    bars = ax.bar(names, totals, color=['#8b5cf6', '#ec4899', '#10b981', '#3b82f6'], width=0.6)
    
    for bar in bars:
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height() + (max(totals)*0.02), 
                str(int(bar.get_height())), ha='center', color='white', fontweight='bold', fontsize=14)
    
    ax.set_xticks(range(len(names)))
    ax.set_xticklabels(names, color='white', fontweight='bold', fontsize=12)
    if totals: ax.set_ylim(0, max(totals) * 1.3)
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.yaxis.set_visible(False)
    
    ax.set_title(reshape_hebrew("ביצועי נציגים (מיילים)"), color='white', pad=20, fontweight='bold', fontsize=16)
    plt.tight_layout(pad=3.0)
    return _save_fig_to_b64(fig)

def plot_tags_donut_b64(tags):
    setup_plt_dark_style()
    tags_sorted = sorted(tags, key=lambda x: x["Total"], reverse=True)[:6]
    if not tags_sorted: return ""
    labels = [reshape_hebrew(t["Tag"]) for t in tags_sorted]
    vals = [t["Total"] for t in tags_sorted]
    
    fig, ax = plt.subplots(figsize=(8, 8))
    colors = ['#8b5cf6', '#ec4899', '#10b981', '#f59e0b', '#3b82f6', '#06b6d4']
    
    wedges, texts, autotexts = ax.pie(vals, labels=labels, autopct='%1.1f%%', startangle=140, 
                                      colors=colors, pctdistance=0.85, 
                                      textprops={'color':"w", 'fontweight':'bold', 'fontsize':12},
                                      wedgeprops=dict(width=0.4, edgecolor='#0f172a', linewidth=8))
    
    ax.set_title(reshape_hebrew("פילוח תיוגים"), color='white', pad=10, fontweight='bold', fontsize=16)
    plt.tight_layout()
    return _save_fig_to_b64(fig)

def plot_calls_bar_b64(emp_counts):
    setup_plt_dark_style()
    if not emp_counts: return ""
    sorted_emps = sorted(emp_counts.items(), key=lambda x: x[1], reverse=True)
    names = [reshape_hebrew(n) for n,v in sorted_emps]
    vals = [v for n,v in sorted_emps]
    
    fig, ax = plt.subplots(figsize=(9, 7))
    bars = ax.bar(names, vals, color=['#4f46e5', '#ec4899', '#10b981', '#f59e0b'], width=0.6)
    
    for bar in bars:
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height() + (max(vals)*0.02), 
                str(int(bar.get_height())), ha='center', va='bottom', color='white', fontweight='bold', fontsize=14)
    
    ax.set_xticks(range(len(names)))
    ax.set_xticklabels(names, color='white', fontweight='bold', fontsize=11)
    if vals: ax.set_ylim(0, max(vals) * 1.3)
    
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_visible(False)
    ax.yaxis.set_visible(False)
    
    ax.set_title(reshape_hebrew("ביצועי נציגים (שיחות)"), color='white', pad=20, fontweight='bold', fontsize=16)
    plt.tight_layout(pad=3.0)
    return _save_fig_to_b64(fig)

def plot_heatmap_b64(hourly_data, title="מפת עוצמת עבודה - משולב"):
    """Creates a work intensity heatmap (Hours vs Activity)"""
    import seaborn as sns
    setup_plt_dark_style()
    
    # Prep data for 1x24 heatmap
    hours = sorted(hourly_data.keys())
    values = [hourly_data[h] for h in hours]
    
    if not any(values): return ""
    
    data = [values]
    fig, ax = plt.subplots(figsize=(10, 2))
    sns.heatmap(data, annot=True, fmt="d", cmap="YlGnBu", cbar=False, ax=ax,
                xticklabels=hours, yticklabels=["עוצמה"], annot_kws={"size": 12, "weight": "bold"})
    
    ax.set_title(reshape_hebrew(title), color='white', fontweight='bold', fontsize=14)
    plt.tight_layout()
    return _save_fig_to_b64(fig)

def plot_queue_distribution_b64(queue_dist):
    """Bar chart for queue wait time buckets"""
    setup_plt_dark_style()
    labels = list(queue_dist.keys())
    vals = list(queue_dist.values())
    if not any(vals): return ""
    
    fig, ax = plt.subplots(figsize=(8, 4))
    colors = ['#10b981', '#fbbf24', '#f59e0b', '#ef4444']
    bars = ax.bar(labels, vals, color=colors, width=0.6)
    
    for bar in bars:
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height() + (max(vals)*0.05), 
                str(int(bar.get_height())), ha='center', color='white', fontweight='bold')
    
    ax.set_title(reshape_hebrew("פילוח זמני המתנה בתור"), color='white', fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    return _save_fig_to_b64(fig)


def plot_services_donut_b64(stats):
    if not stats: return ""
    setup_plt_dark_style()
    labels = ['Vico', 'Tier 1', 'Verticals', 'Shufersal']
    vals = [stats['vico_count'], stats['tier1_count'], stats['vert_count'], stats['shuf_count']]
    data = [(l,v) for l,v in zip(labels, vals) if v > 0]
    if not data: return ""
    
    fig, ax = plt.subplots(figsize=(8, 8))
    colors = ['#6366f1', '#ec4899', '#10b981', '#f59e0b']
    
    ax.pie([v for l,v in data], labels=[l for l,v in data], autopct='%1.0f%%', startangle=140, 
           colors=colors, pctdistance=0.85, 
           textprops={'color':"w", 'fontweight':'bold', 'fontsize':12},
           wedgeprops=dict(width=0.4, edgecolor='#0f172a', linewidth=8))
    
    ax.set_title(reshape_hebrew("פילוח שירותים (Service Share)"), color='white', pad=10, fontweight='bold', fontsize=16)
    plt.tight_layout()
    return _save_fig_to_b64(fig)

def plot_aht_trend_b64(hourly_data):
    """Line chart for average response time trend"""
    setup_plt_dark_style()
    hours = sorted(hourly_data.keys())
    values = [hourly_data[h] for h in hours]
    if not any(values): return ""
    
    fig, ax = plt.subplots(figsize=(10, 3))
    ax.plot(hours, values, marker='o', linestyle='-', color='#ec4899', linewidth=2)
    ax.fill_between(hours, values, color='#ec4899', alpha=0.1)
    
    ax.set_xticks(hours)
    ax.set_title(reshape_hebrew("מגמת זמן תגובה ממוצע (בשעות)"), color='white', fontweight='bold')
    plt.tight_layout()
    return _save_fig_to_b64(fig)

# ===== Star Agent & Top 3 Logic =====
def calculate_star_agent(tickets_agents, verint_stats, survey_stats=None, work_days=1):
    """Calculate combined score with weighted formula: 40% Volume, 30% SLA, 20% FCR, 10% CSAT"""
    combined = {}
    survey_stats = survey_stats or {}
    
    # Target Volume Scaling (50 tickets per day)
    target_volume = 50 * max(1, work_days)

    
    def normalize_for_match(name):
        return "".join(c for c in str(name).lower() if c.isalpha())

    # 1. Process Tickets
    for a in tickets_agents:
        name = a["Agent"]
        combined[name] = {
            "Agent": name, "Name": name, 
            "Total": a.get("Total", 0), "Tickets": a.get("Total", 0),
            "Calls": 0, "SLA": a.get("SLA", 0), "FCR": a.get("FCR", 0), 
            "AvgFirstResponseHours": a.get("AvgFirstResponseHours", 0),
            "AvgCallsPerHour": a.get("AvgCallsPerHour", 0),
            "AvgAHTMin": a.get("AvgAHTMin", 0),
            "Survey": 0
        }

    # 2. Process Calls
    if verint_stats and 'employee_stats' in verint_stats:
        ticket_lookup = {normalize_for_match(n): n for n in combined.keys()}
        for v_name, data in verint_stats["employee_stats"].items():
            v_clean = v_name.strip()
            target_name = VERINT_NAME_MAP.get(v_clean)
            if not target_name and normalize_for_match(v_clean) in ticket_lookup:
                target_name = ticket_lookup[normalize_for_match(v_clean)]
            
            v_key = target_name if target_name else v_clean
            if v_key not in combined:
                combined[v_key] = {
                    "Agent": v_key, "Name": v_key, "Total": 0, "Tickets": 0, "Calls": 0, 
                    "SLA": 0, "FCR": 0, "AvgFirstResponseHours": 0, "AvgCallsPerHour": 0, "Survey": 0, "AvgAHTMin": 0
                }
            combined[v_key]["Calls"] = data["count"]
            combined[v_key]["AvgCallDur"] = data["avg_duration"]

    # 3. Apply Survey Scores & Weighted Formula
    for name, stats in combined.items():
        # CSAT Mapping
        s_data = survey_stats.get(name, 0)
        if isinstance(s_data, dict):
            stats["Survey"] = s_data.get('avg', 0)
        else:
            stats["Survey"] = s_data
        
        
        # Scoring components (Scaled 0-100)
        # Dynamic Volume Score based on Target
        vol_score = min((stats["Tickets"] + stats["Calls"]) / float(target_volume) * 100, 100) 
        sla_score = stats["SLA"]
        fcr_score = stats["FCR"]
        csat_score = (stats["Survey"] / 5.0 * 100) if stats["Survey"] else 0
        
        # Final Score
        final_score = (vol_score * 0.4) + (sla_score * 0.3) + (fcr_score * 0.2) + (csat_score * 0.1)
        stats["Score"] = round(final_score, 1)
        stats["TotalActivity"] = stats["Tickets"] + stats["Calls"]

    sorted_agents = sorted(combined.values(), key=lambda x: x["Score"], reverse=True)
    return sorted_agents

def build_star_agent_html(sorted_agents, title_period="היום"):
    if not sorted_agents: return ""
    star = sorted_agents[0]
    return f"""
    <div style="background: linear-gradient(135deg, #FFD700 0%, #FDB931 100%); color: #ffffff; padding: 25px; border-radius: 16px; margin-bottom: 25px; text-align: center; box-shadow: 0 8px 20px rgba(255, 215, 0, 0.4);">
        <div style="font-size: 28px; font-weight: bold; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); margin-bottom: 15px;">⭐ כוכב/ת {title_period} ⭐</div>
        <div style="font-size: 42px; font-weight: 900; margin: 15px 0; text-shadow: 3px 3px 6px rgba(0,0,0,0.4);">{star['Name']}</div>
        <div style="font-size: 18px; background: rgba(255,255,255,0.25); display: inline-block; padding: 12px 25px; border-radius: 25px; font-weight: bold; margin-top: 10px;">
            🏆 סה"כ {star['TotalActivity']} משימות ל{title_period} העבודה! (מיילים: {star['Tickets']}, שיחות: {star['Calls']})
        </div>
    </div>
    """

def build_top3_html(sorted_agents):
    if len(sorted_agents) < 2: return ""
    top3 = sorted_agents[1:4] # Skip #1 (Star Agent) and show next 3
    if not top3: return ""
    
    html = '<div style="display: flex; justify-content: space-around; margin-bottom: 30px;">'
    medals = ["🥈", "🥉", "🎖️"]
    colors = ["#e2e8f0", "#cd7f32", "#475569"]
    
    for i, agent in enumerate(top3):
        medal = medals[i] if i < len(medals) else "🎖️"
        color = colors[i] if i < len(colors) else "#475569"
        html += f"""
        <div style="background-color: #1e293b; border: 2px solid {color}; border-radius: 12px; padding: 15px; width: 30%; text-align: center;">
            <div style="font-size: 24px;">{medal}</div>
            <div style="font-size: 18px; color: #ffffff; font-weight: bold; margin: 10px 0;">{agent['Name']}</div>
            <div style="font-size: 14px; color: #94a3b8;">פעילות: {agent['TotalActivity']}</div>
        </div>
        """
    html += '</div>'
    return html

# ===== HTML Building Blocks =====
def build_tier2_table_html(agents):
    if not agents: return "<p style='color:#94a3b8; text-align:center;'>אין נתוני פניות זמינים</p>"
    
    rows = []
    agents_sorted = sorted(agents, key=lambda x: x["Total"], reverse=True)
    max_total = agents_sorted[0]["Total"] if agents_sorted else 0
    
    for a in agents_sorted:
        badge = '<span style="background-color: #fbbf24; color: #000; font-size: 11px; font-weight: 900; padding: 2px 6px; border-radius: 4px; margin-right: 8px;">TOP 🔥</span>' if a["Total"] == max_total and max_total > 0 else ""
        sla_color = "#10b981" if a["SLA"] >= 90 else "#ef4444"
        fcr_color = "#3b82f6" if a["FCR"] >= 80 else "#ec4899"
        
        rows.append(f"""
        <tr style="border-bottom: 1px solid #1e293b;">
            <td style="padding:12px; font-weight:bold; color:#f8fafc; font-size:15px;">{a['Agent']} {badge} </td>
            <td style="padding:12px; color:#ffffff; text-align:center; font-weight:bold;">{a['Total']}</td>
            <td style="padding:12px; font-weight:bold; color:{sla_color}; text-align:center;">{a['SLA']}%</td>
            <td style="padding:12px; font-weight:bold; color:{fcr_color}; text-align:center;">{a['FCR']}%</td>
            <td style="padding:12px; color:#94a3b8; text-align:center;">{a['AvgFirstResponseHours']}h</td>
            <td style="padding:12px; color:#ffffff; text-align:center; font-weight:800;">{a['AvgCallsPerHour']}</td>
        </tr>
        """)
    
    return f"""
    <table style="width:100%; border-collapse:collapse; text-align:right; direction:rtl; background:#0f172a; border-radius:12px; overflow:hidden; border: 1px solid #1e293b; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.4);">
        <thead>
            <tr style="background: #2563eb; color:#ffffff; font-size:1em;">
                <th style="padding:15px; text-align:right;">נציג</th>
                <th style="padding:15px; text-align:center;">פניות</th>
                <th style="padding:15px; text-align:center;">SLA %</th>
                <th style="padding:15px; text-align:center;">FCR %</th>
                <th style="padding:15px; text-align:center;">תגובה ראשונה</th>
                <th style="padding:15px; text-align:center;">לשעה</th>
            </tr>
        </thead>
        <tbody>
            {"".join(rows)}
        </tbody>
    </table>
    """

def build_tags_table_html(tags):
    if not tags: return "<p style='color:#94a3b8; text-align:center;'>אין נתוני תגיות זמינים</p>"
    
    rows = []
    tag_items = tags.values() if isinstance(tags, dict) else tags
    sorted_tags = sorted(tag_items, key=lambda x: x.get("Total", 0), reverse=True)
    
    for t in sorted_tags[:10]:
        rows.append(f"""
        <tr style="border-bottom: 1px solid #1e293b;">
            <td style="padding:15px; font-weight:bold; color:#f8fafc;">{t.get('Tag', 'Unknown')}</td>
            <td style="padding:15px; color:#94a3b8; text-align:center;">{t.get('Open', 0)}</td>
            <td style="padding:15px; color:#94a3b8; text-align:center;">{t.get('Closed', 0)}</td>
            <td style="padding:15px; color:#38bdf8; text-align:center; font-weight:bold; font-size:16px;">{t.get('Total', 0)}</td>
            <td style="padding:15px; color:#ec4899; text-align:center;">{t.get('AvgFirstResponseHours', 0)}h</td>
        </tr>
        """)
    
    return f"""
    <table style="width:100%; border-collapse:collapse; text-align:right; direction:rtl; background:#0f172a; border-radius:12px; overflow:hidden; border: 1px solid #1e293b; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.4);">
        <thead>
            <tr style="background: #7c3aed; color:#ffffff; font-size:1em;">
                <th style="padding:15px; text-align:right;">תגית</th>
                <th style="padding:15px; text-align:center;">פתוחות</th>
                <th style="padding:15px; text-align:center;">סגורות</th>
                <th style="padding:15px; text-align:center;">סה"כ</th>
                <th style="padding:15px; text-align:center;">זמן תגובה</th>
            </tr>
        </thead>
        <tbody>
            {"".join(rows)}
        </tbody>
    </table>
    """

def build_verint_table_html(stats, surveys=None):
    if not stats or 'employee_stats' not in stats:
        return "<p style='color:#94a3b8; text-align:center;'>אין נתוני שיחות זמינים</p>"
    
    surveys = surveys or {}
    rows = []
    sorted_items = sorted(stats['employee_stats'].items(), key=lambda x: x[1]['count'], reverse=True)
    max_calls = sorted_items[0][1]['count'] if sorted_items else 0
    
    for name, data in sorted_items:
        s_val = surveys.get(name, "-")
        survey_avg = "-"
        survey_count = 0
        
        if isinstance(s_val, dict):
            survey_avg = s_val.get('avg', "-")
            survey_count = s_val.get('count', 0)
        elif isinstance(s_val, (int, float)):
            survey_avg = s_val

        # Color score based on value (Scale 1-5)
        score_html = f'<span style="color:#94a3b8;">{survey_avg}</span>'
        if isinstance(survey_avg, (int, float)):
            if survey_avg >= 4.5: score_html = f'<span style="color:#10b981; font-weight:bold;">{survey_avg} <small style="color:#94a3b8; font-weight:normal;">({survey_count})</small></span>'
            elif survey_avg >= 3: score_html = f'<span style="color:#fbbf24; font-weight:bold;">{survey_avg} <small style="color:#94a3b8; font-weight:normal;">({survey_count})</small></span>'
            else: score_html = f'<span style="color:#ef4444; font-weight:bold;">{survey_avg} <small style="color:#94a3b8; font-weight:normal;">({survey_count})</small></span>'

        badge = '<span style="background-color: #fbbf24; color: #000; font-size: 11px; font-weight: 900; padding: 2px 6px; border-radius: 4px; margin-right: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.3);">TOP 🔥</span>' if data['count'] == max_calls and max_calls > 0 else ""
        cph = round(data['count'] / 8, 2)
        display_name = VERINT_NAME_MAP.get(name, name)
        
        rows.append(f"""
        <tr style="border-bottom: 1px solid #1e293b;">
            <td style="padding:15px; font-weight:bold; color:#f8fafc; font-size:15px;">{display_name} {badge}</td>
            <td style="padding:15px; color:#ffffff; text-align:center; font-weight:bold;">{data['count']}</td>
            <td style="padding:15px; color:#10b981; text-align:center; font-weight:bold;">{data['avg_duration']}m</td>
            <td style="padding:15px; color:#ffffff; text-align:center; font-weight:800;">{cph}</td>
            <td style="padding:15px; text-align:center;">{score_html}</td>
        </tr>
        """)
    
    return f"""
    <table style="width:100%; border-collapse:collapse; text-align:right; direction:rtl; background:#0f172a; border-radius:12px; overflow:hidden; border: 1px solid #1e293b; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.4);">
        <thead>
            <tr style="background: #2563eb; color:#ffffff; font-size:1em;">
                <th style="padding:15px; text-align:right;">נציג</th>
                <th style="padding:15px; text-align:center;">שיחות</th>
                <th style="padding:15px; text-align:center;">ממוצע שיחה</th>
                <th style="padding:15px; text-align:center;">לשעה</th>
                <th style="padding:15px; text-align:center;">ציון סקר (כמות)</th>
            </tr>
        </thead>
        <tbody>
            {"".join(rows)}
        </tbody>
    </table>
    """

def build_full_html(agents_metrics, tags, verint_stats, survey_stats, images, date_str, info={}, trends={}):
    """
    Redesigned HTML report with advanced analytics, gamification and TRENDS.
    """
    # Determine title period (Today vs Month)
    title_period = "החודש" if len(date_str) > 5 else "היום" # Simple heuristic: '01/2026' vs '01/02/2026' or just check if it was called monthly. 
    # Better: check context. But date_str is format. Let's use heuristics or pass it.
    # Actually, subject_prefix is better but not passed here.
    # Let's check date string length. Month is MM/YYYY (7 chars). Day is DD/MM/YYYY (10 chars).
    if len(date_str) < 9: 
         title_period = "החודש"
    
    star_agent_html = build_star_agent_html(agents_metrics, title_period)
    top3_html = build_top3_html(agents_metrics)
    
    total_calls = verint_stats['total_calls'] if verint_stats else 0
    total_tickets = sum(a['Tickets'] for a in agents_metrics if 'Tickets' in a)
    
    # Trend helper
    def get_trend_html(current, prev, lower_is_better=False):
        if prev is None or prev == 0: return ""
        diff = current - prev
        if diff == 0: return '<span style="color:#94a3b8; font-size:12px; margin-right:5px;">●</span>'
        is_good = diff > 0 if not lower_is_better else diff < 0
        color = "#10b981" if is_good else "#ef4444"
        icon = "▲" if diff > 0 else "▼"
        return f'<span style="color:{color}; font-size:12px; margin-right:5px; font-weight:bold;">{icon}{abs(round(diff,1))}</span>'

    # Calculate averages for KPI cards
    fcr_list = [a['FCR'] for a in agents_metrics if 'FCR' in a]
    avg_fcr = round(sum(fcr_list) / len(fcr_list), 1) if fcr_list else 0
    
    sla_list = [a['SLA'] for a in agents_metrics if 'SLA' in a]
    avg_sla = round(sum(sla_list) / len(sla_list), 1) if sla_list else 0
    
    resp_list = [a['AvgFirstResponseHours'] for a in agents_metrics if 'AvgFirstResponseHours' in a]
    avg_resp = round(sum(resp_list) / len(resp_list), 2) if resp_list else 0

    reopen_rate = round((info.get('reopen_count', 0) / max(total_tickets, 1)) * 100, 1)
    old_closed = info.get('old_tickets_closed', 0)
    
    tickets_table = build_tier2_table_html([a for a in agents_metrics if a.get('Tickets', 0) > 0])
    tags_table = build_tags_table_html(tags)
    verint_table = build_verint_table_html(verint_stats, survey_stats)
    
    now_year = datetime.now().year

    return f"""
    <!DOCTYPE html>
    <html lang="he" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>דוח TIER 2 - {date_str}</title>
    </head>
    <body style="margin: 0; padding: 0; background-color: #030712; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #f8fafc;">
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 950px; background-color: #0f172a; margin: 20px auto; border-radius: 20px; overflow: hidden; box-shadow: 0 20px 50px rgba(0, 0, 0, 0.7); border: 1px solid #1e293b;">
            <!-- כותרת ראשית -->
            <tr>
                <td align="center" style="background: linear-gradient(90deg, #1e40af 0%, #3b82f6 100%); padding: 40px 20px; border-bottom: 4px solid #3b82f6;">
                    <h1 style="margin: 0; color: #ffffff; font-size: 46px; font-weight: 900; letter-spacing: 2px; text-shadow: 0 4px 8px rgba(0,0,0,0.4);">
                         דוח ביצועי TIER 2 PRO
                    </h1>
                    <p style="margin: 15px 0 0; color: #dbeafe; font-size: 20px; font-weight: 600;">נכון לתאריך: {date_str}</p>
                </td>
            </tr>
            
            <!-- KPI Cards Row -->
            <tr>
                <td align="center" style="padding: 30px 25px 15px;">
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                            <td align="center" width="23%" style="background-color: #1e3a8a; border-radius: 12px; padding: 20px; border: 1px solid #3b82f6;">
                                <div style="font-size: 14px; color: #dbeafe; font-weight: bold;">מדד FCR 🎯</div>
                                <div style="font-size: 32px; color: #ffffff; font-weight: 900; margin-top: 8px;">{avg_fcr}%</div>
                            </td>
                            <td width="2%">&nbsp;</td>
                            <td align="center" width="23%" style="background-color: #047857; border-radius: 12px; padding: 20px; border: 1px solid #10b981;">
                                <div style="font-size: 14px; color: #d1fae5; font-weight: bold;">SLA % ⏱️</div>
                                <div style="font-size: 32px; color: #ffffff; font-weight: 900; margin-top: 8px;">{avg_sla}%</div>
                            </td>
                            <td width="2%">&nbsp;</td>
                            <td align="center" width="23%" style="background-color: #1e293b; border-radius: 12px; padding: 20px; border: 1px solid #334155;">
                                <div style="font-size: 14px; color: #94a3b8; font-weight: bold;">מיילים (Tickets) 📧</div>
                                <div style="font-size: 32px; color: #ffffff; font-weight: 900; margin-top: 8px;">
                                    {total_tickets} {get_trend_html(total_tickets, trends.get('prev_tickets'))}
                                </div>
                            </td>
                            <td width="2%">&nbsp;</td>
                            <td align="center" width="23%" style="background-color: #1e293b; border-radius: 12px; padding: 20px; border: 1px solid #334155;">
                                <div style="font-size: 14px; color: #94a3b8; font-weight: bold;">שיחות (Calls) 📞</div>
                                <div style="font-size: 32px; color: #ffffff; font-weight: 900; margin-top: 8px;">
                                    {total_calls} {get_trend_html(total_calls, trends.get('prev_calls'))}
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            
            <!-- Star Agent & Top 3 Section -->
            <tr>
                <td style="padding: 15px 25px 30px;">
                    {star_agent_html}
                    <div style="text-align: center; color: #94a3b8; font-weight: bold; margin-bottom: 20px; font-size: 18px;">🏆 נבחרת המצטיינים</div>
                    {top3_html}
                </td>
            </tr>
            
            <!-- Efficiency Analytics Section -->
            <tr>
                <td style="padding: 10px 25px 40px;">
                    <div style="background-color: #0f172a; border: 1px solid #1e293b; border-radius: 16px; padding: 20px; margin-bottom: 30px;">
                        <div style="font-size: 20px; font-weight: bold; color: #60a5fa; margin-bottom: 20px; text-align: right; border-right: 5px solid #2563eb; padding-right: 12px;">📊 ניתוח יעילות ועומסים</div>
                        <img src="data:image/png;base64,{images.get('efficiency', '')}" width="100%" style="display:block; border-radius: 8px; margin-bottom: 25px;" />
                        <img src="data:image/png;base64,{images.get('resp_trend', '')}" width="100%" style="display:block; border-radius: 8px; margin-bottom: 25px;" />
                        
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                                <td width="49%" align="center" style="background-color: #1e293b; border-radius: 12px; padding: 25px; border: 1px solid #334155;">
                                    <div style="font-size: 15px; color: #94a3b8; font-weight: bold;">אחוז פתיחה מחדש 🔄</div>
                                    <div style="font-size: 32px; color: #f43f5e; font-weight: 800; margin-top: 10px;">{reopen_rate}%</div>
                                    <div style="font-size: 12px; color: #64748b; margin-top: 5px;">פניות שנסגרו ונפתחו שוב ע"י הלקוח</div>
                                </td>
                                <td width="2%">&nbsp;</td>
                                <td width="49%" align="center" style="background-color: #1e293b; border-radius: 12px; padding: 25px; border: 1px solid #334155;">
                                    <div style="font-size: 15px; color: #94a3b8; font-weight: bold;">חיסול בקלוג (ישנות) 💣</div>
                                    <div style="font-size: 32px; color: #10b981; font-weight: 800; margin-top: 10px;">{old_closed}</div>
                                    <div style="font-size: 12px; color: #64748b; margin-top: 5px;">קריאות מימים קודמים שנסגרו היום</div>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            
            <!-- Email (Glassix) Detailed Metrics -->
            <tr>
                <td style="padding: 0 25px 40px;">
                    <div style="font-size: 20px; font-weight: bold; color: #60a5fa; margin-bottom: 15px; text-align: right; border-right: 5px solid #2563eb; padding-right: 12px;">📧 פירוט ביצועי מיילים (Glassix)</div>
                    <div style="overflow-x:auto;">{tickets_table}</div>
                    <div style="margin-top: 25px;">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                                <td width="49%" valign="top" style="background-color: #1e293b; border-radius: 8px; padding: 15px;">
                                    <img src="data:image/png;base64,{images.get('agents_bar', '')}" width="100%" style="display:block;" />
                                </td>
                                <td width="2%">&nbsp;</td>
                                <td width="49%" valign="top" style="background-color: #1e293b; border-radius: 8px; padding: 15px;">
                                    <img src="data:image/png;base64,{images.get('tags_pie', '')}" width="100%" style="display:block;" />
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>

            <!-- Verint Detailed Metrics -->
            <tr>
                <td style="padding: 0 25px 40px;">
                    <div style="font-size: 20px; font-weight: bold; color: #10b981; margin-bottom: 15px; text-align: right; border-right: 5px solid #10b981; padding-right: 12px;">📞 פירוט ביצועי שיחות (Verint)</div>
                    
                    <!-- Service Distribution color boxes -->
                    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="margin: 15px 0;">
                        <tr>
                            <td align="center" width="23.5%" style="background-color: #2e1065; border-radius: 8px; padding: 15px; border: 1px solid #4c1d95;">
                                <div style="font-size: 12px; color: #c084fc; font-weight: bold;">VICO</div>
                                <div style="font-size: 24px; color: #ffffff; font-weight: 900; margin-top: 5px;">{verint_stats['vico_count'] if verint_stats else 0}</div>
                            </td>
                            <td width="2%">&nbsp;</td>
                            <td align="center" width="23.5%" style="background-color: #312e81; border-radius: 8px; padding: 15px; border: 1px solid #3730a3;">
                                <div style="font-size: 12px; color: #818cf8; font-weight: bold;">TIER 1</div>
                                <div style="font-size: 24px; color: #ffffff; font-weight: 900; margin-top: 5px;">{verint_stats['tier1_count'] if verint_stats else 0}</div>
                            </td>
                            <td width="2%">&nbsp;</td>
                            <td align="center" width="23.5%" style="background-color: #064e3b; border-radius: 8px; padding: 15px; border: 1px solid #065f46;">
                                <div style="font-size: 12px; color: #34d399; font-weight: bold;">VERTICALS</div>
                                <div style="font-size: 24px; color: #ffffff; font-weight: 900; margin-top: 5px;">{verint_stats['vert_count'] if verint_stats else 0}</div>
                            </td>
                            <td width="2%">&nbsp;</td>
                            <td align="center" width="23.5%" style="background-color: #451a03; border-radius: 8px; padding: 15px; border: 1px solid #78350f;">
                                <div style="font-size: 12px; color: #fbbf24; font-weight: bold;">SHUFERSAL</div>
                                <div style="font-size: 24px; color: #ffffff; font-weight: 900; margin-top: 5px;">{verint_stats['shuf_count'] if verint_stats else 0}</div>
                            </td>
                        </tr>
                    </table>
                    
                    <div style="overflow-x:auto;">{verint_table}</div>
                    
                    <div style="margin-top: 25px;">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                                <td width="49%" valign="top" style="background-color: #1e293b; border-radius: 8px; padding: 15px;">
                                    <img src="data:image/png;base64,{images.get('calls_bar', '')}" width="100%" style="display:block;" />
                                </td>
                                <td width="2%">&nbsp;</td>
                                <td width="49%" valign="top" style="background-color: #1e293b; border-radius: 8px; padding: 15px;">
                                    <img src="data:image/png;base64,{images.get('services_donut', '')}" width="100%" style="display:block;" />
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>

            <!-- Tags Section -->
            <tr>
                <td style="padding: 0 25px 40px;">
                    <div style="font-size: 20px; font-weight: bold; color: #a855f7; margin-bottom: 15px; text-align: right; border-right: 5px solid #7c3aed; padding-right: 12px;">🏷️ פילוח טיפול לפי נושאים (Tags)</div>
                    <div style="overflow-x:auto;">{tags_table}</div>
                </td>
            </tr>
            
            <!-- Footer -->
            <tr>
                <td align="center" style="background-color: #020617; padding: 35px; border-top: 1px solid #1e293b; color: #64748b; font-size: 14px;">
                    <p style="margin: 0;">🚀 הופק באופן אוטומטי ע"י TIER 2 PERSONAL REPORTER PRO</p>
                    <p style="margin: 8px 0 0;">&copy; {now_year} Verifone Digital Support Team</p>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """


def send_email_html(html, date_str, subject_prefix="דוח יומי"):
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_TO
        mail.CC = EMAIL_CC
        mail.Subject = f"Tier 2 - {subject_prefix} - {date_str}"
        mail.HTMLBody = html
        mail.Send()
        print(f"   [V] מייל נשלח בהצלחה ל: {EMAIL_TO}")
    except Exception as e:
        print(f"   [!] שגיאה בשליחת מייל: {e}")

# ===== Main Entry Point (Automated) =====
# ===== Core Report Execution Function =====
def generate_and_send_report(start_date, end_date, subject_prefix, report_display_title, date_str, total_work_days, trends_enabled=True):
    since = start_date.strftime("%d/%m/%Y 00:00:00:00")
    until = end_date.strftime("%d/%m/%Y 23:59:59:00")

    print(f"\n[*] מפיק {report_display_title} עבור טווח: {since} עד {until}")

    tickets_info = {"agents": [], "tags": [], "total_count": 0, "hourly_volume": {}, "hourly_closed": {}, "queue_dist": {}}
    verint_stats = None
    survey_stats = None
    trends = {}

    # --- Yearly Optimization: Try to load from local Excel files ---
    if "שנתי" in subject_prefix:
        print(f"   [*] דוח שנתי: מנסה לאסוף נתונים מקבצים מקומיים בתיקיית TIER2...")
        target_year = str(start_date.year)
        all_tickets_data = []
        try:
            local_files = [f for f in os.listdir(TIER2_DIR) if f.endswith(".xlsx") and target_year in f and "חודשי" in f]
            for f in local_files:
                f_path = os.path.join(TIER2_DIR, f)
                df = pd.read_excel(f_path)
                all_tickets_data.extend(df.to_dict('records'))
            
            if all_tickets_data:
                tickets_info = parse_tickets(all_tickets_data, work_days=total_work_days)
                print(f"   [V] נטענו בהצלחה {len(all_tickets_data)} פניות מהארכיון המקומי.")
        except Exception as e:
            print(f"   [!] אזהרה בטעינה מקומית: {e}")

    # --- API & Outlook Fetch if needed ---
    if not tickets_info.get("total_count"):
        try:
            token = get_access_token()
            tickets = get_glassix_tickets(token, since, until)
            tickets_info = parse_tickets(tickets, work_days=total_work_days)
            print(f"   [V] Glassix: {len(tickets)} פניות נמשכו.")
            
            # Archive Tickets
            safe_date = date_str.replace("/", "_").replace(" ", "_")
            t_filename = f"Tickets_{subject_prefix.replace(' ', '_')}_{safe_date}.xlsx"
            pd.DataFrame(tickets).to_excel(os.path.join(TIER2_DIR, t_filename), index=False)
            print(f"   [V] נתוני Tickets נשמרו ב-TIER2.")
        except Exception as e:
            print(f"   [!] Glassix Error: {e}")

    # Fetch Verint (Calls)
    if "חודשי" in subject_prefix:
        # Use specific monthly folders under Five9
        calls_csv = fetch_from_outlook(start_date, "Log_VICO_Monthly", "Verint_Calls_Monthly")
        survey_xlsx = fetch_from_outlook(start_date, "Scheduled Report: Survey Result_new_MONTHLY", "Verint_Survey_Monthly")
    else:
        calls_csv = fetch_from_outlook(start_date, "Call Log_VICO", "Verint_Calls")
        survey_xlsx = fetch_from_outlook(start_date, "Survey", "Verint_Survey")


    if calls_csv:
        verint_stats = analyze_verint_csv(calls_csv, start_date, is_survey=False, end_date=end_date)
        if verint_stats:
            print(f"   [V] Verint: {verint_stats['total_calls']} שיחות נותחו.")
            # Archive Calls
            safe_date = date_str.replace("/", "_").replace(" ", "_")
            c_filename = f"Calls_{subject_prefix.replace(' ', '_')}_{safe_date}.xlsx"
            # Since analyze_verint_csv returns a dict of stats, we re-parse or just save the file
            shutil.copy2(calls_csv, os.path.join(TIER2_DIR, c_filename))

    if survey_xlsx:
        survey_stats = analyze_verint_csv(survey_xlsx, start_date, is_survey=True, end_date=end_date)

    # --- Trend calculation (Yesterday) ---
    if trends_enabled and subject_prefix == "דוח יומי":
        prev_date = start_date - timedelta(days=1)
        p_since = prev_date.strftime("%d/%m/%Y 00:00:00:00")
        p_until = prev_date.strftime("%d/%m/%Y 23:59:59:00")
        try:
            print(f"   [*] שולף נתוני השוואה לאתמול...")
            p_token = get_access_token()
            p_tickets = get_glassix_tickets(p_token, p_since, p_until)
            p_info = parse_tickets(p_tickets, work_days=1)
            
            # Simple trend storage
            trends = {
                'prev_total': p_info['total_count'],
                'prev_tickets': p_info['total_count']
            }
            
            # Try to get prev calls if file exists
            p_calls_filename = f"Calls_דוח_יומי_{prev_date.strftime('%d_%m_%Y')}.xlsx"
            p_calls_path = os.path.join(TIER2_DIR, p_calls_filename)
            if os.path.exists(p_calls_path):
                p_verint = analyze_verint_csv(p_calls_path, prev_date)
                if p_verint:
                    trends['prev_calls'] = p_verint['total_calls']
        except: pass

    # 3. Processing Star Agent & Top 3
    print("\n[Step 3] Calculating Performance Gamification...")
    sorted_agents = calculate_star_agent(tickets_info["agents"], verint_stats, survey_stats, work_days=total_work_days)
    

    # 4. Generating Graphs
    print("\n[Step 4] Generating Advanced Visualizations...")
    images = {}
    images['agents_bar'] = plot_agents_bar_b64(tickets_info["agents"])
    images['tags_pie'] = plot_tags_donut_b64(tickets_info["tags"])
    images['efficiency'] = plot_efficiency_b64(tickets_info["hourly_volume"], tickets_info["hourly_closed"])
    images['resp_trend'] = plot_response_trend_b64(tickets_info["hourly_resp_avg"])
    
    if verint_stats:
        images['calls_bar'] = plot_calls_bar_b64(verint_stats['employee_counts'])
        images['services_donut'] = plot_services_donut_b64(verint_stats)

    # 5. Build & Send
    print("\n[Step 5] Building HTML & Sending Final Email...")
    # Note: build_full_html needs to be updated to handle 'trends'
    html = build_full_html(sorted_agents, tickets_info["tags"], verint_stats, survey_stats, images, date_str, tickets_info, trends=trends)
    send_email_html(html, date_str, subject_prefix=subject_prefix)
    
    # --- NEW: Archive HTML Report for Dashboard ---
    try:
        reports_dir = os.path.join(TIER2_DIR, "Reports")
        if not os.path.exists(reports_dir): os.makedirs(reports_dir)
        safe_date = date_str.replace("/", "_").replace(" ", "_")
        html_filename = f"Report_{subject_prefix.replace(' ', '_')}_{safe_date}.html"
        with open(os.path.join(reports_dir, html_filename), "w", encoding="utf-8") as f:
            f.write(html)
        print(f"   [V] עותק דף דוח נשמר ב-TIER2/Reports.")
    except: pass
    
    print(f"\n[SUCCESS] {report_display_title} הופק ונשלח בהצלחה!")

def main():
    print("="*60)
    print("   TIER 2 PERSONAL REPORTER - AUTOMATED")
    print("="*60)

    now = datetime.now()
    days_back = 2 if now.weekday() == 6 else 1
    yesterday = now - timedelta(days=days_back)
    
    # 1. דוח יומי
    print("\n[1/3] מפיק דוח יומי...")
    generate_and_send_report(
        start_date=yesterday, 
        end_date=yesterday, 
        subject_prefix="דוח יומי", 
        report_display_title="דוח ביצועי Tier 2 יומי", 
        date_str=yesterday.strftime('%d/%m/%Y'), 
        total_work_days=1
    )

    # 2. האם ה-1 לחודש?
    if now.day == 1:
        print("\n[2/3] היום ה-1 לחודש - מפיק דוח חודשי...")
        first_day_this_month = now.replace(day=1)
        last_day_prev_month = first_day_this_month - timedelta(days=1)
        start_date_m = last_day_prev_month.replace(day=1)
        
        # חישוב ימי עבודה
        days_in_month = (last_day_prev_month - start_date_m).days + 1
        work_days_m = sum(1 for i in range(days_in_month) if (start_date_m + timedelta(days=i)).weekday() not in [4, 5])
        
        generate_and_send_report(
            start_date=start_date_m, 
            end_date=last_day_prev_month, 
            subject_prefix="דוח חודשי", 
            report_display_title="דוח ביצועי Tier 2 חודשי", 
            date_str=start_date_m.strftime("%m/%Y"), 
            total_work_days=work_days_m
        )

        # 3. האם ה-1 לינואר?
        if now.month == 1:
            print("\n[3/3] היום ה-1 לינואר - מפיק דוח שנתי...")
            last_year = now.year - 1
            start_date_y = datetime(last_year, 1, 1)
            end_date_y = datetime(last_year, 12, 31)
            
            days_in_year = (end_date_y - start_date_y).days + 1
            work_days_y = sum(1 for i in range(days_in_year) if (start_date_y + timedelta(days=i)).weekday() not in [4, 5])

            generate_and_send_report(
                start_date=start_date_y, 
                end_date=end_date_y, 
                subject_prefix="דוח שנתי", 
                report_display_title="דוח ביצועי Tier 2 שנתי", 
                date_str=str(last_year), 
                total_work_days=work_days_y
            )

    print("\n[FINISH] כל הדוחות המתוזמנים נשלחו בהצלחה.")

if __name__ == "__main__":
    main()
