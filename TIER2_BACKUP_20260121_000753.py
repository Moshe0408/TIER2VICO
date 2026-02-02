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
EMAIL_TO = "liorb5@verifone.com; danv1@verifone.com; niv.arieli@verifone.com; moshei1@verifone.com"
EMAIL_CC = "jonas.maman@verifone.com; nadav.lieber@verifone.com"

TIER2_MAP = {
    "niv.arieli": "ניב אריאלי",
    "din.weissman": "דין וייסמן",
    "lior.burstein": "ליאור בורשטיין",
}

DELAY_SECONDS = 120

# ===== הגדרות כלליות =====
DOWNLOAD_DIR = os.path.join(os.getcwd(), "Verint_Reports")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

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

    for t in tickets:
        state = (t.get("state") or "").lower()
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

        tags_list = t.get("tags", [])
        if isinstance(tags_list, str): tags_list = [x.strip() for x in tags_list.split(",") if x.strip()]
        elif not isinstance(tags_list, list): tags_list = []

        if agent_key not in agents:
            agents[agent_key] = {
                "AgentKey": agent_key, "Agent": agent_display,
                "Open": 0, "Closed": 0, "Snoozed": 0, "Other": 0, "Total": 0,
                "SLA": 0, "FCR": 0, "AvgFirstResponseHours": 0, "AvgCallsPerHour": 0
            }
        agents[agent_key]["Total"] += 1
        if state == "open": agents[agent_key]["Open"] += 1
        elif state == "closed": agents[agent_key]["Closed"] += 1
        elif state == "snoozed": agents[agent_key]["Snoozed"] += 1
        else: agents[agent_key]["Other"] += 1

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
                dt_cust = datetime.fromisoformat(first_customer.replace("Z", "+00:00"))
                dt_agent = datetime.fromisoformat(first_agent.replace("Z", "+00:00"))
                effective_start = dt_cust
                is_weekend = dt_cust.weekday() in [4, 5]
                is_outside_hours = dt_cust.hour < 8 or dt_cust.hour >= 17
                if is_weekend or is_outside_hours or dt_cust.date() < dt_agent.date():
                    effective_start = dt_agent.replace(hour=8, minute=0, second=0, microsecond=0)
                diff_hours = (dt_agent - effective_start).total_seconds() / 3600
                if diff_hours < 0: diff_hours = 0
                if 0 <= diff_hours <= 1000:
                    first_response_times[agent_key].append(diff_hours)
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

    for tag_name, stat in tags.items():
        stat["Total"] = stat["Open"] + stat["Closed"] + stat["Snoozed"] + stat["Other"]
        stat["Share"] = round((stat["Total"] / total_tickets) * 100, 1) if total_tickets > 0 else 0
        times = first_response_times_tags.get(tag_name, [])
        stat["AvgFirstResponseHours"] = round(sum(times) / len(times), 2) if times else 0
    
    return list(agents.values()), list(tags.values())

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

def analyze_verint_csv(file_path, start_date, is_survey=False):
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
    
    # Flexible column identification
    emp_col = None
    for c in df.columns:
        if str(c).upper() in ['EMPLOYEE', 'AGENT NAME', 'AGENT', 'נציג']:
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
                survey_stats[name] = round(group[score_col].mean(), 1)
        return survey_stats

    # Extracting Duration
    duration_col = None
    for c in ['TALK TIME', 'TALK_TIME', 'HANDLE TIME', 'Interaction Duration', 'Duration']:
        if c in df.columns: duration_col = c; break
    
    if duration_col:
        df['Duration_Sec'] = df[duration_col].apply(process_duration)
    else:
        df['Duration_Sec'] = 0

    # Date Filtering (Robust)
    time_col = None
    for col in ['Start Time', 'DATE', 'Date', 'TIMESTAMP MILLISECOND', 'TIME']:
        if col in df.columns: time_col = col; break
    
    if time_col:
        try:
            df[time_col] = pd.to_datetime(df[time_col], errors='coerce')
            df_filtered = df[df[time_col].dt.date == start_date.date()]
            if not df_filtered.empty:
                df = df_filtered
            else:
                print(f"   [!] אזהרה: לא נמצאו שורות לתאריך {start_date.date()}, משתמש בכל הדוח.")
        except: pass

    # Campaign Mapping
    campaign_col = 'Campaign' if 'Campaign' in df.columns else 'CAMPAIGN'
    vico = 0; vert = 0; shuf = 0; tier1 = 0
    
    if campaign_col in df.columns:
        vico = len(df[df[campaign_col] == 'IL_TLV_VICO_Service'])
        vert = len(df[df[campaign_col] == 'IL_TLV_Verticals'])
        shuf = len(df[df[campaign_col] == 'IL_TLV_Shufersal'])
        tier1 = len(df) - (vico + vert + shuf)

    employee_stats = {}
    for name, group in df.groupby(emp_col):
        avg_dur = round(group['Duration_Sec'].mean() / 60, 2)
        employee_stats[name] = {'count': len(group), 'avg_duration': avg_dur}

    return {
        'total_calls': len(df),
        'avg_duration_min': round(df['Duration_Sec'].mean()/60, 2) if len(df)>0 else 0,
        'vico_count': vico, 'tier1_count': tier1, 'vert_count': vert, 'shuf_count': shuf,
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
        survey_score = surveys.get(name, "-")
        # Color score based on value (Scale 1-5)
        score_html = f'<span style="color:#94a3b8;">{survey_score}</span>'
        if isinstance(survey_score, (int, float)):
            if survey_score >= 4.5: score_html = f'<span style="color:#10b981; font-weight:bold;">{survey_score}</span>'
            elif survey_score >= 3: score_html = f'<span style="color:#fbbf24; font-weight:bold;">{survey_score}</span>'
            else: score_html = f'<span style="color:#ef4444; font-weight:bold;">{survey_score}</span>'

        badge = '<span style="background-color: #fbbf24; color: #000; font-size: 11px; font-weight: 900; padding: 2px 6px; border-radius: 4px; margin-right: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.3);">TOP 🔥</span>' if data['count'] == max_calls and max_calls > 0 else ""
        cph = round(data['count'] / 8, 2)
        
        rows.append(f"""
        <tr style="border-bottom: 1px solid #1e293b;">
            <td style="padding:15px; font-weight:bold; color:#f8fafc; font-size:15px;">{name} {badge}</td>
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
                <th style="padding:15px; text-align:center;">ציון סקר</th>
            </tr>
        </thead>
        <tbody>
            {"".join(rows)}
        </tbody>
    </table>
    """

def build_full_html(agents, tags, verint_stats, survey_stats, img_agents_b64, img_tags_b64, img_calls_b64, img_services_donut, date_str):
    total_calls = verint_stats['total_calls'] if verint_stats else 0
    avg_call_dur = verint_stats['avg_duration_min'] if verint_stats else 0
    
    total_vol = sum(a['Total'] for a in agents)
    avg_resp = round(sum(a['AvgFirstResponseHours'] for a in agents)/len(agents), 2) if agents else 0
    avg_fcr = round(sum(a['FCR'] for a in agents)/len(agents), 1) if agents else 0
    
    tier2_table = build_tier2_table_html(agents)
    tags_table = build_tags_table_html(tags)
    verint_table = build_verint_table_html(verint_stats, survey_stats)
    
    now_year = datetime.now().year
    report_type = "יומי"

    return f"""
    <!DOCTYPE html>
    <html lang="he" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>דוח ביצועי {report_type} - Tier 2</title>
    </head>
    <body style="margin: 0; padding: 0; background-color: #030712; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 900px; background-color: #0f172a; margin: 20px auto; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.6); border: 1px solid #1e293b;">
            <!-- כותרת ראשית -->
            <tr>
                <td align="center" style="background-color: #1e40af; padding: 35px 20px; border-bottom: 4px solid #3b82f6;">
                    <h1 style="margin: 0; color: #ffffff; font-size: 42px; font-weight: 900; letter-spacing: 1.5px; text-shadow: 0 2px 4px rgba(0,0,0,0.3);">
                         דוח ביצועי יומי - Tier 2
                    </h1>
                    <p style="margin: 10px 0 0; color: #bfdbfe; font-size: 18px; font-weight: 600;">נכון לתאריך: {date_str}</p>
                </td>
            </tr>
            
            <!-- KPI Cards - Order like Image 0 -->
            <tr>
                <td align="center" style="padding: 30px 25px;">
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                            <!-- FCR Index -->
                            <td align="center" width="31%" style="background-color: #1e3a8a; border-radius: 4px; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                                <div style="font-size: 14px; color: #dbeafe; font-weight: bold;">מדד FCR 🎯</div>
                                <div style="font-size: 42px; color: #ffffff; font-weight: 900; margin-top: 10px;">{avg_fcr}%</div>
                            </td>
                            <td width="3.5%">&nbsp;</td>
                            
                            <!-- Response Time -->
                            <td align="center" width="31%" style="background-color: #047857; border-radius: 4px; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                                <div style="font-size: 14px; color: #d1fae5; font-weight: bold;">זמן תגובה ⏳</div>
                                <div style="font-size: 42px; color: #ffffff; font-weight: 900; margin-top: 10px;">{avg_resp}h</div>
                            </td>
                            <td width="3.5%">&nbsp;</td>

                             <!-- Total Tickets -->
                            <td align="center" width="31%" style="background-color: #4338ca; border-radius: 4px; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                                <div style="font-size: 14px; color: #e0e7ff; font-weight: bold;">סה"כ פניות 📩</div>
                                <div style="font-size: 42px; color: #ffffff; font-weight: 900; margin-top: 10px;">{total_vol}</div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            
            <!-- טבלת פניות -->
            <tr>
                <td style="padding: 0 25px 30px;">
                    <div style="font-size: 18px; font-weight: bold; color: #60a5fa; margin-bottom: 15px; text-align: right; border-right: 4px solid #2563eb; padding-right: 12px;">
                       📊 פירוט יעדים ומדדים (צוות T2&lrm;(
                    </div>
                    <div style="overflow-x:auto;">{tier2_table}</div>
                </td>
            </tr>
            
            <!-- גרפים מיילים -->
            <tr>
                <td style="padding: 30px 25px;">
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                             <td align="center" valign="top" width="48%" style="background-color: #1e293b; padding: 20px; border-radius: 4px; border: 1px solid #334155;">
                                <img src="data:image/png;base64,{img_agents_b64}" width="100%" style="display: block; max-width: 100%; height: auto;">
                            </td>
                            <td width="4%">&nbsp;</td>
                            <td align="center" valign="top" width="48%" style="background-color: #1e293b; padding: 20px; border-radius: 4px; border: 1px solid #334155;">
                                <img src="data:image/png;base64,{img_tags_b64}" width="100%" style="display: block; max-width: 100%; height: auto;">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            
            <!-- טבלת תגיות -->
            <tr>
                <td style="padding: 0 25px 50px;">
                    <div style="font-size: 18px; font-weight: bold; color: #a78bfa; margin-bottom: 15px; text-align: right; border-right: 4px solid #7c3aed; padding-right: 12px;">
                        🏷️ פירוט לפי תגיות (נושאים)
                    </div>
                    <div style="overflow-x:auto;">{tags_table}</div>
                </td>
            </tr>
            
            <!-- כותרת Verint -->
            <tr>
                <td align="center" style="background-color: #1d4ed8; padding: 25px; border-bottom: 4px solid #3b82f6;">
                    <h2 style="margin: 0; color: #ffffff; font-size: 32px; font-weight: 800; letter-spacing: 1px;">
                        📞 דוח מדדי שיחות
                    </h2>
                </td>
            </tr>
            
            <!-- KPI Cards - Verint Order like Image 3 -->
            <tr>
                <td align="center" style="padding: 30px 25px;">
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                            <!-- Active Agents -->
                            <td align="center" width="31%" style="background-color: #1e3a8a; border-radius: 4px; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                                <div style="font-size: 14px; color: #dbeafe; font-weight: bold;">נציגים פעילים 👥</div>
                                <div style="font-size: 42px; color: #ffffff; font-weight: 900; margin-top: 10px;">{len(verint_stats['employee_counts']) if verint_stats else 0}</div>
                            </td>
                            <td width="3.5%">&nbsp;</td>

                             <!-- Avg duration -->
                            <td align="center" width="31%" style="background-color: #047857; border-radius: 4px; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                                <div style="font-size: 14px; color: #d1fae5; font-weight: bold;">ממוצע שיחה ⏱️</div>
                                <div style="font-size: 42px; color: #ffffff; font-weight: 900; margin-top: 10px;">{avg_call_dur}</div>
                            </td>
                            <td width="3.5%">&nbsp;</td>
                            
                            <!-- Total calls -->
                            <td align="center" width="31%" style="background-color: #6b21a8; border-radius: 4px; padding: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.3);">
                                <div style="font-size: 14px; color: #f3e8ff; font-weight: bold;">סה"כ שיחות 📞</div>
                                <div style="font-size: 42px; color: #ffffff; font-weight: 900; margin-top: 10px;">{total_calls}</div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            
            <!-- פילוח שירותים Boxes (Image 3 style) -->
            <tr>
                <td style="padding: 0 25px 30px;">
                    <div style="font-size: 16px; font-weight: bold; color: #10b981; margin-bottom: 15px; text-align: center;">
                        📊 פילוח שיחות לפי שירותים (Service Share)
                    </div>
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                            <!-- Vico -->
                            <td align="center" width="23.5%" style="background-color: #0f172a; border-radius: 4px; padding: 15px; border: 1px solid #1e293b;">
                                <div style="font-size: 12px; color: #94a3b8; font-weight: bold;">VICO</div>
                                <div style="font-size: 28px; color: #6366f1; font-weight: 900; margin-top: 5px;">{verint_stats['vico_count'] if verint_stats else 0}</div>
                            </td>
                            <td width="2%">&nbsp;</td>

                             <!-- Tier 1 -->
                            <td align="center" width="23.5%" style="background-color: #0f172a; border-radius: 4px; padding: 15px; border: 1px solid #1e293b;">
                                <div style="font-size: 12px; color: #94a3b8; font-weight: bold;">TIER 1</div>
                                <div style="font-size: 28px; color: #ec4899; font-weight: 900; margin-top: 5px;">{verint_stats['tier1_count'] if verint_stats else 0}</div>
                            </td>
                            <td width="2%">&nbsp;</td>
                            
                            <!-- Verticals -->
                            <td align="center" width="23.5%" style="background-color: #0f172a; border-radius: 4px; padding: 15px; border: 1px solid #1e293b;">
                                <div style="font-size: 12px; color: #94a3b8; font-weight: bold;">VERTICALS</div>
                                <div style="font-size: 28px; color: #10b981; font-weight: 900; margin-top: 5px;">{verint_stats['vert_count'] if verint_stats else 0}</div>
                            </td>
                            <td width="2%">&nbsp;</td>

                            <!-- Shufersal -->
                            <td align="center" width="23.5%" style="background-color: #0f172a; border-radius: 4px; padding: 15px; border: 1px solid #1e293b;">
                                <div style="font-size: 12px; color: #94a3b8; font-weight: bold;">SHUFERSAL</div>
                                <div style="font-size: 28px; color: #f59e0b; font-weight: 900; margin-top: 5px;">{verint_stats['shuf_count'] if verint_stats else 0}</div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>

            <!-- טבלת נציגים (Verint) -->
            <tr>
                <td style="padding: 0 25px 30px;">
                    <div style="font-size: 18px; font-weight: bold; color: #60a5fa; margin-bottom: 15px; text-align: right; border-right: 4px solid #3b82f6; padding-right: 12px;">
                        📊 פירוט יעדים ומדדים (שיחות)
                    </div>
                    <div style="overflow-x:auto;">{verint_table}</div>
                </td>
            </tr>

            <!-- גרפים שיחות -->
            <tr>
                <td style="padding: 0 25px 50px;">
                    <table border="0" cellpadding="0" cellspacing="0" width="100%">
                        <tr>
                            <td align="center" valign="top" width="48%" style="background-color: #1e293b; border-radius: 4px; border: 1px solid #334155; padding: 20px;">
                                <img src="data:image/png;base64,{img_calls_b64}" width="100%" style="display: block; max-width: 100%; height: auto;">
                            </td>
                            <td width="4%">&nbsp;</td>
                            <td align="center" valign="top" width="48%" style="background-color: #1e293b; border-radius: 4px; border: 1px solid #334155; padding: 20px;">
                                <img src="data:image/png;base64,{img_services_donut}" width="100%" style="display: block; max-width: 100%; height: auto;">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            
            <!-- Footer -->
            <tr>
                <td align="center" style="background-color: #020617; padding: 30px; border-top: 1px solid #334155; color: #64748b; font-size: 12px;">
                    <p style="margin: 0;">🚀 הופק באופן אוטומטי ע"י Verint & Glassix Reporter PRO</p>
                    <p style="margin: 5px 0 0;">&copy; {now_year} Verifone Tier 2 Support Team</p>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """

def send_email_html(html, date_str):
    try:
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = EMAIL_TO
        mail.CC = EMAIL_CC
        mail.Subject = f"Tier 2 - דוח יומי - {date_str}"
        mail.HTMLBody = html
        mail.Send()
        print(f"   [V] מייל נשלח בהצלחה ל: {EMAIL_TO}")
    except Exception as e:
        print(f"   [!] שגיאה בשליחת מייל: {e}")

# ===== Main Entry Point =====
def main():
    print("="*60)
    print("   COMBINED REPORTER - OUTLOOK PRO EDITION")
    print("   Tickets (Glassix) + Calls (Outlook Verint)")
    print(f"   Sending test email to: {EMAIL_TO}")
    print("="*60)

    now = datetime.now()
    # ביום ראשון מדווחים על יום שישי (2 ימים אחורה), בשאר הימים על אתמול (1 יום אחורה)
    days_back = 2 if now.weekday() == 6 else 1
    report_date = now - timedelta(days=days_back)
    date_str = report_date.strftime('%d/%m/%Y')

    print(f"[*] היום יום {now.strftime('%A')}, מפיק דוח עבור יום {report_date.strftime('%A')} ({date_str})")

    # 1. Glassix
    print("\n[Step 1] Fetching Glassix Tickets...")
    try:
        since = report_date.strftime("%d/%m/%Y 00:00:00:00")
        until = report_date.strftime("%d/%m/%Y 23:59:59:00")
        token = get_access_token()
        tickets = get_glassix_tickets(token, since, until)
        agents, tags = parse_tickets(tickets)
        print(f"   [V] Glassix: {len(tickets)} פניות נמשכו.")
    except Exception as e:
        print(f"   [!] Glassix Error: {e}")
        agents, tags = [], []

    # 2. Verint Reports (Outlook)
    print("\n[Step 2] Fetching Verint Reports from Outlook...")
    calls_csv = fetch_from_outlook(report_date, "Call Log_VICO", "Verint_Calls")
    survey_xlsx = fetch_from_outlook(report_date, "Survey", "Verint_Survey")

    verint_stats = None
    if calls_csv:
        verint_stats = analyze_verint_csv(calls_csv, report_date)
        if verint_stats:
            print(f"   [V] Verint: {verint_stats['total_calls']} שיחות נותחו.")
        else:
            print("   [!] כשל בניתוח קובץ ה-Verint.")
    else:
        print("   [!] לא נמצא דוח שיחות בתיקייה המיועדת.")

    survey_stats = None
    if survey_xlsx:
        survey_stats = analyze_verint_csv(survey_xlsx, report_date, is_survey=True)
        if survey_stats:
            print(f"   [V] סקרים: {len(survey_stats)} נציגים עם ציונים נמצאו.")

    # 3. Generating Graphs
    print("\n[Step 3] Generating Graphs...")
    img_agents = plot_agents_bar_b64(agents)
    img_tags = plot_tags_donut_b64(tags)
    img_calls = plot_calls_bar_b64(verint_stats['employee_counts']) if (verint_stats and 'employee_counts' in verint_stats) else ""
    img_services = plot_services_donut_b64(verint_stats) if verint_stats else ""

    # 4. Build & Send
    print("\n[Step 4] Building HTML & Sending Email...")
    html = build_full_html(agents, tags, verint_stats, survey_stats, img_agents, img_tags, img_calls, img_services, date_str)
    send_email_html(html, date_str)
    
    print("\n[SUCCESS] הדוח המשולב הופק ונשלח!")

if __name__ == "__main__":
    main()
