import pandas as pd
import os
import json
import webbrowser
from datetime import datetime

# CONFIGURATION
VICO_ANI = '97239029740'
TIER1_DNIS = '97239029740'
VERTICALS_DNIS = '972732069574'
SHUFERSAL_DNIS = '972732069576'

def dur_to_sec(val):
    if pd.isna(val) or val == '': return 0
    parts = str(val).split(':')
    try:
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        elif len(parts) == 2:
            return int(parts[0]) * 60 + int(parts[1])
    except:
        return 0
    return 0

def generate():
    print("🚀 Starting Professional Dashboard Generation...")
    
    # 1. Load Data
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]
    if not csv_files:
        print("❌ No CSV files found in the directory.")
        return

    frames = []
    for f in csv_files:
        try:
            print(f"📖 Reading {f}...")
            # Try different encodings
            df = None
            for enc in ['utf-8', 'latin1', 'cp1255', 'iso-8859-8']:
                try:
                    df = pd.read_csv(f, encoding=enc)
                    if 'Start Time' in df.columns:
                        break
                except:
                    continue
            
            if df is not None:
                frames.append(df)
            else:
                print(f"⚠️ Could not read {f} with supported encodings.")
        except Exception as e:
            print(f"⚠️ Error reading {f}: {e}")

    if not frames:
        print("❌ No data could be read from CSV files.")
        return

    full_df = pd.concat(frames, ignore_index=True)
    
    # Clean Start Time
    full_df['Start Time'] = pd.to_datetime(full_df['Start Time'], dayfirst=True, errors='coerce')
    full_df = full_df.dropna(subset=['Start Time'])
    
    # Calculate durations
    full_df['sec'] = full_df['Interaction Duration'].apply(dur_to_sec)
    
    # ANI/DNIS Cleaning
    full_df['ani_clean'] = full_df['Dialed From (ANI)'].astype(str).str.replace('+', '', regex=False).str.replace('.0', '', regex=False)
    full_df['dnis_clean'] = full_df['Dialed To (DNIS)'].astype(str).str.replace('+', '', regex=False).str.replace('.0', '', regex=False)

    # 2. Main Metrics
    total_calls = len(full_df)
    avg_dur = round(full_df['sec'].mean() / 60, 2)
    
    vico_count = len(full_df[full_df['ani_clean'] == VICO_ANI])
    tier1_count = len(full_df[full_df['dnis_clean'] == TIER1_DNIS])
    vert_count = len(full_df[full_df['dnis_clean'] == VERTICALS_DNIS])
    shuf_count = len(full_df[full_df['dnis_clean'] == SHUFERSAL_DNIS])

    # 3. Monthly Volume
    monthly_series = full_df.groupby(full_df['Start Time'].dt.month).size()
    monthly_data = [int(monthly_series.get(i, 0)) for i in range(1, 13)]

    # 4. FCR Calculation (Monthly Weighted)
    fcr_df = full_df[full_df['dnis_clean'] != SHUFERSAL_DNIS].copy()
    fcr_df['month'] = fcr_df['Start Time'].dt.month
    
    monthly_scores = []
    total_weight = 0
    for m in range(1, 13):
        m_data = fcr_df[fcr_df['month'] == m]
        if len(m_data) == 0: continue
        
        caller_counts = m_data['ani_clean'].value_counts()
        unique_callers = len(caller_counts)
        ones = len(caller_counts[caller_counts == 1])
        score = ones / unique_callers
        
        monthly_scores.append({'score': score, 'weight': unique_callers})
        total_weight += unique_callers

    fcr_final = (sum(s['score'] * s['weight'] for s in monthly_scores) / total_weight * 100) if total_weight > 0 else 0
    fcr_final = round(fcr_final, 1)

    # 5. Agent Stats
    emp_stats = full_df.groupby('Employee').agg({'ani_clean': 'count', 'sec': 'mean'}).reset_index()
    emp_stats = emp_stats.sort_values('ani_clean', ascending=False)
    
    # 6. Build JS Data
    js_stats = {
        "total": total_calls,
        "avg": avg_dur,
        "vico": vico_count,
        "tier1": tier1_count,
        "vert": vert_count,
        "shuf": shuf_count,
        "monthly": monthly_data,
        "fcr": fcr_final,
        "emps": emp_stats['Employee'].head(6).tolist(),
        "empCounts": emp_stats['ani_clean'].head(6).tolist(),
        "empAvgs": [round(s/60, 2) for s in emp_stats['sec'].head(6).tolist()]
    }

    # 7. Generate Premium HTML Template
    html_content = f"""<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>מערכת BI - צוות Tier 2 Vico</title>
    <link href="https://fonts.googleapis.com/css2?family=Assistant:wght@200;400;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0"></script>
    <style>
        :root {{
            --bg: #07080c;
            --card: #11141b;
            --primary: #5865f2;
            --secondary: #eb459e;
            --accent: #57f287;
            --text-dim: #949ba4;
            --border: rgba(255, 255, 255, 0.08);
            --vico: #5865f2; --tier1: #eb459e; --verticals: #57f287; --shufersal: #fee75c;
        }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; font-family: 'Assistant', sans-serif; }}
        body {{ background: var(--bg); color: #fff; padding: 40px; min-height: 100vh; overflow-x: hidden; }}
        .dashboard {{ max-width: 1600px; margin: 0 auto; }}

        header {{ margin-bottom: 40px; border-right: 12px solid var(--primary); padding-right: 25px; }}
        header h1 {{ font-size: 3.5rem; font-weight: 800; letter-spacing: -1.5px; }}

        .kpi-row {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 30px; }}
        .kpi-card {{ background: var(--card); border-radius: 30px; padding: 30px 40px; border: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; box-shadow: 0 10px 40px rgba(0,0,0,0.3); }}
        .kpi-val {{ font-size: 3.5rem; font-weight: 800; }}
        .kpi-val span {{ font-size: 1.2rem; color: var(--primary); margin-right: 10px; }}
        .kpi-label {{ color: var(--text-dim); font-size: 1.2rem; font-weight: 600; }}

        .service-row {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 40px; }}
        .srv-card {{ background: var(--card); padding: 25px; border-radius: 22px; border: 1px solid var(--border); text-align: center; }}
        .srv-val {{ font-size: 2.2rem; font-weight: 800; margin-top: 5px; }}

        .grid {{ display: grid; grid-template-columns: 1.2fr 0.8fr; gap: 30px; }}
        .section {{ background: var(--card); border-radius: 35px; padding: 40px; border: 1px solid var(--border); }}
        .header-s {{ font-size: 1.8rem; font-weight: 800; margin-bottom: 35px; display: flex; align-items: center; gap: 15px; }}
        .header-s::before {{ content: ''; width: 8px; height: 28px; background: var(--primary); border-radius: 4px; }}

        .insight-card {{ background: rgba(255,255,255,0.02); border-radius: 24px; padding: 25px; margin-bottom: 20px; border-left: 6px solid var(--primary); display: flex; align-items: center; gap: 20px; }}
        .ins-icon {{ font-size: 2.2rem; background: rgba(255,255,255,0.03); width: 65px; height: 65px; display: flex; align-items: center; justify-content: center; border-radius: 18px; }}
        .ins-content h4 {{ color: var(--text-dim); font-size: 0.95rem; text-transform: uppercase; letter-spacing: 1px; }}
        .ins-content p {{ font-size: 1.6rem; font-weight: 800; }}

        .goals {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; margin-top: 40px; }}
        .goal-c {{ background: rgba(255,255,255,0.02); padding: 35px; border-radius: 30px; text-align: center; border: 1px solid var(--border); transition: 0.3s; }}
        .goal-c:hover {{ transform: translateY(-5px); border-color: var(--primary); background: rgba(88,101,242,0.05); }}
        .goal-icon {{ font-size: 2.8rem; display: block; margin-bottom: 15px; }}

        .success-box {{ margin-top: 60px; padding: 80px; text-align: center; border-radius: 45px; background: linear-gradient(135deg, rgba(88, 101, 242, 0.1), rgba(87, 242, 135, 0.1)); border: 1px solid var(--border); position: relative; }}
        .success-box h2 {{ font-size: 4.8rem; font-weight: 900; background: linear-gradient(to right, #fff, #5865f2, #57f287, #fff); background-size: 200% auto; -webkit-background-clip: text; -webkit-text-fill-color: transparent; animation: shine 5s linear infinite; }}
        @keyframes shine {{ to {{ background-position: 200% center; }} }}

        .agent-list {{ display: flex; flex-direction: column; gap: 12px; }}
        .agent-card {{ display: grid; grid-template-columns: 1.5fr 1fr 1fr; background: rgba(255,255,255,0.015); padding: 20px 30px; border-radius: 20px; border: 1px solid var(--border); align-items: center; }}
        canvas {{ width: 100% !important; height: 400px !important; }}
    </style>
</head>
<body>
    <div class="dashboard">
        <header>
            <h1>ביצועי שנת 2025 צוות Tier 2</h1>
            <p style="color:var(--text-dim); font-size:1.3rem; margin-top:5px;">מערכת BI מתקדמת • Strategic Analytics</p>
        </header>

        <div class="kpi-row">
            <div class="kpi-card" style="border-color: rgba(87,242,135,0.3);">
                <div><p class="kpi-label">סה"כ שיחות שנתי</p><div class="kpi-val">{js_stats['total']:,}<span>שיחות</span></div></div>
                <div style="font-size:3.5rem;">📞</div>
            </div>
            <div class="kpi-card" style="border-color: rgba(88,101,242,0.3);">
                <div><p class="kpi-label">ממוצע שיחה צוותי</p><div class="kpi-val">{js_stats['avg']}<span>דקות</span></div></div>
                <div style="font-size:3.5rem;">⏳</div>
            </div>
        </div>

        <div class="service-row">
            <div class="srv-card">
                <p class="kpi-label">(יוצא) Vico</p><div class="srv-val" style="color:var(--vico);">{js_stats['vico']:,}</div>
            </div>
            <div class="srv-card">
                <p class="kpi-label">(משכנו) Tier 1</p><div class="srv-val" style="color:var(--tier1);">{js_stats['tier1']:,}</div>
            </div>
            <div class="srv-card">
                <p class="kpi-label">Verticals</p><div class="srv-val" style="color:var(--verticals);">{js_stats['vert']:,}</div>
            </div>
            <div class="srv-card">
                <p class="kpi-label">Shufersal</p><div class="srv-val" style="color:var(--shufersal);">{js_stats['shuf']:,}</div>
            </div>
        </div>

        <div class="grid">
            <div class="section" style="grid-column: span 2;">
                <div class="header-s">נפח פעילות חודשי</div>
                <canvas id="mChart"></canvas>
            </div>

            <div class="section">
                <div class="header-s">פילוח שירותים (Market Share)</div>
                <div style="height:420px;"><canvas id="pChart"></canvas></div>
            </div>

            <div class="section">
                <div class="header-s">ניתוח מגמות ותובנות</div>
                <div class="insight-card">
                    <div class="ins-icon">💎</div>
                    <div class="ins-content"><h4>פתרון בפנייה ראשונה</h4><p>{js_stats['fcr']}% מדד FCR חודשי</p></div>
                </div>
                <div class="insight-card">
                    <div class="ins-icon">📈</div>
                    <div class="ins-content"><h4>שיא פעילות חודשי</h4><p id="peakMonth">-</p></div>
                </div>
                <div class="insight-card">
                    <div class="ins-icon">🎯</div>
                    <div class="ins-content"><h4>דומיננטיות שירות</h4><p>Verticals מהווה {round(js_stats['vert']/js_stats['total']*100)}%</p></div>
                </div>
            </div>

            <div class="section">
                <div class="header-s">ביצועי נציגים (כמות שיחות)</div>
                <canvas id="eChart"></canvas>
            </div>

            <div class="section">
                <div class="header-s">דירוג איכות נציגים</div>
                <div id="aList" class="agent-list"></div>
            </div>
        </div>

        <div class="section" style="margin-top:40px; border:none; background:transparent; padding:0;">
            <div class="header-s" style="justify-content:center;">חזון ויעדים לשנת 2026</div>
            <div class="goals">
                <div class="goal-c"><span class="goal-icon">🚀</span><h3>חדשנות</h3><p>כלי אבחון AI לקיצור זמנים</p></div>
                <div class="goal-c"><span class="goal-icon">🎯</span><h3>מצוינות</h3><p>שמירה על FCR מעל 75%</p></div>
                <div class="goal-c"><span class="goal-icon">💎</span><h3>מקצועיות</h3><p>שיפור מתמיד של הידע הטכני</p></div>
                <div class="goal-c"><span class="goal-icon">🤝</span><h3>סינרגיה</h3><p>עבודת צוות ותמיכה הדדית</p></div>
            </div>
        </div>

        <div class="success-box">
            <div style="margin-bottom: 25px;">
                <svg width="120" height="120" viewBox="0 0 24 24" fill="none">
                    <path d="M6 9V2H18V9M6 9C6 11.2091 7.79086 13 10 13H14C16.2091 13 18 11.2091 18 9M6 9H4C2.89543 9 2 8.10457 2 7V5C2 3.89543 2.89543 3 4 3H6M18 9H20C21.1046 9 22 8.10457 22 7V5C22 3.89543 21.1046 3 20 3H18M12 13V17M12 17H8M12 17H16M7 22H17M10 22L12 17L14 22" stroke="url(#g)" stroke-width="1.2"/>
                    <defs><linearGradient id="g" x1="2" y1="2" x2="22" y2="22"><stop stop-color="#fee75c"/><stop offset="1" stop-color="#f39c12"/></linearGradient></defs>
                </svg>
            </div>
            <h2>בהצלחה ב-2026!</h2>
            <p style="font-size:1.8rem; color:var(--text-dim);">ממשיכים קדימה תמיד • צוות Tier 2 Vico</p>
        </div>

        <footer style="margin-top:40px; text-align:center; padding:40px; color:var(--text-dim);">VERIFONE ISRAEL • BI SYSTEM • 2026</footer>
    </div>

    <script>
        Chart.register(ChartDataLabels);
        const data = {json.dumps(js_stats)};
        
        const boxL = {{ backgroundColor: '#000', color: '#fff', borderRadius: 5, padding: 8, font: {{ weight: '800' }}, display: true, anchor: 'end', align: 'top', offset: 10 }};

        new Chart(document.getElementById('mChart'), {{
            type: 'line',
            data: {{ labels: ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'], datasets: [{{ data: data.monthly, borderColor: '#5865f2', borderWidth: 5, pointRadius: 6, tension: 0.4, fill: true, backgroundColor: 'rgba(88,101,242,0.1)' }}] }},
            options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }}, datalabels: boxL }}, scales: {{ y: {{ display: false }}, x: {{ grid: {{ display: false }}, ticks: {{ color: '#949ba4', font: {{ weight: '700' }} }} }} }} }}
        }});

        new Chart(document.getElementById('pChart'), {{
            type: 'doughnut',
            data: {{ labels: ['Verticals', 'Vico', 'Tier 1', 'Shufersal'], datasets: [{{ data: [data.vert, data.vico, data.tier1, data.shuf], backgroundColor: ['#57f287','#5865f2','#eb459e','#fee75c'], borderWidth: 0 }}] }},
            options: {{ responsive: true, maintainAspectRatio: false, cutout: '72%', plugins: {{ datalabels: {{ color: '#fff', font: {{ weight: '900', size: 15 }}, formatter: (v, ctx) => {{ const s = ctx.chart.data.datasets[0].data.reduce((a,b)=>a+b); return Math.round(v/s*100) + '%'; }} }} }} }}
        }});

        new Chart(document.getElementById('eChart'), {{
            type: 'bar',
            data: {{ labels: data.emps, datasets: [{{ data: data.empCounts, backgroundColor: '#5865f2', borderRadius: 12 }}] }},
            options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }}, datalabels: boxL }}, scales: {{ y: {{ display: false }}, x: {{ grid: {{ display: false }}, ticks: {{ color: '#949ba4', font: {{ weight: '700' }} }} }} }} }}
        }});

        const list = document.getElementById('aList');
        data.emps.forEach((n, i) => {{
            const div = document.createElement('div');
            div.className = 'agent-card';
            div.innerHTML = `<div><b style="font-size:1.2rem;">${{n}}</b></div><div style="text-align:center">שיחות: <b>${{data.empCounts[i].toLocaleString()}}</b></div><div style="text-align:center">ממוצע: <b style="color:var(--accent)">${{data.empAvgs[i]}}</b></div>`;
            list.appendChild(div);
        }});

        const mNames = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר'];
        const maxIdx = data.monthly.indexOf(Math.max(...data.monthly));
        document.getElementById('peakMonth').innerText = mNames[maxIdx] + " (" + Math.max(...data.monthly).toLocaleString() + " שיחות)";
    </script>
</body>
</html>"""

    with open('presentation.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("✅ High-End Dashboard generated: presentation.html")
    webbrowser.open('file://' + os.path.realpath('presentation.html'))

if __name__ == "__main__":
    generate()
