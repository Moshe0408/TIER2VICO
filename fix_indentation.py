# -*- coding: utf-8 -*-
import os
import re

path = r'c:\Users\moshei1\OneDrive - Verifone\Desktop\TIP\STFPNOW\בדיקות\Dashboard_App.py'

# Read the file
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    content = f.read()

# 1. Fix the messy CUSTOMER_LOGOS duplication
# Find the start of the first CUSTOMER_LOGOS and the end of the last one before handler class
customer_logos_clean = """CUSTOMER_LOGOS = {
    "shufersal": {
        "name": "שופרסל", 
        "logo": "https://upload.wikimedia.org/wikipedia/he/f/f0/ShufersalLogo.svg", 
        "desc": "רשת הקמעונאות הגדולה והמובילה בישראל, המפעילה מאות סניפים תחת מותגים שונים ומהווה עוגן משמעותי בשוק הצריכה המקומי.",
        "fallbacks": ["https://logo.clearbit.com/shufersal.co.il", "https://www.shufersal.co.il/online/static/media/logo.dfdfdfdf.png"]
    },
    "ikea": {
        "name": "איקאה", 
        "logo": "https://diversityisrael.org.il/wp-content/uploads/%D7%9C%D7%95%D7%92%D7%95-%D7%90%D7%99%D7%A7%D7%90%D7%941.png", 
        "desc": "תאגיד רהיטים בינלאומי המציע מגוון רחב של פתרונות לעיצוב הבית. הרשת ידועה בחוויית הקניה הייחודית שלה ובפריסת מרכזי ענק.",
        "fallbacks": ["https://logo.clearbit.com/ikea.co.il", "https://www.ikea.co.il/images/logo.png"]
    },
    "mcdonald": {
        "name": "מקדונלד'ס", 
        "logo": "https://upload.wikimedia.org/wikipedia/commons/3/36/McDonald%27s_Golden_Arches.svg", 
        "desc": "רשת המזון המהיר הגדולה והמוכרת בעולם. בישראל הרשת מובילה את התחום עם פריסה ארצית רחבה וחדשנות בשירות הדיגיטלי.",
        "fallbacks": ["https://logo.clearbit.com/mcdonalds.co.il", "https://www.mcdonalds.co.il/assets/images/logo.png"]
    },
    "aroma": {
        "name": "ארומה", 
        "logo": "https://upload.wikimedia.org/wikipedia/he/c/c9/Aroma_Espresso_Bar_Logo.svg", 
        "desc": "רשת בתי קפה ישראלית מובילה, המפורסמת בזכות האספרסו האיכותי, הכריכים הטריים והמאפים המיוחדים שלה.",
        "fallbacks": ["https://logo.clearbit.com/aroma.co.il", "https://online.aroma.co.il/static/media/logo.png"]
    }
}"""

content = re.sub(r'CUSTOMER_LOGOS = \{.*?\}\n\},.*?\n\}', customer_logos_clean, content, flags=re.DOTALL)
# One more pass in case it's still there
content = re.sub(r'CUSTOMER_LOGOS = \{.*?\}\n}(,)?', customer_logos_clean, content, flags=re.DOTALL)


# 2. Fix the handler class and is_authenticated (indentation fix)
handler_fix = """class handler(http.server.SimpleHTTPRequestHandler):
    def is_authenticated(self):
        try:
            cookie_header = self.headers.get('Cookie')
            if not cookie_header: return False
            import http.cookies, base64
            C = http.cookies.SimpleCookie(cookie_header)
            sid = C.get('sid')
            if sid:
                try:
                    # Stateless: sid is base64(email:expiry_timestamp)
                    val = base64.b64decode(sid.value.encode()).decode()
                    email, expiry = val.split(':')
                    if float(expiry) > datetime.now().timestamp():
                        return True
                except: pass
        except Exception: pass
        return False
"""

content = re.sub(r'class handler\(http\.server\.SimpleHTTPRequestHandler\):.*?def is_authenticated\(self\):.*?return False\n', handler_fix, content, flags=re.DOTALL)

# 3. Restore Hebrew in UI methods
# Login UI
login_ui_fixed = """    def get_login_ui(self):
        return r\"\"\"
<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>Vico | Login</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Internal+Sans:wght@400;700&display=swap');
        :root { --primary: #3b82f6; --accent: #10b981; --bg: #030712; --glass: rgba(17, 24, 39, 0.7); --border: rgba(255, 255, 255, 0.1); --accent-glow: rgba(16, 185, 129, 0.2); }
        * { box-sizing: border-box; }
        body { margin: 0; padding: 0; background-color: var(--bg); color: #fff; font-family: 'Inter', system-ui, -apple-system, sans-serif; display: flex; align-items: center; justify-content: center; min-height: 100vh; overflow: hidden; }
        .scene { position: fixed; top: 0; left: 0; width: 100%; height: 100%; z-index: -1; }
        .orb { position: absolute; border-radius: 50%; filter: blur(80px); opacity: 0.4; animation: float 20s infinite alternate; }
        .orb-1 { width: 500px; height: 500px; background: #1e3a8a; top: -100px; right: -100px; }
        .orb-2 { width: 400px; height: 400px; background: #1e1b4b; bottom: -50px; left: -50px; animation-delay: -5s; }
        .orb-3 { width: 300px; height: 300px; background: #312e81; top: 40%; left: 30%; animation-delay: -10s; }
        @keyframes float { from { transform: translate(0,0) rotate(0deg); } to { transform: translate(40px, 40px) rotate(10deg); } }
        .card-container { z-index: 10; width: 100%; max-width: 440px; animation: slideUp 1s cubic-bezier(0.2, 0.8, 0.2, 1); }
        @keyframes slideUp { from { opacity: 0; transform: translateY(30px); } to { opacity: 1; transform: translateY(0); } }
        .card { background: var(--glass); backdrop-filter: blur(40px); -webkit-backdrop-filter: blur(40px); padding: 70px 50px; border-radius: 50px; border: 1px solid rgba(255, 255, 255, 0.08); box-shadow: 0 30px 60px rgba(0,0,0,0.8), inset 0 0 0 1px rgba(255,255,255,0.05); text-align: center; }
        .logo-wrap { margin-bottom: 40px; }
        .logo { height: 30px; filter: drop-shadow(0 0 15px rgba(255,255,255,0.4)); transition: 0.5s; }
        .logo:hover { transform: scale(1.05); }
        .title-wrap { margin-bottom: 45px; }
        .title-wrap h1 { font-size: 38px; font-weight: 800; margin: 0; background: linear-gradient(to bottom, #fff, #94a3b8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; letter-spacing: -1px; }
        .title-wrap p { color: #8b949e; font-size: 16px; margin: 12px 0 0; font-weight: 400; }
        .form-grid { display: flex; flex-direction: column; gap: 30px; }
        .input-box { text-align: right; }
        .input-box label { display: block; font-size: 12px; font-weight: 800; color: var(--accent); margin-bottom: 12px; margin-right: 5px; text-transform: uppercase; letter-spacing: 1px; }
        .field-wrap { position: relative; }
        input { width: 100%; background: rgba(255, 255, 255, 0.03); border: 1px solid rgba(255,255,255,0.1); padding: 20px 25px; border-radius: 24px; color: #fff; font-size: 17px; font-weight: 500; outline: none; transition: 0.4s cubic-bezier(0.4, 0, 0.2, 1); box-sizing: border-box; text-align: left; direction: ltr; }
        input:focus { background: rgba(255, 255, 255, 0.06); border-color: var(--accent); box-shadow: 0 0 30px var(--accent-glow); transform: translateY(-2px); }
        .action-btn { margin-top: 15px; width: 100%; background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%); color: #fff; padding: 22px; border-radius: 24px; font-size: 20px; font-weight: 800; cursor: pointer; border: none; transition: 0.4s; box-shadow: 0 15px 35px -10px rgba(59, 130, 246, 0.5); }
        .error-notif { display: none; background: rgba(239, 68, 68, 0.1); border: 1px solid rgba(239, 68, 68, 0.2); color: #f87171; padding: 18px; border-radius: 20px; margin-top: 30px; font-weight: 700; font-size: 14px; }
        .legal { margin-top: 45px; font-size: 13px; color: #484f58; font-weight: 600; border-top: 1px solid rgba(255,255,255,0.05); padding-top: 30px; }
    </style>
</head>
<body>
    <div class="scene"><div class="orb orb-1"></div><div class="orb orb-2"></div><div class="orb orb-3"></div></div>
    <div class="card-container">
        <div class="card">
            <div class="logo-wrap"><img src=\"https://upload.wikimedia.org/wikipedia/commons/9/98/Verifone_Logo.svg\" class=\"logo\" alt=\"Verifone\" style=\"filter: brightness(0) invert(1);\"></div>
            <div class=\"title-wrap\"><h1>מרכז הבקרה Vico</h1><p>התחברות לאזור המורשה של Tier 2</p></div>
            <div class=\"form-grid\">
                <div class=\"input-box\"><label>זיהוי משתמש (Email)</label><div class=\"field-wrap\"><input type=\"email\" id=\"u-mail\" placeholder=\"name@verifone.com\" required></div></div>
                <div class=\"input-box\"><label>סיסמת גישה</label><div class=\"field-wrap\"><input type=\"password\" id=\"u-pass\" placeholder=\"••••••••\" required></div></div>
                <button class=\"action-btn\" id=\"l-btn\" onclick=\"handleAuth()\">כניסה למערכת</button>
            </div>
            <div id=\"msg\" class=\"error-notif\">שגיאת אימות: פרטי המשתמש אינם תואמים.</div>
            <div class=\"legal\">מערכת פנימית של Verifone &copy; 2026. כל הזכויות שמורות.</div>
        </div>
    </div>
    <script type=\"module\">
        import { initializeApp } from \"https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js\";
        import { getAuth, signInWithEmailAndPassword } from \"https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js\";
        const config = { apiKey: \"AIzaSyB3pruogaljwaw9FVyrD3MvPOHgpyGfxzs\", authDomain: \"tier-2-vico.firebaseapp.com\", projectId: \"tier-2-vico\", storageBucket: \"tier-2-vico.firebasestorage.app\", messagingSenderId: \"272065575004\", appId: \"1:272065575004:web:11ed615295a56dbc824e99\" };
        const app = initializeApp(config);
        const auth = getAuth(app);
        window.handleAuth = async () => {
            const email = document.getElementById('u-mail').value;
            const pass = document.getElementById('u-pass').value;
            const btn = document.getElementById('l-btn');
            const msg = document.getElementById('msg');
            if(!email || !pass) return;
            btn.disabled = true; btn.innerText = \"מעבד...\"; msg.style.display = 'none';
            try {
                const userCredential = await signInWithEmailAndPassword(auth, email, pass);
                const idToken = await userCredential.user.getIdToken();
                const response = await fetch('/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ idToken, email }) });
                if (response.ok) { const data = await response.json(); window.location.href = data.url; }
                else { throw new Error(\"Validation failure\"); }
            } catch (error) { console.error(error); msg.style.display = 'block'; btn.disabled = false; btn.innerText = \"כניסה למערכת\"; }
        };
    </script>
</body>
</html>
\"\"\"
"""

content = re.sub(r'def get_login_ui\(self\):.*?\"\"\"\n', login_ui_fixed, content, flags=re.DOTALL)

# Main Dashboard UI (fix Hebrew title)
content = content.replace("<title>Tier 2 Vico | Intelligence Dashboard</title>", "<title>מערכת Vico | דאשבורד ניהולי</title>")

# Fix common garbled patterns
content = content.replace("׳–׳™׳”׳•׳™ ׳ž׳©׳×׳ž׳©", "זיהוי משתמש")
content = content.replace("׳¡׳™׳¡׳ž׳× ׳’׳™׳©׳”", "סיסמת גישה")
content = content.replace("׳›׳ ׳™׳¡׳” ׳œ׳ž׳¢׳¨׳›׳×", "כניסה למערכת")

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Fix completed.")
