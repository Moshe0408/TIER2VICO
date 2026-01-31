# -*- coding: utf-8 -*-
import os

path = r'c:\Users\moshei1\OneDrive - Verifone\Desktop\TIP\STFPNOW\בדיקות\Dashboard_App.py'

# 1. Read the file robustly
with open(path, 'r', encoding='utf-8', errors='ignore') as f:
    lines = f.readlines()

# 2. Extract strictly good parts
header_part = lines[:139] # Lines 1 to 139 (indices 0 to 138)
body_part = lines[910:]   # Line 911 to end (index 910 to end)

# 3. Define the clean middle section
middle_part = """
CUSTOMER_LOGOS = {
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
}


class handler(http.server.SimpleHTTPRequestHandler):
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

    def get_login_ui(self):
        return r\"\"\"
<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>Vico | Login</title>
    <style>
        :root { --primary: #3b82f6; --accent: #10b981; --bg: #030712; --glass: rgba(17, 24, 39, 0.7); --border: rgba(255, 255, 255, 0.1); --accent-glow: rgba(16, 185, 129, 0.2); }
        * { box-sizing: border-box; }
        body { margin: 0; padding: 0; background-color: var(--bg); color: #fff; font-family: system-ui, -apple-system, sans-serif; display: flex; align-items: center; justify-content: center; min-height: 100vh; overflow: hidden; direction: rtl; }
        .card { background: var(--glass); backdrop-filter: blur(40px); padding: 70px 50px; border-radius: 50px; border: 1px solid var(--border); text-align: center; width: 440px; }
        .logo { height: 30px; margin-bottom: 40px; filter: brightness(0) invert(1); }
        h1 { font-size: 38px; margin: 0; background: linear-gradient(to bottom, #fff, #94a3b8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }
        p { color: #8b949e; margin: 15px 0 40px; }
        .input-box { text-align: right; margin-bottom: 25px; }
        label { display: block; font-size: 12px; font-weight: 800; color: var(--accent); margin-bottom: 10px; }
        input { width: 100%; background: rgba(255,255,255,0.05); border: 1px solid var(--border); padding: 20px; border-radius: 20px; color: #fff; font-size: 16px; outline: none; }
        input:focus { border-color: var(--accent); box-shadow: 0 0 20px var(--accent-glow); }
        .action-btn { width: 100%; background: linear-gradient(135deg, #3b82f6, #1d4ed8); color: #fff; padding: 20px; border-radius: 20px; font-size: 18px; font-weight: 800; cursor: pointer; border: none; margin-top: 10px; transition: 0.3s; }
        .action-btn:hover { transform: translateY(-3px); box-shadow: 0 10px 30px rgba(59,130,246,0.4); }
        .error-notif { display: none; background: rgba(239,68,68,0.1); color: #f87171; padding: 15px; border-radius: 15px; margin-top: 20px; font-weight: 700; }
        .legal { margin-top: 40px; font-size: 12px; color: #484f58; border-top: 1px solid var(--border); padding-top: 20px; }
    </style>
</head>
<body>
    <div class="card">
        <img src=\"https://upload.wikimedia.org/wikipedia/commons/9/98/Verifone_Logo.svg\" class=\"logo\" alt=\"Verifone\">
        <h1>מרכז הבקרה Vico</h1>
        <p>התחברות לאזור המורשה של Tier 2</p>
        <div class=\"input-box\"><label>זיהוי משתמש (Email)</label><input type=\"email\" id=\"u-mail\" placeholder=\"name@verifone.com\"></div>
        <div class=\"input-box\"><label>סיסמת גישה</label><input type=\"password\" id=\"u-pass\" placeholder=\"••••••••\"></div>
        <button class=\"action-btn\" id=\"l-btn\" onclick=\"handleAuth()\">כניסה למערכת</button>
        <div id=\"msg\" class=\"error-notif\">פרטי המשתמש אינם תואמים.</div>
        <div class=\"legal\">Verifone &copy; 2026. כל הזכויות שמורות.</div>
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
                else { throw new Error(\"Login failed\"); }
            } catch (error) { msg.style.display = 'block'; btn.disabled = false; btn.innerText = \"כניסה למערכת\"; }
        };
    </script>
</body>
</html>
\"\"\"
"""

# 4. Construct final content
final_content = "".join(header_part) + middle_part + "".join(body_part)

# 5. Clean up duplicated charset info in the body part
final_content = final_content.replace('charset=utf-8; charset=utf-8', 'charset=utf-8')

# 6. Save the file
with open(path, 'w', encoding='utf-8') as f:
    f.write(final_content)

print("Reconstruction complete.")
