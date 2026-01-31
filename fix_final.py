# -*- coding: utf-8 -*-
import os
import re
import base64

path = r'c:\Users\moshei1\OneDrive - Verifone\Desktop\TIP\STFPNOW\בדיקות\Dashboard_App.py'

# Read with a robust way to handle the current mess
try:
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
except:
    with open(path, 'r', encoding='windows-1255', errors='ignore') as f:
        content = f.read()

# 1. Fix Headers/Encoding Declaration (remove BOM and duplicates)
content = re.sub(r'^.*?import http\.server', 'import http.server', content, flags=re.DOTALL)
content = '# -*- coding: utf-8 -*-\n' + content

# 2. Fix TIER2_MAP
tier2_map_clean = """TIER2_MAP = {
    "niv.arieli": "ניב אריאלי", "din.weissman": "דין וייסמן", "lior.burstein": "ליאור בורשטיין", "liorb5": "ליאור בורשטיין",
    "avivs": "אביב סולר", "ebrahimf": "אברהים פריג", "orenw1": "אורן וייס", "ahmado": "אחמד עודה",
    "almancha": "אלמנך עלמיה", "zahiyas1": "זהייה אבו שמאלה", "tals": "טל שוקר", "yuvala1": "יובל אגרון",
    "yuliano": "יוליאן אולרסקו", "yoadc": "יועד כחלון", "nuphars": "נוּפר שלום", "idoh": "עידו הרמל",
    "aviele": "אביאל אלשוילי", "avivk": "אביב כץ", "bari": "בר ישראלי", "veral2": "ורה ליברמן",
    "danv1": "דן וייסמן", "niva2": "ניב אריאלי", "nadavl1": "נדב", "paulp": "פאול",
    "moshei1": "משה איטח", "nadav.lieber": "נדב", "erezm1": "ארז", "almanch.alme": "אלמנך עלמיה",
    "dan.vico": "דן ויקו", "danv": "דן ויקו", "shira": "שיר אברהם"
}"""
content = re.sub(r'TIER2_MAP = \{.*?\}', tier2_map_clean, content, flags=re.DOTALL)

# 3. Fix CUSTOMER_LOGOS
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
content = re.sub(r'CUSTOMER_LOGOS = \{.*?\}', customer_logos_clean, content, flags=re.DOTALL)

# 4. Fix UI Strings (Login Page)
content = content.replace('<h1>׳ž׳¨׳›׳– ׳”׳‘׳§׳¨׳” Vico</h1>', '<h1>מרכז הבקרה Vico</h1>')
content = content.replace('׳”׳×׳—׳‘׳¨׳•׳× ׳œ׳ ׳–׳•׳¨ ׳”׳ž׳•׳¨׳©׳” ׳©׳œ Tier 2', 'התחברות לאזור המורשה של Tier 2')
content = content.replace('׳–׳™׳”׳•׳™ ׳ž׳©׳×׳ž׳© (Email)', 'זיהוי משתמש (Email)')
content = content.replace('׳¡׳™׳¡׳ž׳× ׳’׳™׳©׳”', 'סיסמת גישה')
content = content.replace('׳›׳ ׳™׳¡׳” ׳œ׳ž׳¢׳¨׳›׳×', 'כניסה למערכת')
content = content.replace('׳©׳’׳™׳ ׳× ׳ ׳™׳ž׳•׳×: ׳₪׳¨׳˜׳™ ׳”׳ž׳©׳×׳ž׳© ׳ ׳™׳ ׳  ׳×׳•׳ ׳ž׳™׳ .', 'שגיאת אימות: פרטי המשתמש אינם תואמים.')
content = content.replace('׳ž׳¢׳¨׳›׳× ׳₪׳ ׳™׳ž׳™׳× ׳©׳œ Verifone &copy; 2026. ׳›׳œ ׳”׳–׳כ׳•׳™׳•׳× ׳©׳ž׳•׳¨׳•׳×.', 'מערכת פנימית של Verifone &copy; 2026. כל הזכויות שמורות.')
content = content.replace('"׳ž׳¢׳‘׳“..."', '"מעבד..."')

# 5. Fix Dir Logic in do_GET
content = content.replace("'/׳ž׳“׳¨׳™׳›׳™׳ /'", "'/מדריכים/'")

# 6. Fix Auth - Move to Stateless (Simple Base64 Cookie for Vercel)
stateless_auth = """    def is_authenticated(self):
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
content = re.sub(r'def is_authenticated\(self\):.*?return False', stateless_auth, content, flags=re.DOTALL)

# 7. Fix Login POST to use Stateless Auth
# Firebase Login
content = content.replace('sid = str(uuid.uuid4())\n                        SESSIONS[sid] = {\n                            "user": email,\n                            "expiry": datetime.now() + timedelta(days=1)\n                        }', 
                          'import base64\n                        expiry = str((datetime.now() + timedelta(days=1)).timestamp())\n                        sid = base64.b64encode(f"{email}:{expiry}".encode()).decode()')

# Hardcoded Login
content = content.replace('sid = str(uuid.uuid4())\n                    SESSIONS[sid] = {"user": email, "expiry": datetime.now() + timedelta(days=1)}',
                          'import base64\n                    expiry = str((datetime.now() + timedelta(days=1)).timestamp())\n                    sid = base64.b64encode(f"{email}:{expiry}".encode()).decode()')

# 8. Ensure charset=utf-8 everywhere
content = content.replace("'Content-type', 'text/html'", "'Content-type', 'text/html; charset=utf-8'")
content = content.replace("'Content-Type', 'text/html'", "'Content-Type', 'text/html; charset=utf-8'")

with open(path, 'w', encoding='utf-8') as f:
    f.write(content)

print("Successfully fixed Hebrew encoding and and applied Stateless Auth in Dashboard_App.py")
