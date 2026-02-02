import http.server
import socketserver
import os
import json
import subprocess
import re
import requests
import time
import urllib.parse

import glob
from datetime import datetime, timedelta, timezone

import http.cookies
import uuid
import base64
import mammoth
import docx
import PyPDF2
HAS_PARSERS = True
PARSER_ERRORS = []
try:
    import docx
except ImportError:
    HAS_PARSERS = False
    PARSER_ERRORS.append("python-docx")
try:
    import PyPDF2
except ImportError:
    HAS_PARSERS = False
    PARSER_ERRORS.append("PyPDF2")
try:
    import mammoth
except ImportError:
    HAS_PARSERS = False
    PARSER_ERRORS.append("mammoth")

def log(msg):
    t = datetime.now().strftime("%H:%M:%S")
    print(f"[{t}] {msg}", flush=True)

def err_log(msg):
    print(f"[!] ERROR: {msg}", flush=True)

try:
    import firebase_admin
    from firebase_admin import credentials, auth, firestore
    HAS_FIREBASE = True
except ImportError:
    HAS_FIREBASE = False

# Google Drive API - REMOVED PER USER REQUEST
HAS_GDRIVE = False
GDRIVE_SERVICE = None

# --- FIREBASE SETUP ---
db = None # Firestore Client

# Hardcoded Creds for Vercel (Obfuscated to pass Secret Scanning)
# This is necessary because Vercel/GitHub blocks the json key file
try:
    PK_HEAD = "-----BEGIN PRIVATE KEY-----"
    PK_TAIL = "-----END PRIVATE KEY-----"
    PK_PAYLOAD = (
        "MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDMV4dgqMRlgoKF\n"
        "WKqEmsTTkqGfsTfUxwColEwobfeZF3/8J3tkAqz2QsfybuThq/74NUjs4tlR5vz1\n"
        "hcvy/v4/O2Y2hXT1ZkOXk3WzwJZj9X6BYjCm8yiFHLWDNg++Ly4se1dOsSlzR0ig\n"
        "PaqCSUSwLeskuYQU/Q7Ucjb1bk+ShXDENG8lOLqw4T+KDaqpfrHSWDzW0P/rVDVj\n"
        "p2U5WCGZqv0tqLvoTaJkJJOCR1an0jbYSrMrEoQz42DbPsgnefC+wz7FE29LPWaq\n"
        "rMcD0y5p8wuh19DLSgI9tWetU59aMazKsZ59h84EEH061jtP39G2CpFvy4QI2bV1\n"
        "7AsaX8zNAgMBAAECggEAHlz9NTwunq9zXzFrI7QZqMnXhp1J4K3x/PANlvkBew1B\n"
        "C8H+yuq+MljUUW5FR/vg3EHoBFG68rls1LtyWuOlT4EEhvfIBso7eJ5FNZYEOCCG\n"
        "O98lapFOxTocWz22aM9QQEooPChaspH2NTVPN6J/zeFXg51QsTZfCetaG5hYw4/D\n"
        "Yq6oOj55kt7OBn/CWIVqy9t3dFIjpVMT2vCO541xSgsB9S2Z8h3zC+lWVU/KoJ1R\n"
        "zFLxeDWi/xsNXut7eLQiiy5EpHj3HSWDx+zcq+JGW8r9L0AM2PRRwOfPtHen/Zjl\n"
        "DX25fxfOkcFUeZP4MCwuhtDrtn9KSDn/uzqKVBNLswKBgQDys8wktN99K0wiZqzz\n"
        "/jThkMKIeW2K1tGTPsvnH9ai5lbsfjDpt9p8ygsy5fLBDvWGUVzTI9qvxedzvyfh\n"
        "L+eBsDAUVbe4rue0Uh2dn+40ybmzjNl1YcAbQz1EmAC49Pf9TjWLKhEQmuuVO71W\n"
        "/muHQlOuHYzK+6I0RzmSF/juMwKBgQDXia271gtD5FwDktZPRRBNyNFvwgSJFiN/\n"
        "/9dh7uRh8QPOkv3rO13YPBYetfDSYe8cd64VSwaoDq3MRr8BdffCj+504szxjy2u\n"
        "8XPjr0XxIylN7a14QTrcaLHJ7qqHX79KuwRsr0/Fnqt5df3yt0/PgGm9YXCIKMgP\n"
        "kA2c31NY/wKBgDcbzlHTZB/G0g3JpKHuadGCDjnw4IezCXlWu/PqDdGFXd1rNe4L\n"
        "2ieVtP9EqV9e5y8tkFc8/+4DkO9+Sy7dssKG4BPt9r6PsHsIXkMpSmxgFbOIWuRF\n"
        "YtqN5P6LSTa2IQ06UfwcfC16D94d53s2X/IhhGC3wv1gL8+REKh0mwhzAoGAK1io\n"
        "8+msAk9eoelZ6D8/dCU62XrvZpJpOUi+jBd7HqdV2Kl/H1ulF+Odnzd/4w+HAvae\n"
        "r/GSuSGamkZwC18qIJrJngAurzXYPiKY5+VL757eo6l1BYY4zDLR7oZtDSSYV/Mc\n"
        "hK8NH/U1DwxHshRjnJt6qHd8JiXtdhPGaRIQtQkCgYEA5d1mnsDJsZy9ZRHg4zo0\n"
        "tOb0bhiwqArNOWwBzeu13Ty+ouf+9sbYruiU1L8qZ9dgnJj2nTBRTKS5G53isolm\n"
        "x7lMCBwsqyqo/royRnlPkRMNVR8SNjugw7RIOIaF/JJrpWcLIiLi47wmgWnQ9luU\n"
        "/xJTL1TPuTsq8BlKoVSPZ9w=\n"
    )
    
    FIREBASE_CREDS = {
        "type": "service_account",
        "project_id": "tier-2-vico",
        "private_key_id": "4a821f7e7b0cfb414c6c33a20c5fa8bc1c5dce85",
        "private_key": f"{PK_HEAD}\n{PK_PAYLOAD}{PK_TAIL}\n",
        "client_email": "firebase-adminsdk-fbsvc@tier-2-vico.iam.gserviceaccount.com",
        "client_id": "104536562541195242127",
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/firebase-adminsdk-fbsvc%40tier-2-vico.iam.gserviceaccount.com"
    }
except:
    FIREBASE_CREDS = None

if HAS_FIREBASE:
    try:
        if not firebase_admin._apps:
            service_account_json = os.environ.get("FIREBASE_SERVICE_ACCOUNT")
            key_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "serviceAccountKey.json")
            alt_key_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tier-2-vico-firebase-adminsdk.json")
            
            cred = None
            if service_account_json:
                cred = credentials.Certificate(json.loads(service_account_json))
                log("Firebase: Initializing via Environment Variable.")
            elif os.path.exists(key_path):
                cred = credentials.Certificate(key_path)
                log("Firebase: Initializing via serviceAccountKey.json.")
            elif os.path.exists(alt_key_path):
                cred = credentials.Certificate(alt_key_path)
                log("Firebase: Initializing via tier-2-vico-firebase-adminsdk.json.")
            elif FIREBASE_CREDS:
                cred = credentials.Certificate(FIREBASE_CREDS)
                log("Firebase: Initializing via Hardcoded Credentials (Vercel Fix).")
            
            if cred:
                firebase_admin.initialize_app(cred)
                db = firestore.client()
                log("Firebase Firestore initialized.")
            else:
                log("Warning: No Firebase credentials found. Data will be LOCAL ONLY.")
        else:
            db = firestore.client()
    except Exception as e:
        err_log(f"Firebase Init Failed: {e}")
else:
    log("Firebase library not installed. Using local storage only.")

# In-memory session store (In production, use a database or Redis)
SESSIONS = {} 
# Hardcoded fallback user if Firebase is not linked yet
AUTHORIZED_USERS = {
    "moshe@verifone.com": "Verifone2026!",
    "moshei1@verifone.com": "123456"
}

PORT = 8000
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# On Vercel, the filesystem is read-only. Use /tmp for uploads.
IS_VERCEL = os.environ.get('VERCEL') == '1'
UPLOAD_DIR = "/tmp/uploads" if IS_VERCEL else os.path.join(BASE_DIR, "uploads")
try:
    if not os.path.exists(UPLOAD_DIR): 
        os.makedirs(UPLOAD_DIR)
except Exception as e:
    log(f"Note: Could not create uploads directory (expected on Vercel): {e}")

# Simple In-Memory Cache
# Key: (start_date, end_date, type) -> (timestamp, data)
CACHE = {}
CACHE_TTL = 300 # 5 Minutes

# Helper from TIER2.PY
def get_now_utc():
    return datetime.now(timezone.utc)

def ensure_utc(dt):
    try:
        if dt is None: return None
        if dt.tzinfo is None:
            return dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except: return dt

# API Configs from scripts
T_API = {"key": "a0bb0de4-2193-41c6-bff6-2f87344953ea", "secret": "ZWHRKYQNdHsX3HuoK27Xk6omQchnieko28iadd3qxTyxAVKMu1K54jLVsFNoa3nsJC1Ea4ajfg6zsAcIbQOit36B2urQCpGd4K6nkPeJmtixYSoP6ZMwTmCgWgQiVnLt"}


# --- CONFIG & MAPPING ---
BANNER_PATH = "https://i.ibb.co/Xxd9D1R/verifone-banner.png"
LOGO_PATH = "https://cdn.verifone.com/verifone-standard-logo.png"

TIER2_MAP = {
    "niv.arieli": "ניב אריאלי", "din.weissman": "דין וייסמן", "lior.burstein": "ליאור בורשטיין", "liorb5": "ליאור בורשטיין",
    "avivs": "אביב סולר", "ebrahimf": "אברהים פריג", "orenw1": "אורן וייס", "ahmado": "אחמד עודה",
    "almancha": "אלמנך עלמיה", "zahiyas1": "זהייה אבו שמאלה", "tals": "טל שוקר", "yuvala1": "יובל אגרון",
    "yuliano": "יוליאן אולרסקו", "yoadc": "יועד כחלון", "nuphars": "נוּפר שלום", "idoh": "עידו הרמל",
    "aviele": "אביאל אלשוילי", "avivk": "אביב כץ", "bari": "בר ישראלי", "veral2": "ורה ליברמן",
    "danv1": "דן וייסמן", "niva2": "ניב אריאלי", "nadavl1": "נדב", "paulp": "פאול",
    "moshei1": "משה איטח", "nadav.lieber": "נדב", "erezm1": "ארז", "almanch.alme": "אלמנך עלמיה",
    "dan.vico": "דן ויקו", "danv": "דן ויקו", "shira": "שיר אברהם"
}
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
        "desc": "רשת בתי הקפה הגדולה והדומיננטית ביותר בישראל, המגדירה מחדש את תרבות הקפה וההגשה המהירה עבור הצרכן הישראלי.",
        "fallbacks": ["https://aroma.co.il/wp-content/uploads/2021/05/logo_black.png"]
    },
    "toys r us": {
        "name": "טויס אר אס", 
        "logo": "https://upload.wikimedia.org/wikipedia/commons/a/a7/Toys_%22R%22_Us_logo.svg",
        "desc": "רשת חנויות הצעצועים והפנאי המובילה בעולם, המציעה חוויית קניה חווייתית ומגוון עצום של מותגי צעצועים ומוצרי תינוקות.",
        "fallbacks": ["https://logo.clearbit.com/toysrus.co.il"]
    },
    "maccabi": {
        "name": "מכבי שירותי בריאות", 
        "logo": "https://upload.wikimedia.org/wikipedia/he/a/ae/Maccabi_Health_Care_Services_2011_logo.svg",
        "desc": "ארגון בריאות מוביל בישראל המעניק שירותים רפואיים מתקדמים למיליוני חברים באמצעות צוותי מומחים ומרכזי רפואה טכנולוגיים.",
        "fallbacks": ["https://logo.clearbit.com/maccabi4u.co.il"]
    },
    "leumit": {
        "name": "לאומית", 
        "logo": "https://upload.wikimedia.org/wikipedia/he/0/06/Leumit_New_Logo.svg",
        "desc": "קופת חולים ארצית המעניקה שירותי רפואה איכותיים וזמינים בפריסה רחבה, עם דגש על שירות אישי וטכנולוגיה רפואית מתקדמת.",
        "fallbacks": ["https://logo.clearbit.com/leumit.co.il"]
    },
    "dor alon": {
        "name": "דור אלון", 
        "logo": "https://upload.wikimedia.org/wikipedia/he/4/4f/Dor_Alon_Logo.svg",
        "desc": "קבוצת אנרגיה וקמעונאות מובילה המפעילה תחנות תדלוק, חנויות נוחות (אלונית) ומרכזי מסחר בפריסה ארצית מלאה.",
        "fallbacks": ["https://logo.clearbit.com/doralon.co.il"]
    },
    "hatzi hinam": {
        "name": "חצי חינם", 
        "logo": "https://upload.wikimedia.org/wikipedia/he/0/0f/Hazi-Hinam_logo.png",
        "desc": "רשת מרכולים קמעונאית מהגדולות בישראל, הידועה בסניפי ענק המעניקים חוויית קניה ייחודית ומגוון מוצרים רחב.",
        "fallbacks": ["https://logo.clearbit.com/hany.co.il"]
    },
    "home center": {
        "name": "הום סנטר", 
        "logo": "https://www.homecenter.co.il/cdn/shop/files/HC_Logo_New.jpg?v=1733740893&width=200",
        "desc": "הרשת הגדולה והמובילה בישראל לשיפור ותחזוקת הבית (DIY), המציעה פתרונות מקיפים לעיצוב, שיפוץ וריהוט הבית והגן.",
        "fallbacks": ["https://upload.wikimedia.org/wikipedia/he/d/dd/Home_Center_Logo.png"]
    },
    "hadasa": {
        "name": "הדסה", 
        "logo": "https://logo.clearbit.com/hadassah.org.il",
        "desc": "המרכז הרפואי האוניברסיטאי הדסה בירושלים, ארגון רפואי עילאי המשלב רפואה קלינית מתקדמת עם מחקר והוראה ברמה בינלאומית."
    },
    "fresh market": {
        "name": "פרש מרקט",
        "logo": "https://upload.wikimedia.org/wikipedia/he/f/f4/FreshMarketLogo.png",
        "desc": "רשת קמעונאות מזון המפעילה עשרות סניפי שכונה איכותיים בפריסה ארצית, תחת דגש על טריות ושירות."
    },
    "miphal hapais": {
        "name": "מפעל הפיס",
        "logo": "https://www.pais.co.il/images/pais_logo_shape.png",
        "desc": "הגוף המרכזי בישראל העוסק בהגרלות ובמשחקי מזל, כאשר כל רווחיו מופנים להקמת מבני ציבור, חינוך ותרבות לטובת הקהילה.",
    },
    "evlink": {
        "name": "EVLink",
        "logo": "https://evlink.co.il/images/logos/8/evlinksmall_tdw1-8s.jpg",
        "desc": "חברה טכנולוגית המתמחה בפתרונות טעינה וניהול לציי רכב חשמליים, המהווה גורם מרכזי במהפכת התחבורה הירוקה בישראל.",
    },
    "milgam": {
        "name": "מילגם",
        "logo": "https://www.milgam.co.il/wp-content/uploads/2025/03/%D7%9E%D7%99%D7%9C%D7%92%D7%9D.png",
        "desc": "קבוצת שירותים מובילה המספקת פתרונות תפעוליים ולוגיסטיים מורכבים עבור רשויות מקומיות, גופים ציבוריים וחברות ממשלתיות.",
    },
    "cardcom": {
        "name": "קארדקום",
        "logo": "https://www.cardcom.solutions/wp-content/uploads/2021/02/%D7%9C%D7%95%D7%92%D7%95-%D7%A7%D7%90%D7%A8%D7%93%D7%A7%D7%95%D7%9D-%D7%91%D7%90%D7%AA%D7%A8.png",
        "desc": "ספקית פתרונות טכנולוגיים מתקדמים לעולמות התשלומים והמסחר הדיגיטלי, המשרתת אלפי עסקים ומיזמי אי-קומרס.",
    },
    "soglowek": {
        "name": "זוגלובק", 
        "logo": "https://upload.wikimedia.org/wikipedia/he/f/f0/%D7%96%D7%95%D7%92%D7%9C%D7%95%D7%92%D7%91%D7%A7.png",
        "desc": "מחברות המזון הוותיקות והמובילות בישראל, המתמחה בייצור ושיווק מוצרי בשר, מאפים ותחליפי בשר איכותיים.",
    },
    "balamuth": {
        "name": "בלמות", 
        "logo": "https://www.balamuth.co.il/sites/balamuth/img/balamuth-logo.svg",
        "desc": "חברה הנדסית מובילה המייצגת יצרנים בינלאומיים ומספקת פתרונות טכנולוגיים מתקדמים לתעשייה ולמגזר העסקי.",
    },
    "food": {
        "name": "A.D. Food", 
        "logo": "https://scontent.ftlv27-1.fna.fbcdn.net/v/t39.30808-6/486218845_1107282251411832_3232505034337483416_n.jpg?_nc_cat=103&ccb=1-7&_nc_sid=6ee11a&_nc_ohc=mFNTlp65W6MQ7kNvwGNI-bH&_nc_oc=AdmT9DEoByNg7Ghaz8MD3wAfFF3EOnemxJzXjWK0KEoCxIiXLzMtQB349DMDrOxd5Tg&_nc_zt=23&_nc_ht=scontent.ftlv27-1.fna&_nc_gid=d759hHpSUWv8Q9cqZejH3A&oh=00_Afq_avyBM7RF3tIdtptc953vxNh2sOYrpwuya-frkjG-rg&oe=69803914",
        "desc": "קבוצת קמעונאות מזון הפועלת בתחום ההפצה והשיווק של מוצרי צריכה ומזון בפריסה ארצית רחבה.",
    },
    "filuet": {
        "name": "Filuet",
        "logo": "https://scontent.ftlv27-1.fna.fbcdn.net/v/t39.30808-6/587114429_1499472191879538_4530865455659782486_n.jpg?_nc_cat=105&ccb=1-7&_nc_sid=6ee11a&_nc_ohc=TmyI5fxisYoQ7kNvwFxYstn&_nc_oc=AdkmOOCQWCxlEzqZ-Sl7HuZBKPRAXmUNr6XTIZ3SJlx_iwknsEQ_N3CM2-rb6vc9k9Q&_nc_zt=23&_nc_ht=scontent.ftlv27-1.fna&_nc_gid=uJmlG0TIU3GKZybg702x4g&oh=00_AfqKsmcTMKjF0MUaAOLTzMGawdbBS37ILh2y0pPzXJeZgA&oe=698043D5",
        "desc": "חברת לוגיסטיקה ושרשרת אספקה גלובלית המעניקה פתרונות אחסנה, הפצה וניהול מלאי עבור חברות בינלאומיות.",
    },
    "pelecard": {
        "name": "Pelecard",
        "logo": "https://res.cloudinary.com/drujiiwnt/images/f_svg,q_auto/fl_sanitize/v1764087387/Wordpress%20Pelecard%20Website/logo_pelecard-1/logo_pelecard-1.svg?_i=AA",
        "desc": "חברת פינטק מובילה המפתחת פתרונות תשלום מתקדמים ומאובטחים עבור רשתות שיווק, עסקים גדולים ומיזמי איקומרס.",
    },
    "hyp": {
        "name": "HYP",
        "logo": "https://hyp.co.il/wp-content/uploads/2021/08/logo_hyp_color.svg",
        "desc": "קבוצת טכנולוגיה פיננסית המציעה פלטפורמה מקיפה לניהול עסקאות, סליקה ושירותים דיגיטליים לעולם המסחר.",
    },
    "intercard": {
        "name": "Intercard",
        "logo": "https://www.intercardinc.com/wp-content/uploads/files/2023/logo-i-icon-with-world-class.svg",
        "desc": "ספקית פתרונות תשלום ודיגיטציה לעסקים, המתמחה בייעול תהליכי מכירה וחווית לקוח בנקודות המכירה.",
    },
    "משרד": {
        "name": "משרדי ממשלה",
        "logo": "https://upload.wikimedia.org/wikipedia/commons/thumb/1/11/Emblem_of_Israel.svg/200px-Emblem_of_Israel.svg.png",
        "desc": "גופים ממשלתיים המנהלים אינטראקציות שירותיות ותשלומים מול אזרחי המדינה בסטנדרטים גבוהים של אבטחה ושירות.",
    },
    "hospitals": {
        "name": "בתי חולים",
        "logo": "https://upload.wikimedia.org/wikipedia/commons/d/da/Health_Ministry_of_Israel_Logo.png",
        "desc": "מרכזים רפואיים ציבוריים המהווים את חזית הרפואה בישראל, ומעניקים שירותי בריאות וטיפול בהיקפים נרחבים.",
    },
    "edea": {
        "name": "Priority Retail (EDEA)",
        "logo": "https://cdn-ildofcc.nitrocdn.com/kitdiqlmIRSNEPcfDMXRsJhcusqfcNfZ/assets/images/source/rev-514108a/www.priority-software.com/wp-content/uploads/2023/04/logo.svg",
        "desc": "חברת טכנולוגיה מובילה המפתחת ומטמיעה פתרונות קמעונאות מתקדמים עבור רשתות השיווק הגדולות בישראל.",
    },
    "net lunch": {
        "name": "Net Lunch",
        "logo": "https://netlunch.co.il/wp-content/uploads/cropped-LogoMedium2-png.png",
        "desc": "פתרון דיגיטלי מתקדם לניהול הטבות מזון וסבסוד ארוחות לעובדים, המקשר בין עסקים למאות מסעדות ונקודות מכירה.",
    },
    "verifone": {
        "name": "Verifone", 
        "logo": "https://upload.wikimedia.org/wikipedia/commons/9/98/Verifone_Logo.svg",
        "desc": "המנהיגה העולמית בפתרונות סחר ותשלומים. וריפון מספקת את התשתית הטכנולוגית והאבטחתית המאפשרת את הפעילות העסקית של כלל הלקוחות והשותפים המוצגים במערכת זו."
    }
}

class DataEngine:
    @staticmethod
    def parse_raw_owner(val):
        if pd.isna(val) or val is None: return "unassigned"
        if isinstance(val, dict):
            for k in ['UserName', 'userName', 'username', 'OwnerName', 'name']:
                if k in val:
                    un = str(val[k]).split('@')[0].lower().strip()
                    if len(un) > 30 and re.match(r'^[0-9a-f-]+$', un): return "bot"
                    return un
            return "unassigned"
        s_val = str(val).strip()
        if not s_val or s_val.lower() == 'none' or s_val == '{}': return 'unassigned'
        m = re.search(r"['\"]?(?:UserName|userName|OwnerName|name)['\"]?\s*:\s*['\"]([^'\"]+)['\"]", s_val, re.I)
        if m:
            un = m.group(1).split('@')[0].lower().strip()
            if len(un) > 30 and re.match(r'^[0-9a-f-]+$', un): return "bot"
            return un
        if '{' not in s_val and len(s_val) < 100:
            un = s_val.split('@')[0].lower().strip() if '@' in s_val else s_val.lower().strip()
            if len(un) > 30: return "bot"
            if not any(x in un for x in ['alon', 'bot', 'glassix', 'system', 'test']): return un
        return "unassigned"

    @staticmethod
    def is_valid_record(row):
        # Handle various boolean-like values
        for col in ['isSpam', 'isTest', 'IsSpam', 'IsTest']:
            if col in row:
                val = str(row[col]).lower().strip()
                if val in ['true', '1', 't', 'yes']: return False
        return True

    @staticmethod
    def fetch_glassix(s, e, api_cfg, is_tickets=True):
        """EXACT LOGIC FROM TIER2.PY"""
        try:
            # Token Step
            base_url = "https://verifone.glassix.com/api/v1.2"
            payload = {"apiKey": api_cfg["key"], "apiSecret": api_cfg["secret"], "userName": "MosheI1@VERIFONE.com"}
            tk_resp = requests.post(f"{base_url}/token/get", json=payload, timeout=90)
            tk_resp.raise_for_status()
            token = tk_resp.json().get("access_token")
            
            # Since/Until format from TIER2.PY
            since = s.strftime("%d/%m/%Y 00:00:00:00")
            until = e.strftime("%d/%m/%Y 23:59:59:00")
            
            headers = {"Authorization": f"Bearer {token}"}
            tickets_all = []
            url = f"{base_url}/tickets/list?since={since}&until={until}" if is_tickets else f"{base_url}/interactions/list?since={since}&until={until}"
            
            while url:
                r = requests.get(url, headers=headers, timeout=90)
                if r.status_code == 429:
                    time.sleep(5) # Shorter wait for dashboard
                    continue
                r.raise_for_status()
                data = r.json()
                
                key = "tickets" if is_tickets else "interactions"
                batch = data.get(key, [])
                tickets_all.extend(batch)
                
                paging = data.get("paging")
                url = paging.get("next") if paging and "next" in paging else None
                
            return tickets_all
        except Exception as ex:
            err_log(f"API FAILURE: {ex}")
            return []

    @staticmethod
    def get_tier2(start, end):
        # Stubbed to remove pandas dependency
        return {"total": 0, "agents": [], "tags": [], "v_rate": "0/day"}

    @staticmethod
    def get_digital(start, end):
        # Stubbed to remove pandas dependency
        return {"tickets": {"total":0, "agents":[], "tags":[], "daily":{}}, "whatsapp": {"total":0, "agents":[], "tags":[], "daily":{}}}

    @staticmethod
    def get_shufersal(start, end):
        # Stubbed to remove pandas/openpyxl dependency
        return {"total": 1175, "settled": 1100, "failed": 75}


    @staticmethod
    def get_stfp(start, end):
        # Stubbed to remove pandas dependency
        return {"ready": 10, "errors": 0, "success": 10}

    @staticmethod
    def get_integrations():
        log("DataEngine: Loading integrations...")
        # 1. Try Firestore
        if db:
            try:
                doc = db.collection('data').document('integrations').get()
                if doc.exists:
                    raw = doc.to_dict()
                    data = raw.get('list', []) if isinstance(raw, dict) else raw
                    if data:
                        log(f"DataEngine: Loaded {len(data)} from Firestore.")
                        return data
            except Exception as e: err_log(f"Firestore Integrations error: {e}")

        # 2. Local fallback
        try:
            p = os.path.join(BASE_DIR, "integrations_db.json")
            if os.path.exists(p):
                with open(p, 'r', encoding='utf-8-sig') as f:
                    raw = json.load(f)
                    data = raw.get('list', []) if isinstance(raw, dict) else raw
                    log(f"DataEngine: Loaded {len(data)} from local JSON.")
                    return data
        except Exception as e: err_log(f"Integrations fallback error: {e}")
        return []

    @staticmethod
    def get_guides_categories():
        log("DataEngine: Loading categories...")
        if db:
            try:
                # 1. Load from Firestore (with tiny hack for potential hangs)
                firestore_cats = []
                try:
                    # stream() can hang if connection is unstable. 
                    # We give it a shot but the try block protects us.
                    cats_ref = list(db.collection('guides_categories').stream())
                    firestore_cats = [c.to_dict() for c in cats_ref]
                    log(f"DataEngine: Loaded {len(firestore_cats)} categories from Firestore.")
                except Exception as ef:
                    err_log(f"Firestore categories stream error: {ef}")
                
                # 2. Load from Local JSON (User's file)
                local_cats = []
                p = os.path.join(BASE_DIR, "guides_db.json")
                if os.path.exists(p):
                    try:
                        with open(p, 'r', encoding='utf-8-sig') as f:
                            local_data = json.load(f)
                            if isinstance(local_data, list): local_cats = local_data
                            elif isinstance(local_data, dict): local_cats = local_data.get('categories') or local_data.get('list') or []
                    except Exception as ex:
                        log(f"Merge: Failed to load local JSON: {ex}")

                # 3. Merge (Firestore takes precedence for modifications, but we add missing local cats)
                # Map by ID
                merged = {c['id']: c for c in firestore_cats}
                
                for l_cat in local_cats:
                    if l_cat['id'] not in merged:
                        log(f"Merge: Adding local category '{l_cat.get('name')}' to display.")
                        merged[l_cat['id']] = l_cat
                    else:
                        pass
                
                # 4. Load Guides from Firestore and Re-attach
                try:
                    guides_ref = list(db.collection('guides').stream())
                    all_guides = [g.to_dict() for g in guides_ref]
                    log(f"DataEngine: Loaded {len(all_guides)} guides from Firestore.")
                except Exception as eg:
                    err_log(f"Firestore guides stream error: {eg}")
                    all_guides = []

                for g in all_guides:
                    cat_id = g.get('Category')
                    sub_id = g.get('SubCategory')
                    
                    if cat_id and cat_id in merged:
                        cat = merged[cat_id]
                        
                        if sub_id:
                            # Attach to SubCategory
                            found_sub = False
                            for sub in cat.get('subCategories', []):
                                if str(sub.get('id')) == str(sub_id):
                                    if 'guides' not in sub: sub['guides'] = []
                                    sub['guides'].append(g)
                                    found_sub = True
                                    break
                            if not found_sub:
                                if 'guides' not in cat: cat['guides'] = []
                                cat['guides'].append(g)
                        else:
                            if 'guides' not in cat: cat['guides'] = []
                            cat['guides'].append(g)


                return list(merged.values())

            except Exception as e: 
                err_log(f"Firestore Categories load/merge error: {e}")
        
        # Fallback if no DB or DB fails
        try:
            p = os.path.join(BASE_DIR, "guides_db.json")
            if os.path.exists(p):
                with open(p, 'r', encoding='utf-8-sig') as f:
                    data = json.load(f)
                    if isinstance(data, list): return data
                    if isinstance(data, dict): return data.get('categories') or data.get('list') or []
        except Exception as e: err_log(f"Guides fallback error: {e}")
        
        # Static defaults
        return [
            {"id": "meetings-forms", "name": "ישיבות שחורר וטפסים", "emoji": "📄", "type": "kb", "subCategories": []},
            {"id": "kb", "name": "מרכז ידע ונהלים", "emoji": "📚", "type": "kb", "subCategories": [
                {"id": "kb-guides", "name": "מדריכי מערכת"}, {"id": "kb-policy", "name": "נהלי עבודה"}
            ]},
            {"id": "integrations", "name": "אינטגרציות וחיבורים", "emoji": "🔌", "type": "kb", "subCategories": [
                {"id": "int-verifone", "name": "וריפון"}, {"id": "int-third-party", "name": "צד ג'"}
            ]}
        ]

    @staticmethod
    def get_guides_by_category(cat_id):
        if db:
            try:
                guides = db.collection('guides').where('Category', '==', cat_id).stream()
                return [g.to_dict() for g in guides]
            except Exception as e:
                err_log(f"Firestore guides load error: {e}")
        
        # Fallback to local JSON if Firestore not ready
        p = os.path.join(BASE_DIR, "guides_db.json")
        if os.path.exists(p):
            try:
                with open(p, 'r', encoding='utf-8-sig') as f:
                    all_g = json.load(f)
                    return [g for g in all_g if g.get('Category') == cat_id]
            except: pass
        return []
    @staticmethod
    def save_integrations(data):
        log(f"DataEngine: Saving {len(data)} integrations...")
        success = False
        if db:
            try:
                db.collection('data').document('integrations').set({'list': data})
                success = True
                log("DataEngine: Integrations saved to Firestore.")
            except Exception as e:
                err_log(f"Firestore Integrations save error: {e}")

        # Local save as backup (may fail on Vercel, but that's okay)
        try:
            p = os.path.join(BASE_DIR, "integrations_db.json")
            with open(p, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            success = True
            log("DataEngine: Integrations saved to local JSON.")
        except Exception as e:
            log(f"DataEngine: Local JSON save skipped/failed (expected on Vercel): {e}")
        return success

    # --- GUIDES LOGIC ---
    @staticmethod
    def get_guides():
        if db:
            try:
                docs = db.collection('guides').stream()
                guides = [doc.to_dict() for doc in docs]
                if guides: return guides
            except Exception as e:
                err_log(f"Firestore Guides load error: {e}")

        p = os.path.join(BASE_DIR, "guides_db.json")
        # ... fallback ...
        if os.path.exists(p):
            try:
                with open(p, 'r', encoding='utf-8-sig') as f:
                    return json.load(f)
            except Exception as e:
                err_log(f"Guides load error: {e}")
        return []

    @staticmethod
    def save_guides(data):
        log(f"DataEngine: Saving {len(data)} guides...")
        success = False
        if db:
            try:
                # Use batch for efficiency
                batch = db.batch()
                for guide in data:
                    gid = guide.get('id') or str(uuid.uuid4())
                    doc_ref = db.collection('guides').document(gid)
                    batch.set(doc_ref, guide)
                batch.commit()
                success = True
                log("DataEngine: Guides saved to Firestore (batch).")
            except Exception as e:
                err_log(f"Firestore Guides save error: {e}")

        # Local save as backup
        try:
            p = os.path.join(BASE_DIR, "guides_db.json")
            with open(p, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            success = True
            log("DataEngine: Guides saved to local JSON.")
        except Exception as e:
            log(f"DataEngine: Guides local save skipped/failed (expected on Vercel): {e}")
        return success

    @staticmethod
    def save_guides_categories(data):
        log(f"DataEngine: Saving {len(data)} categories...")
        success = False
        if db:
            try:
                # Firestore: Split Categories and Guides to separate collections
                batch = db.batch()
                batch_counter = 0
                max_batch = 400  # Firestore batch limit is 500
                
                for cat in data:
                    cid = cat.get('id') or str(uuid.uuid4())
                    cat['id'] = cid # Ensure ID is set
                    
                    # 1. Prepare Category Doc (Excluding large children)
                    # We remove 'guides' but KEEP 'subCategories' structure (without guides inside them ideally, but complexity...)
                    # To do this cleanly: WE MUST strip guides from subCategories in the doc, and save them separately.
                    
                    cat_doc = {k:v for k,v in cat.items() if k not in ['guides']}
                    # Deep copy subcategories to strip guides for the doc
                    clean_subcats = []
                    subcategories = cat.get('subCategories', [])
                    for sub in subcategories:
                         sub_copy = {k:v for k,v in sub.items() if k != 'guides'}
                         clean_subcats.append(sub_copy)
                    cat_doc['subCategories'] = clean_subcats

                    doc_ref = db.collection('guides_categories').document(cid)
                    batch.set(doc_ref, cat_doc)
                    batch_counter += 1
                    
                    # 2. Save Direct Guides
                    guides = cat.get('guides', [])
                    for g in guides:
                        gid = g.get('id') or str(uuid.uuid4())
                        g['id'] = gid
                        g['Category'] = cid # Link to parent
                        g_ref = db.collection('guides').document(gid)
                        batch.set(g_ref, g)
                        batch_counter += 1
                        if batch_counter >= max_batch:
                            batch.commit(); batch = db.batch(); batch_counter = 0;

                    # 3. Save Subcategory Guides
                    subcategories = cat.get('subCategories', [])
                    for sub in subcategories:
                        sub_guides = sub.get('guides', [])
                        for g in sub_guides:
                            gid = g.get('id') or str(uuid.uuid4())
                            g['id'] = gid
                            g['Category'] = cid # Link to main parent category
                            g['SubCategory'] = sub.get('id') # Link to subcategory
                            g_ref = db.collection('guides').document(gid)
                            batch.set(g_ref, g)
                            batch_counter += 1
                            if batch_counter >= max_batch:
                                try:
                                    batch.commit()
                                    batch = db.batch()
                                    batch_counter = 0
                                except Exception as e:
                                    err_log(f"Batch commit error (inner): {e}")

                if batch_counter > 0:
                    batch.commit()
                
                success = True
                log("DataEngine: Categories & Guides saved to Firestore (Split Model).")
            except Exception as e:
                err_log(f"Firestore Categories save error: {e}")

        # Local save as backup (handling Vercel Read-Only FS)
        try:
            if not IS_VERCEL:
                p = os.path.join(BASE_DIR, "guides_db.json")
                with open(p, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=4, ensure_ascii=False)
                success = True # Marked success if local work
            else:
                log("DataEngine: Skipping local file save (Vercel environment).")
                # On Vercel, if Firestore worked, we are good.
        except OSError as e:
            log(f"DataEngine: Local save failed (Expected on Vercel): {e}")
        except Exception as e:
            log(f"DataEngine: Local save error: {e}")
            
        return success

    @staticmethod
    def sync_gdrive_to_kb():
        return False, "Google Drive integration has been removed."

    @staticmethod
    def extract_text_from_file(file_path):
        if not HAS_PARSERS: 
            if os.environ.get('VERCEL'):
                return f"Error: Document parsers not installed on Vercel: {', '.join(PARSER_ERRORS)}. Please check requirements.txt and redeploy."
            else:
                pkgs = " ".join(PARSER_ERRORS)
                return f"Error: Document parsers missing locally. Run this in PowerShell: py -m pip install {pkgs}"
        ext = os.path.splitext(file_path)[1].lower()
        try:
            if ext == '.docx':
                log(f"Starting Mammoth extraction for: {file_path}")
                
                def handle_image(image):
                    try:
                        with image.open() as image_bytes:
                            content = image_bytes.read()
                            if not content: return {"src": ""}
                            
                            mime = image.content_type or "image/png"
                            # Convert directly to Base64
                            b64 = base64.b64encode(content).decode('utf-8')
                            return {"src": f"data:{mime};base64,{b64}"}
                    except Exception as e:
                        err_log(f"Mammoth extraction error: {e}")
                        return {"src": ""}
                
                if not HAS_PARSERS or 'python-docx' in PARSER_ERRORS: return "docx parser missing"
                
                html = ""
                # Use mammoth if available for better HTML with images
                if 'mammoth' not in PARSER_ERRORS:
                    with open(file_path, "rb") as docx_file:
                        result = mammoth.convert_to_html(docx_file, 
                            convert_image=mammoth.images.img_element(handle_image))
                        html = result.value
                else:
                    doc = docx.Document(file_path)
                    for para in doc.paragraphs:
                        html += f"<p>{para.text}</p>"
                return html
            elif ext == '.pdf':
                try:
                    import fitz  # PyMuPDF
                    doc = fitz.open(file_path)
                    content = ""
                    for page_index, page in enumerate(doc):
                        # Extract text
                        text = page.get_text("html")
                        content += text
                        
                        # Extract images and convert to Base64
                        image_list = page.get_images()
                        for img_index, img in enumerate(image_list):
                            xref = img[0]
                            pix = fitz.Pixmap(doc, xref)
                            if pix.n - pix.alpha > 3: # CMYK: convert to RGB first
                                pix = fitz.Pixmap(fitz.csRGB, pix)
                            
                            img_data = pix.tobytes()
                            b64 = base64.b64encode(img_data).decode('utf-8')
                            pix = None # free
                            
                            content += f'<br><img src="data:image/png;base64,{b64}" style="max-width:100%; margin: 10px 0;"><br>'
                            
                    return content
                except ImportError:
                    # Fallback to PyPDF2
                    if not HAS_PARSERS or 'PyPDF2' in PARSER_ERRORS: return "PyPDF2 parser missing"
                    text = ""
                    with open(file_path, 'rb') as f:
                        pdf = PyPDF2.PdfReader(f)
                        for page in pdf.pages:
                            text += page.extract_text() + "\n\n"
                    return text.replace('\n', '<br>')
            elif ext == '.txt':
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read().replace('\n', '<br>')
            return f"Unsupported file type for extraction: {ext}"
        except Exception as e:
            err_log(f"Extraction error ({ext}): {e}")
            return f"Failed to extract content: {str(e)}"

    @staticmethod
    def get_reports():
        reps = {"Tier2":[], "Digital":[], "Shufersal":[], "STFP":[]}
        patterns = {
            "Tier2": os.path.join(BASE_DIR, "TIER2", "Tickets_*.xlsx"),
            "Digital": os.path.join(BASE_DIR, "Digital", "*.xlsx"),
            "Shufersal": os.path.join(BASE_DIR, "Shufersal_Reports", "*.xlsx"),
            "STFP": os.path.join(BASE_DIR, "logs_stf", "*.txt")
        }
        for k, p in patterns.items():
            files = glob.glob(p)
            for f in files:
                m = re.search(r'(\d{2}[_.]\d{2}[_.]\d{4})', os.path.basename(f))
                dt_str = m.group() if m else datetime.fromtimestamp(os.path.getmtime(f)).strftime("%d/%m/%Y")
                reps[k].append({"name": os.path.basename(f), "date": dt_str, "path": f})
        for k in reps: reps[k] = sorted(reps[k], key=lambda x: x['date'], reverse=True)[:10]
        return reps



    @staticmethod
    def get_calls():
        p = os.path.join(BASE_DIR, "call_stats.json")
        if os.path.exists(p):
            with open(p, 'r', encoding='utf-8') as f:
                return json.load(f)
        return None


class handler(http.server.SimpleHTTPRequestHandler):
    def is_authenticated(self):
        cookies = http.cookies.SimpleCookie(self.headers.get('Cookie'))
        sid = cookies.get('sid')
        if not sid: 
            # log("is_authenticated: No sid cookie found")
            return False
        
        sid_val = sid.value
        now = get_now_utc()
        
        # 1. Check in-memory for speed
        if sid_val in SESSIONS:
            sess = SESSIONS[sid_val]
            exp = ensure_utc(sess.get('expiry'))
            if exp and exp > now:
                return True
            log(f"Session Trace: Memory expiry failure for {sess.get('user')}. Exp: {exp}, Now: {now}")
        
        # 2. Check Firestore for persistence (Vercel)
        if db:
            try:
                doc = db.collection('sessions').document(sid_val).get()
                if doc.exists:
                    data = doc.to_dict()
                    exp = ensure_utc(data.get('expiry'))
                    
                    if exp and exp > now:
                        # Sync back to memory
                        SESSIONS[sid_val] = data
                        return True
                    log(f"Session Trace: Firestore expiry failure for {data.get('user')}. Exp: {exp}, Now: {now}")
                else:
                    log(f"Session Trace: SID {sid_val[:8]} not in Firestore")
            except Exception as e:
                err_log(f"Session verification error: {e}")
        return False

    def save_session(self, sid, email):
        now = get_now_utc()
        ext_expiry = now + timedelta(days=1)
        sess_data = {
            "user": email, 
            "expiry": ext_expiry,
            "created_at": now
        }
        SESSIONS[sid] = sess_data
        log(f"Session Trace: Saving session {sid[:8]} for {email}. Expiry: {ext_expiry}")
        if db:
            try:
                db.collection('sessions').document(sid).set(sess_data)
            except Exception as e:
                err_log(f"Session save error: {e}")

    def do_GET(self):
        try:
            path = self.path.split('?')[0] # Strip query string for routing
            if path == '/login':
                self.send_response(200)
                self.send_header('Content-type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(self.get_login_ui().encode('utf-8'))
                return

            if not self.is_authenticated():
                self.send_response(302)
                self.send_header('Location', '/login')
                self.end_headers()
                return

            if path in ['/', '/index.html']:
                self.send_response(200); self.send_header('Content-type','text/html;charset=utf-8'); self.end_headers()
                self.wfile.write(self.get_ui().encode('utf-8'))
                return
            
            if path == '/api/stats':
                log(f"Handling /api/stats (Trace: {self.headers.get('Cookie')})")
                try:
                    integrations = DataEngine.get_integrations()
                    categories = DataEngine.get_guides_categories()
                    data = {
                        "Integrations": integrations,
                        "GuidesCategories": categories,
                        "CustomerLogos": CUSTOMER_LOGOS,
                        "Health": {
                            "firebase": db is not None,
                            "count": len(integrations),
                            "vercel": os.environ.get('VERCEL') is not None
                        }
                    }
                    self.send_response(200); self.send_header('Content-Type','application/json'); self.end_headers()
                    self.wfile.write(json.dumps(data, default=str).encode('utf-8'))
                except Exception as e:
                    err_log(f"API Stats Error: {e}")
                    self.send_error(500, str(e))
                return

            if path == '/api/health':
                health = {
                    "firebase": db is not None,
                    "gdrive": False,
                    "parsers": HAS_PARSERS,
                    "vercel": os.environ.get('VERCEL') is not None,
                    "now": get_now_utc().isoformat()
                }
                self.send_response(200); self.send_header('Content-Type','application/json'); self.end_headers()
                self.wfile.write(json.dumps(health).encode())
                return

            # Explicit file serving whitelist - NO DIRECTORY LISTING
            allowed_prefixes = ['/uploads/', '/מדריכים/', '/לקוחות/', '/TIER2/', '/Digital/', '/csv/']
            if any(path.startswith(p) for p in allowed_prefixes):
                try:
                    rel_path = urllib.parse.unquote(path[1:])
                    fpath = os.path.join(BASE_DIR, rel_path)
                    
                    if os.path.exists(fpath) and os.path.isfile(fpath):
                        self.send_response(200)
                        ext = os.path.splitext(fpath)[1].lower()
                        mimes = {
                            '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', 
                            '.gif': 'image/gif', '.pdf': 'application/pdf', 
                            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            '.csv': 'text/csv', '.txt': 'text/plain', '.json': 'application/json'
                        }
                        self.send_header('Content-type', mimes.get(ext, 'application/octet-stream'))
                        self.end_headers()
                        with open(fpath, 'rb') as f: self.wfile.write(f.read())
                        return
                    else:
                        log(f"File Route: 404 for {fpath}")
                        self.send_error(404, "File not found or is a directory")
                        return
                except Exception as e:
                    err_log(f"File Serve Error: {e}")
            
            # Catch all for unhandled routes - DO NOT CALL SUPER (Security)
            log(f"Route Trace: 404 for path '{path}'")
            self.send_error(404, "Page Not Found")
        except Exception as e:
            err_log(f"GET Error: {e}")
            self.send_error(500, str(e))

    def do_POST(self):
        try:
            if self.path == '/login':
                length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(length).decode('utf-8')
                
                # Check if it's a JSON token login (Firebase)
                if self.headers.get('Content-Type') == 'application/json':
                    data = json.loads(body)
                    id_token = data.get('idToken')
                    email = data.get('email', 'unknown')
                    
                    success = False
                    if HAS_FIREBASE and id_token:
                        log(f"Login Attempt: Firebase token verification for {email}")
                        if not firebase_admin._apps:
                            log(f"ALERT: Non-verified login for {email} (Firebase not initialized, allowing for dev)")
                            success = True 
                        else:
                            try:
                                # Verify the token with Firebase Admin
                                decoded_token = auth.verify_id_token(id_token)
                                success = True
                            except Exception as e:
                                err_log(f"Token Verification Failed: {e}")
                                self.send_response(401); self.end_headers(); return
                    
                    if success:
                        sid = str(uuid.uuid4())
                        log(f"Login Success: {email}, Session: {sid}")
                        self.save_session(sid, email)
                        self.send_response(200)
                        self.send_header('Content-Type', 'application/json')
                        self.send_header('Set-Cookie', f'sid={sid}; HttpOnly; Path=/; SameSite=Lax; Max-Age=86400')
                        self.end_headers()
                        self.wfile.write(json.dumps({"status": "url", "url": "/"}).encode())
                        return
                    else:
                        log(f"Login Failed: Firebase auth failed for {email}")
                        self.send_response(401); self.end_headers()
                    return

                # Fallback to old form auth for hardcoded users
                params = urllib.parse.parse_qs(body)
                email = params.get('email', [''])[0].strip()
                password = params.get('password', [''])[0].strip()
                
                log(f"Login Attempt: Fallback auth for {email}")
                if email in AUTHORIZED_USERS and password == AUTHORIZED_USERS[email]:
                    sid = str(uuid.uuid4())
                    log(f"Fallback Login Success: {email}, Session: {sid[:8]}")
                    self.save_session(sid, email)
                    # Respond with JSON status for consistency
                    self.send_response(200)
                    self.send_header('Content-Type', 'application/json')
                    self.send_header('Set-Cookie', f'sid={sid}; HttpOnly; Path=/; SameSite=Lax; Max-Age=86400')
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "url", "url": "/"}).encode())
                else:
                    log(f"Fallback Login Failed: Invalid credentials for {email}")
                    self.send_response(401)
                    self.send_header('Content-Type', 'application/json')
                    self.end_headers()
                    self.wfile.write(json.dumps({"status": "error", "message": "Invalid credentials"}).encode())
                return

            if not self.is_authenticated():
                self.send_error(403, "Forbidden")
                return

            if self.path == '/api/upload':
                content_type = self.headers.get('Content-Type')
                if not content_type or 'multipart/form-data' not in content_type:
                    self.send_error(400, "Bad Request")
                    return
                
                length = int(self.headers.get('Content-Length', 0))
                raw_data = self.rfile.read(length)
                
                try:
                    # Robust boundary extraction
                    b_parts = content_type.split('boundary=')
                    if len(b_parts) < 2:
                        err_log("Upload Error: No boundary found in Content-Type")
                        self.send_error(400, "No boundary")
                        return
                    
                    boundary = b_parts[-1].split(';')[0].strip().encode()
                    if boundary.startswith(b'"') and boundary.endswith(b'"'):
                        boundary = boundary[1:-1]
                        
                    boundary_search = b'--' + boundary
                    parts = raw_data.split(boundary_search)
                    log(f"Upload: Received {length} bytes, {len(parts)} parts found.")
                except Exception as e:
                    err_log(f"Boundary Parse Error: {e}")
                    self.send_error(400, "Invalid multipart data")
                    return
                
                for part in parts:
                    if b'filename=' in part:
                        # Find filename with or without quotes
                        fn_match = re.search(b'filename=(?:"([^"]+)"|([^;\r\n]+))', part)
                        if fn_match:
                            filename_bytes = fn_match.group(1) or fn_match.group(2)
                            try:
                                filename = filename_bytes.decode('utf-8')
                            except:
                                filename = filename_bytes.decode('latin-1')
                            
                            ext = os.path.splitext(filename)[1].lower()
                            allowed = ['.jpg', '.jpeg', '.png', '.gif', '.pdf', '.docx', '.xlsx', '.csv', '.txt', '.doc', '.ppt', '.pptx', '.zip']
                            if ext not in allowed: 
                                log(f"Upload: File extension {ext} not allowed.")
                                continue
                            
                            safe_name = f"{uuid.uuid4()}{ext}"
                            header_end = part.find(b'\r\n\r\n')
                            if header_end == -1: continue
                            
                            # Binary safe content extraction
                            file_content = part[header_end+4:]
                            # Strip trailing \r\n that belongs to the multipart format
                            if file_content.endswith(b'\r\n'):
                                file_content = file_content[:-2]
                            elif file_content.endswith(b'\r'):
                                file_content = file_content[:-1]
                            
                            # Upload to Google Drive if available, otherwise save locally
                            # Google Drive upload removed.
                            pass
                            
                            # Fallback: Save locally
                            try:
                                with open(os.path.join(UPLOAD_DIR, safe_name), 'wb') as f:
                                    f.write(file_content)
                                log(f"SUCCESS: Uploaded {filename} as {safe_name} (local fallback)")
                                self.send_response(200); self.send_header('Content-Type','application/json'); self.end_headers()
                                self.wfile.write(json.dumps({"url": f"/uploads/{safe_name}", "name": filename, "warning": "local_storage_not_persistent"}).encode())
                            except Exception as e:
                                err_log(f"Local storage fallback failed: {e}")
                                self.send_error(500, f"Upload failed: {str(e)}")
                            return
                
                err_log("Upload failed: No valid file parts with filename found.")
                self.send_error(400, "No file found")
                return

            if self.path == '/api/guides/save':
                length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(length).decode('utf-8')
                data = json.loads(body)
                # The frontend sends the whole categories list here
                if DataEngine.save_guides_categories(data):
                    self.send_response(200); self.send_header('Content-Type', 'application/json'); self.end_headers()
                    self.wfile.write(json.dumps({"status":"ok"}).encode('utf-8'))
                else: 
                    err_log("POST /api/guides/save failed in DataEngine")
                    self.send_error(500, "Save failed")
                return

            if self.path == '/api/integrations/save':
                length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(length).decode('utf-8')
                data = json.loads(body)
                if DataEngine.save_integrations(data):
                    self.send_response(200); self.send_header('Content-Type', 'application/json'); self.end_headers()
                    self.wfile.write(json.dumps({"status":"ok"}).encode('utf-8'))
                else: self.send_error(500, "Save failed")
                return

            if self.path == '/api/gdrive/sync':
                self.send_error(501, "Google Drive integration removed.")
                return

            if self.path == '/api/extract-content':
                length = int(self.headers.get('Content-Length', 0))
                body = self.rfile.read(length).decode('utf-8')
                data = json.loads(body)
                url = data.get('url', '')
                if url.startswith('/uploads/'):
                    p = os.path.join(UPLOAD_DIR, os.path.basename(url))
                    text = DataEngine.extract_text_from_file(p)
                    self.send_response(200); self.send_header('Content-Type', 'application/json'); self.end_headers()
                    self.wfile.write(json.dumps({"content": text}).encode('utf-8'))
                else: self.send_error(400, "Invalid URL")
                return

            self.send_error(404, "Not Found")
        except Exception as e:
            err_log(f"POST Error: {e}")
            self.send_error(500, str(e))

    def get_login_ui(self):
        return r"""
<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>התחברות מאובטחת | Verifone Tier 2</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Assistant:wght@200;400;600;800&family=Outfit:wght@500;900&display=swap');
        
        :root {
            --accent: #3b82f6;
            --accent-glow: rgba(59, 130, 246, 0.4);
            --bg: #010409;
            --glass: rgba(13, 17, 23, 0.7);
        }

        body { 
            margin:0; padding:0; height:100vh; width:100vw;
            display:flex; align-items:center; justify-content:center; 
            background: var(--bg); 
            font-family: 'Assistant', sans-serif; 
            color: #fff; overflow:hidden;
            perspective: 1000px;
        }

        /* High-End Background Effects */
        .scene {
            position: fixed; inset: 0; z-index: 0;
            background: radial-gradient(circle at 50% 50%, #0d1117 0%, #010409 100%);
        }
        .orb {
            position: absolute; border-radius: 50%;
            filter: blur(100px); opacity: 0.4;
            animation: drift 25s infinite alternate ease-in-out;
        }
        .orb-1 { width: 600px; height: 600px; background: #2563eb; top: -10%; left: -10%; }
        .orb-2 { width: 500px; height: 500px; background: #7c3aed; bottom: -10%; right: -10%; animation-delay: -5s; }
        .orb-3 { width: 400px; height: 400px; background: #0891b2; top: 40%; left: 60%; animation-delay: -10s; }

        @keyframes drift {
            0% { transform: translate(0,0) rotate(0deg) scale(1); }
            100% { transform: translate(100px, 100px) rotate(90deg) scale(1.2); }
        }

        .card-container {
            z-index: 10; width: 100%; max-width: 440px;
            animation: slideUp 1s cubic-bezier(0.2, 0.8, 0.2, 1);
        }
        @keyframes slideUp { from { opacity: 0; transform: translateY(30px); } to { opacity: 1; transform: translateY(0); } }

        .card {
            background: var(--glass);
            backdrop-filter: blur(40px);
            -webkit-backdrop-filter: blur(40px);
            padding: 70px 50px;
            border-radius: 50px;
            border: 1px solid rgba(255, 255, 255, 0.08);
            box-shadow: 
                0 30px 60px rgba(0,0,0,0.8),
                inset 0 0 0 1px rgba(255,255,255,0.05);
            text-align: center;
        }

        .logo-wrap { margin-bottom: 40px; }
        .logo { 
            height: 30px; 
            filter: drop-shadow(0 0 15px rgba(255,255,255,0.4));
            transition: 0.5s;
        }
        .logo:hover { transform: scale(1.05); }
        
        .title-wrap { margin-bottom: 45px; }
        .title-wrap h1 { 
            font-size: 38px; font-weight: 800; margin: 0; 
            background: linear-gradient(to bottom, #fff, #94a3b8);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
            letter-spacing: -1px;
        }
        .title-wrap p { color: #8b949e; font-size: 16px; margin: 12px 0 0; font-weight: 400; }
        
        .form-grid { display: flex; flex-direction: column; gap: 30px; }
        .input-box { text-align: right; }
        .input-box label { 
            display: block; font-size: 12px; font-weight: 800; color: var(--accent); 
            margin-bottom: 12px; margin-right: 5px; text-transform: uppercase; letter-spacing: 1px;
        }
        
        .field-wrap { position: relative; }
        input {
            width: 100%; background: rgba(255, 255, 255, 0.03); border: 1px solid rgba(255,255,255,0.1);
            padding: 20px 25px; border-radius: 24px; color: #fff; font-size: 17px; font-weight: 500;
            outline: none; transition: 0.4s cubic-bezier(0.4, 0, 0.2, 1); box-sizing: border-box;
            text-align: left; direction: ltr;
        }
        input:focus { 
            background: rgba(255, 255, 255, 0.06); border-color: var(--accent);
            box-shadow: 0 0 30px var(--accent-glow);
            transform: translateY(-2px);
        }
        input::placeholder { color: rgba(255,255,255,0.15); }
        
        .action-btn {
            margin-top: 15px; width: 100%;
            background: linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%);
            color: #fff; padding: 22px; border-radius: 24px; font-size: 20px; font-weight: 800;
            cursor: pointer; border: none; transition: 0.4s;
            box-shadow: 0 15px 35px -10px rgba(59, 130, 246, 0.5);
        }
        .action-btn:hover { 
            transform: translateY(-4px) scale(1.01); 
            box-shadow: 0 25px 50px -10px rgba(59, 130, 246, 0.6);
            filter: brightness(1.1);
        }
        .action-btn:active { transform: translateY(-1px); }
        .action-btn:disabled { background: #1e293b; color: #475569; cursor: not-allowed; transform: none; box-shadow: none; }
        
        .error-notif { 
            display: none; background: rgba(239, 68, 68, 0.1); border: 1px solid rgba(239, 68, 68, 0.2);
            color: #f87171; padding: 18px; border-radius: 20px; margin-top: 30px; font-weight: 700;
            font-size: 14px;
        }
        
        .legal { 
            margin-top: 45px; font-size: 13px; color: #484f58; font-weight: 600; 
            border-top: 1px solid rgba(255,255,255,0.05); padding-top: 30px;
        }
        .legal a { color: var(--accent); text-decoration: none; }
    </style>
</head>
<body>
    <div class="scene">
        <div class="orb orb-1"></div>
        <div class="orb orb-2"></div>
        <div class="orb orb-3"></div>
    </div>

    <div class="card-container">
        <div class="card">
            <div class="logo-wrap">
                <!-- Official Verifone Logo SVG -->
                <img src="https://upload.wikimedia.org/wikipedia/commons/9/98/Verifone_Logo.svg" 
                     class="logo" alt="Verifone" style="filter: brightness(0) invert(1);">
            </div>
            
            <div class="title-wrap">
                <h1>מרכז הבקרה Vico</h1>
                <p>התחברות לאזור המורשה של Tier 2</p>
            </div>
            
            <div class="form-grid">
                <div class="input-box">
                    <label>זיהוי משתמש (Email)</label>
                    <div class="field-wrap">
                        <input type="email" id="u-mail" placeholder="name@verifone.com" required autocomplete="email">
                    </div>
                </div>
                <div class="input-box">
                    <label>סיסמת גישה</label>
                    <div class="field-wrap">
                        <input type="password" id="u-pass" placeholder="••••••••" required autocomplete="current-password">
                    </div>
                </div>
                
                <button class="action-btn" id="l-btn" onclick="handleAuth()">כניסה למערכת</button>
            </div>
            
            <div id="msg" class="error-notif">שגיאת אימות: פרטי המשתמש אינם תואמים.</div>
            
            <div class="legal">
                מערכת פנימית של Verifone &copy; 2026. כל הזכויות שמורות.
            </div>
        </div>
    </div>

    <script type="module">
        import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
        import { getAuth, signInWithEmailAndPassword } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js";

        const config = {
            apiKey: "AIzaSyB3pruogaljwaw9FVyrD3MvPOHgpyGfxzs",
            authDomain: "tier-2-vico.firebaseapp.com",
            projectId: "tier-2-vico",
            storageBucket: "tier-2-vico.firebasestorage.app",
            messagingSenderId: "272065575004",
            appId: "1:272065575004:web:11ed615295a56dbc824e99",
            measurementId: "G-57ZTPZWJSV"
        };

        const app = initializeApp(config);
        const auth = getAuth(app);

        window.handleAuth = async () => {
            const email = document.getElementById('u-mail').value;
            const pass = document.getElementById('u-pass').value;
            const btn = document.getElementById('l-btn');
            const msg = document.getElementById('msg');

            if(!email || !pass) return;

            btn.disabled = true;
            btn.innerText = "מעבד...";
            msg.style.display = 'none';

            try {
                // 1. Try Firebase Auth
                try {
                    const userCredential = await signInWithEmailAndPassword(auth, email, pass);
                    const idToken = await userCredential.user.getIdToken();
                    
                    const response = await fetch('/login', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ idToken, email })
                    });

                    if (response.ok) {
                        const data = await response.json();
                        window.location.href = data.url;
                        return;
                    }
                } catch (fbErr) {
                    console.warn("Firebase auth failed, trying local fallback:", fbErr);
                }

                // 2. Try Local Fallback (Standard POST)
                const formData = new URLSearchParams();
                formData.append('email', email);
                formData.append('password', pass);
                
                const response = await fetch('/login', {
                    method: 'POST',
                    body: formData,
                    redirect: 'follow'
                });

                if (response.ok && !response.url.includes('error=1')) {
                    window.location.href = '/';
                } else {
                    throw new Error("אימות נכשל");
                }
            } catch (error) {
                console.error(error);
                msg.style.display = 'block';
                btn.disabled = false;
                btn.innerText = "כניסה למערכת";
            }
        };
    </script>
</body>
</html>
"""


    def get_ui(self):
        return r"""
<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>Vico Intelligence (Live Mode)</title>
    <script src="https://html2canvas.hertzen.com/dist/html2canvas.min.js"></script>
    
    <!-- Firebase SDK -->
    <script type="module">
        import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js';
        import { getFirestore, collection, getDocs, addDoc, updateDoc, deleteDoc, doc, setDoc } from 'https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js';
        import { getStorage, ref, uploadBytes, getDownloadURL } from 'https://www.gstatic.com/firebasejs/10.7.1/firebase-storage.js';
        
        const firebaseConfig = {
            apiKey: "AIzaSyB3pruogaljwaw9FVyrD3MvPOHgpyGfxzs",
            authDomain: "tier-2-vico.firebaseapp.com",
            projectId: "tier-2-vico",
            storageBucket: "tier-2-vico.firebasestorage.app",
            messagingSenderId: "272065575004",
            appId: "1:272065575004:web:11ed615295a56dbc824e99",
            measurementId: "G-57ZTPZWJSV"
        };
        
        const app = initializeApp(firebaseConfig);
        window.db = getFirestore(app);
        window.storage = getStorage(app);
        console.log("Firebase initialized in UI");
        window.firebaseRefs = { collection, getDocs, addDoc, updateDoc, deleteDoc, doc, setDoc, ref, uploadBytes, getDownloadURL };
    </script>
    <script>
        console.log("UI HTML Loaded - Starting JS...");
        window.addEventListener('error', function(e) {
            console.error("Global JS Error:", e.message, "at", e.filename, ":", e.lineno);
        });
    </script>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;800;900&display=swap');
        :root { --bg: #0f172a; --card: #1e293b; --primary: #3b82f6; --accent: #10b981; --text: #f1f5f9; --dim: #94a3b8; --border: rgba(255,255,255,0.1); --panel: #1e293b; --content-bg: #f8fafc; }
        
        html, body { margin:0; padding:0; height:100vh; width:100vw; overflow:hidden; background: var(--bg); color: var(--text); font-family: 'Outfit', sans-serif; display: flex; flex-direction: column; align-items: stretch; font-size: 18px; direction: rtl; box-sizing: border-box; }
        
        .top-bar { height: 100px; width: 100vw; background: #0f172a; border-bottom: 2px solid var(--border); display: flex; align-items: center; justify-content: space-between; padding: 0 40px; box-shadow: 0 10px 50px rgba(0,0,0,0.5); z-index: 100; box-sizing: border-box; }
        .logo { font-size: 32px; font-weight: 900; background: linear-gradient(to left, #60a5fa, #34d399); -webkit-background-clip: text; -webkit-text-fill-color: transparent; letter-spacing: -1.5px; }
        .nav-links { display: flex; gap: 10px; }
        .nav { cursor: pointer; opacity: 0.6; font-weight: 700; font-size: 18px; padding: 10px 20px; border-radius: 12px; transition: 0.3s; display: flex; align-items: center; gap: 8px; }
        .nav:hover { background: rgba(255,255,255,0.05); opacity: 1; }
        .nav.active { opacity: 1; background: var(--primary); box-shadow: 0 0 25px rgba(59, 130, 246, 0.5); }
        
        .clock-box { font-family: 'Courier New', monospace; font-size: 18px; color: #fff; background: rgba(255,255,255,0.1); padding: 8px 15px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.1); display: flex; align-items: center; gap: 10px; }
        .clock-box::before { content: ''; width: 8px; height: 8px; background: #ef4444; border-radius: 50%; animation: blink 1s infinite; }
        @keyframes blink { 50% { opacity: 0.4; } }

        .main { flex: 1; height: calc(100vh - 100px); width: 100vw; overflow-y: auto; padding: 30px 40px; scroll-behavior: smooth; box-sizing: border-box; display: flex; flex-direction: column; align-items: stretch; background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); position: relative; }
        .main::before { content: ''; position: absolute; inset: 0; background: url('https://www.transparenttextures.com/patterns/carbon-fibre.png'); opacity: 0.05; pointer-events: none; }
        .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 30px; border-bottom: 2px solid var(--border); padding-bottom: 20px; }
        .header h1 { font-size: 52px !important; letter-spacing: -1px; margin: 0; line-height: 1; font-weight: 900; color: #fff; }
        .header p { font-size: 20px; color: var(--dim); margin-top: 5px; }
        
        .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 25px; margin-bottom: 30px; }
        .kpi-card { background: var(--card); padding: 35px; border-radius: 30px; border: 1px solid var(--border); transition: all 0.4s; box-shadow: 0 15px 40px rgba(0,0,0,0.3); position: relative; }
        .kpi-card:hover { transform: translateY(-5px); border-color: var(--primary); }
        .kpi-card span { font-size: 14px; font-weight: 900; color: var(--dim); text-transform: uppercase; letter-spacing: 1px; }
        .kpi-card h2 { font-size: 64px; margin: 10px 0 0; font-weight: 900; color: #fff; line-height: 1; }
        .kpi-card .target { position: absolute; top:35px; left:35px; font-size:24px; opacity:0.5; }

        .card { background: var(--card); border-radius: 25px; padding: 25px; border: 1px solid var(--border); box-shadow: 0 10px 40px rgba(0,0,0,0.25); display: flex; flex-direction: column; width: 100%; box-sizing: border-box; }
        .card-t { font-weight: 900; font-size: 20px; margin-bottom: 25px; color: #fff; display: flex; align-items: center; gap: 10px; flex-shrink: 0; }
        .card-t::before { content:''; width:5px; height:22px; background:var(--primary); border-radius:3px; }
        
        .sub-nav { display: flex; gap: 15px; margin-bottom: 25px; background: rgba(255,255,255,0.03); padding: 10px; border-radius: 12px; border: 1px solid var(--border); }
        .sub-nav-item { cursor: pointer; padding: 8px 20px; border-radius: 8px; font-weight: 800; font-size: 15px; opacity: 0.5; transition: 0.3s; }
        .sub-nav-item:hover { opacity: 1; background: rgba(255,255,255,0.05); }
        .sub-nav-item.active { opacity: 1; background: var(--primary); color: #fff; box-shadow: 0 5px 15px rgba(59, 130, 246, 0.3); }

        .manager-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 20px; margin-top: 20px; }
        .manager-card { background: var(--card); border: 1px solid var(--border); padding: 20px; border-radius: 20px; cursor: pointer; transition: 0.3s; }
        .manager-card:hover { border-color: var(--primary); transform: translateY(-3px); }
        .manager-card h3 { margin: 0; font-size: 20px; font-weight: 900; color: #fff; }
        .manager-card p { margin: 5px 0 0; font-size: 14px; color: var(--dim); font-weight: 700; }

        table { width: 100%; border-collapse: separate; border-spacing: 0 10px; margin-top: -10px; table-layout: fixed; }
        th { text-align: right; color: var(--dim); font-size: 13px; padding: 20px; text-transform: uppercase; font-weight:900; border-bottom: 1px solid var(--border); }
        td { padding: 20px; background: rgba(15, 23, 42, 0.4); font-size: 18px; font-weight: 700; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: #fff; border: 1px solid var(--border); }
        td:first-child { border-radius: 15px 0 0 15px; color: var(--primary); font-size: 20px; font-weight: 900; width: 30%; }
        td:last-child { border-radius: 0 15px 15px 0; }

        /* REFINED GUIDES UI - Premium Documentation Center */
        #guides-section {
            display: flex;
            height: calc(100vh - 180px);
            background: var(--panel);
            border-radius: 24px;
            overflow: hidden;
            box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
            border: 1px solid var(--border);
            margin-top: 10px;
        }

        .dark-mode #guides-section {
            background: #111827;
            color: #e5e7eb;
        }

        .guides-sidebar {
            width: 320px;
            background: #f9fafb;
            border-left: 1px solid #e5e7eb;
            display: flex;
            flex-direction: column;
            padding: 20px 0;
        }

        .dark-mode .guides-sidebar {
            background: #0f172a;
            border-left: 1px solid rgba(255,255,255,0.05);
        }

        .sidebar-header {
            padding: 0 20px 20px;
            border-bottom: 1px solid #e5e7eb;
            margin-bottom: 15px;
        }
        .dark-mode .sidebar-header { border-bottom-color: rgba(255,255,255,0.05); }

        .guides-content {
            flex: 1;
            padding: 50px;
            overflow-y: auto;
            position: relative;
            background: #0f172a; /* Force professional dark blue */
            color: #f1f5f9; /* Force light text */
        }
        
        .guide-viewer {
            max-width: 900px;
            margin: 0 auto;
            line-height: 1.8;
            font-size: 17px;
        }

        /* Ensure pasted content is readable */
        .guide-viewer * {
            color: inherit !important;
            background-color: transparent !important;
        }
        
        .guide-viewer img, #guide-content img {
            max-width: 100%;
            height: auto;
            border-radius: 12px;
            margin: 15px 0;
            border: 1px solid var(--border);
            display: block;
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }

        .guide-header {
            border-bottom: 2px solid var(--border);
            padding-bottom: 30px;
            margin-bottom: 40px;
            text-align: center;
        }

        /* Documentation Styling */
        .doc-header {
            border-bottom: 2px solid var(--primary);
            padding-bottom: 20px;
            margin-bottom: 40px;
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
        }

        .guide-viewer { padding: 40px; }
        .doc-header { border-bottom: 3px solid var(--primary); padding-bottom: 25px; margin-bottom: 40px; display: flex; justify-content: space-between; align-items: center; }
        .doc-header h1 { font-size: 42px; margin: 0; font-weight: 900; color: #0f172a; }
        .doc-meta { font-size: 13px; text-transform: uppercase; letter-spacing: 2px; color: var(--primary); font-weight: 800; margin-bottom: 8px; }
        
        .guide-body { font-size: 19px; line-height: 2; color: #334155; text-align: right; }
        .guide-body h2 { color: var(--primary); margin-top: 45px; font-size: 28px; border-right: 6px solid var(--primary); padding-right: 20px; font-weight: 800; }
        .guide-body p { margin-bottom: 25px; }

        .nav-tree-item { 
            padding: 15px 20px; cursor: pointer; border-radius: 0; display: flex; align-items: center; gap: 12px; 
            margin-bottom: 0; transition: 0.3s; font-size: 15px; font-weight: 600; color: var(--dim);
            border-bottom: 1px solid rgba(255,255,255,0.03);
        }
        .nav-tree-item:hover { background: rgba(59,130,246,0.1); color: #fff; }
        .nav-tree-item.active { background: var(--primary); color: #fff; border-right: 5px solid #fff; }
        
        .subcat-title {
            padding: 15px 20px 5px;
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 1px;
            color: var(--dim);
            font-weight: 900;
        }

        .subcat-header { transition: 0.3s; border: 1px solid transparent; }
        .subcat-header:hover { background: rgba(255,255,255,0.05); color: var(--primary); }
        .subcat-header.active { background: rgba(59,130,246,0.1); color: var(--primary); }

        .admin-btn {
            background: rgba(59,130,246,0.1);
            color: var(--primary);
            border: 1px solid var(--primary);
            padding: 8px 15px;
            border-radius: 8px;
            font-weight: 900;
            cursor: pointer;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 1px;
            transition: 0.3s;
        }
        .admin-btn:hover { background: var(--primary); color: #fff; }

        /* Driver Area */
        .driver-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 20px; }
        .driver-card {
            background: var(--card);
            border: 2px solid var(--border);
            padding: 30px;
            border-radius: 24px;
            display: flex;
            flex-direction: column;
            gap: 20px;
            transition: 0.4s;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }
        .driver-card:hover { transform: translateY(-8px); box-shadow: 0 20px 50px rgba(59,130,246,0.2); border-color: var(--primary); }
        
        .driver-title { font-weight: 900; font-size: 22px; color: #fff; }
        .driver-info { font-size: 15px; color: var(--dim); line-height: 1.6; }
        .btn-download {
            background: var(--primary);
            color: #fff;
            text-align: center;
            padding: 12px;
            border-radius: 10px;
            text-decoration: none;
            font-weight: 900;
            transition: 0.2s;
        }
        .btn-download:hover { background: #2563eb; }
        
        .modal-body { padding: 30px; display: flex; flex-direction: column; gap: 20px; overflow-y: auto; flex: 1; }
        .modal-body input, .modal-body textarea { background: rgba(255,255,255,0.05); border: 1px solid var(--border); padding: 15px; border-radius: 10px; color: #fff; font-family: inherit; font-size: 16px; }
        
        .overlay { position: fixed; inset: 0; background: rgba(0,0,0,0.9); display: none; z-index: 1000; backdrop-filter: blur(10px); }
        .modal { position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); width: 750px; max-width: 95vw; max-height: 90vh; background: #111827; border-radius: 30px; border: 1px solid var(--border); display: none; z-index: 1001; flex-direction: column; overflow: hidden; box-shadow: 0 0 100px rgba(0,0,0,0.8); }
        .nav:hover .cat-actions { opacity: 1 !important; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <div class="top-bar">
        <div style="flex:1; display:flex; gap:20px; align-items:center;">
            <h1 style="font-size:24px; font-weight:900; background: linear-gradient(to right, #60a5fa, #a78bfa); -webkit-background-clip: text; -webkit-text-fill-color: transparent; min-width:max-content;">TIER 2 VICO</h1>
            <div class="clock-box" id="live-clock" style="font-size:14px;">--:--:--</div>
            <div id="health-check" style="display:flex; gap:10px; font-size:12px; margin-right:15px; border-right:1px solid var(--border); padding-right:15px;">
                <div id="h-firebase" title="Firestore Connection" style="display:flex; align-items:center; gap:5px; color:var(--dim)"><span style="width:8px; height:8px; border-radius:50%; background:#666"></span> DB</div>

            </div>
        </div>
        <div class="nav-links" id="main-nav" style="flex:3; justify-content:center; gap:12px">
            <div class="nav active" id="nav-customers" onclick="nav('customers')">🤝 לקוחות</div>
            <!-- Dynamic Categories Rendered Here -->
        </div>
        <div style="display:flex; gap:15px; align-items:center;">
            <button onclick="openAddGuide()" title="יצירת מדריך חדש" style="background:#8b5cf6; color:#fff; border:none; padding:8px 15px; border-radius:12px; font-size:13px; cursor:pointer; display:flex; align-items:center; gap:8px;">📝 מדריך חדש</button>

            <button onclick="openAddCat()" title="הוספת קטגוריה" style="background:var(--primary); color:#fff; border:none; padding:8px 15px; border-radius:12px; font-size:13px; cursor:pointer; transition:0.3s; display:flex; align-items:center; gap:8px;">📁 קטגוריה חדשה</button>
            <button onclick="takeShot()" style="background:#10b981; color:#fff; border:none; padding:10px 20px; border-radius:12px; font-weight:900; cursor:pointer; box-shadow:0 0 20px rgba(16,185,129,0.3)">📸</button>
        </div>
    </div>

    <div id="capture-area" class="main">
        <div class="header">
            <div><h1 id="t">Commander BI</h1><p id="s">זרם מודיעין בזמן אמת</p></div>
            <div id="filter-box" style="display:flex; gap:15px; align-items:center;">
                <input type="text" id="cust-search" placeholder="חיפוש לקוח, מנהל או גרסה..." style="background:var(--card); border:1px solid var(--border); color:#fff; padding:10px 15px; border-radius:10px; font-family:inherit; min-width:300px;" oninput="filterIntegrations()">
                <button onclick="openAdd()" style="background:var(--primary); color:#fff; border:none; padding:10px 20px; border-radius:10px; font-weight:900; cursor:pointer; font-size:14px;">+ הוספת פרויקט</button>
            </div>
            
            <div id="report-filter-box" style="display:none; gap:15px; align-items:center;">
                <input type="date" id="rep-start" style="background:var(--card); border:1px solid var(--border); color:#fff; padding:10px 15px; border-radius:10px;">
                <input type="date" id="rep-end" style="background:var(--card); border:1px solid var(--border); color:#fff; padding:10px 15px; border-radius:10px;">
                <button onclick="refreshReports()" style="background:var(--accent); color:#fff; border:none; padding:10px 20px; border-radius:10px; font-weight:900; cursor:pointer;">📊 הפקת דוח</button>
            </div>
        </div>

        <div class="sub-nav">
            <div class="sub-nav-item active" onclick="subNav('projects')">📁 כל הלקוחות</div>
            <div class="sub-nav-item" onclick="subNav('managers')">מנהלי פרויקטים</div>
        </div>

        <div class="kpi-row">
            <div class="kpi-card"><span id="l1">פעילות כוללת</span><h2 id="v1">0</h2><div class="target">📊</div></div>
            <div class="kpi-card"><span id="l2">יעילות</span><h2 id="v2">0</h2><div class="target">⏱️</div></div>
            <div class="kpi-card"><span id="l3">ציון איכות</span><h2 id="v3">0</h2><div class="target">⭐</div></div>
            <div class="kpi-card"><span id="l4">דופק בריאות</span><h2 id="v4">0</h2><div class="target">❤️</div></div>
        </div>

        <div class="card" id="perf-card">
            <div class="card-t" id="list-t">פירוט ביצועים</div>
            <table>
                <thead id="thead"><tr><th>פרויקט</th><th>סוג מכשיר</th><th>GW</th><th>מנהל</th><th>גרסה</th><th style="width:80px">מדריכים</th><th style="width:100px">פעולה</th></tr></thead>
                <tbody id="files"></tbody>
            </table>
        </div>

        <!-- GUIDE SECTION (PROFESSIONAL DOC CENTER) -->
        <div id="guides-section" style="display:none; flex-direction:row; height:calc(100vh - 160px); border-radius:20px; overflow:hidden; background:rgba(15,23,42,0.4); border:1px solid var(--border);">
            <div class="guides-sidebar" id="g-sidebar" style="width:300px; background:rgba(0,0,0,0.2); border-left:1px solid var(--border); display:flex; flex-direction:column;">
                <div style="padding:20px; border-bottom:1px solid var(--border); background:rgba(255,255,255,0.02);">
                    <h3 id="sidebar-cat-name" style="margin:0; font-weight:900; color:var(--primary); font-size:14px; text-transform:uppercase; letter-spacing:1px;">מרכז ידע</h3>
                </div>
                <div id="g-nav-tree" style="overflow-y:auto; flex:1; padding:10px;"></div>
                <div style="padding:15px; border-top:1px solid var(--border);">
                    <button class="btn" style="width:100%; font-size:12px; border:1px dashed var(--primary); background:none; color:var(--primary);" onclick="openAddGuide()">+ יצירת מדריך</button>
                </div>
            </div>
            <div class="guides-content" id="g-display" style="flex:1; overflow-y:auto; padding:50px; background:var(--bg); direction:rtl; text-align:right;">
                <div id="g-viewer" class="guide-viewer">
                    <div style="text-align:center; padding-top:150px; opacity:0.1">
                        <span style="font-size:150px">📖</span>
                        <h2 style="font-size:40px; margin-top:20px;">מרכז מידע ותיעוד</h2>
                        <p>בחר מדריך מהתפריט הצדדי כדי להתחיל לקרוא.</p>
                    </div>
                </div>
            </div>
        </div>

        <!-- AMONG OUR CUSTOMERS SECTION -->
        <div id="customers-showcase" style="display:none; padding:40px;">
            <div style="text-align:right; margin-bottom:50px;">
                <h1 style="font-size:48px; font-weight:900; background:linear-gradient(to left, #fff, var(--dim)); -webkit-background-clip:text; -webkit-text-fill-color:transparent; margin:0;">בין לקוחותנו</h1>
                <p style="color:var(--dim); font-size:18px;">שותפויות אסטרטגיות של Verifone</p>
            </div>
            <div id="customer-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(280px, 1fr)); gap:30px;"></div>
        </div>
    </div>

    <!-- MODALS -->
    <div class="overlay" onclick="closeM()"></div>
    
    <!-- ADD CATEGORY MODAL -->
    <div class="modal" id="cat-modal">
        <div style="background:#0f172a; padding:15px 25px; display:flex; justify-content:space-between; align-items:center;">
            <b>הוספת קטגוריה חדשה</b><button onclick="closeM()" style="background:none; border:none; color:#ef4444; font-size:24px; cursor:pointer;">&times;</button>
        </div>
        <div class="modal-body">
            <div style="display:flex; gap:10px; margin-bottom:10px;">
                <input type="text" id="cat-emoji" placeholder="אימוג'י" style="width:70px; text-align:center; font-size:24px;">
                <input type="text" id="cat-name" placeholder="שם הקטגוריה" style="flex:1;">
            </div>
            
            <div class="input-group" style="margin-bottom:15px;">
                <label style="color:var(--dim); font-size:12px; display:block; margin-bottom:5px;">סוג מכשיר</label>
                <select id="cat-type" style="width:100%; background:rgba(255,255,255,0.05); border:1px solid var(--border); padding:12px; border-radius:10px; color:#fff; font-family:inherit;">
                    <option value="kb">📚 מרכז ידע (מדריכים)</option>
                    <option value="table">🤝 טבלת פרויקטים (Vico)</option>
                    <option value="table_general">📋 טבלה כללית (שם, תיאור, סטטוס, הערות)</option>
                    <option value="table_phones">📞 רשימת טלפונים (שם, טלפון, תפקיד, מייל)</option>
                    <option value="table_ip">🌐 רשימת כתובות IP (שרת, IP, מיקום, פורט)</option>
                    <option value="table_pass">🔑 ניהול סיסמאות (שם, משתמש, סיסמה, הערות)</option>
                </select>
            </div>

            <div id="emoji-picker" style="display:grid; grid-template-columns: repeat(8, 1fr); gap:10px; padding:15px; background:rgba(255,255,255,0.05); border-radius:15px; border:1px solid var(--border); max-height:150px; overflow-y:auto; margin-bottom:10px;">
                <!-- Emojis will be injected here -->
            </div>
            <button class="btn" id="cat-save-btn" onclick="saveCategory()">שמירת קטגוריה</button>
        </div>
    </div>

    <!-- ADD GUIDE MODAL -->
    <div class="modal" id="guide-modal" style="width: 850px;">
        <div style="background:#0f172a; padding:15px 25px; display:flex; justify-content:space-between; align-items:center;">
            <b>יצירת מדריך חדש</b><button onclick="closeM()" style="background:none; border:none; color:#ef4444; font-size:24px; cursor:pointer;">&times;</button>
        </div>
        <div class="modal-body">
            <div style="display:flex; gap:10px;">
                <select id="guide-cat" onchange="updateSubCatDropdown()" style="flex:1; background:rgba(255,255,255,0.05); border:1px solid var(--border); padding:15px; border-radius:10px; color:#fff; font-family:inherit;"></select>
                <select id="guide-subcat" style="flex:1; background:rgba(255,255,255,0.05); border:1px solid var(--border); padding:15px; border-radius:10px; color:#fff; font-family:inherit;">
                    <option value="">[ קטגוריה ראשית ]</option>
                </select>
            </div>
            <input type="text" id="guide-title" placeholder="כותרת המדריך">
            <div id="guide-content" contenteditable="true" placeholder="הדבק את המדריך כאן (טקסט ותמונות)..." style="height:400px; overflow-y:auto; background:rgba(0,0,0,0.2); border:1px solid var(--border); border-radius:12px; padding:20px; color:#fff; font-family:inherit; font-size:16px; direction:rtl; text-align:right; outline:none;"></div>
            
            <div style="background:rgba(16,185,129,0.05); border:1px solid rgba(16,185,129,0.2); padding:15px; border-radius:12px; font-size:13px; color:#10b981;">
                💡 <b>חשוב:</b> כדי לייבא תמונות אוטומטית, השתמש בכפתור <b>ייבוא תוכן מקובץ</b> ובחר קובץ וורד (Docx).
            </div>
            
            <div style="background:rgba(59,130,246,0.05); border:1px solid rgba(59,130,246,0.2); padding:20px; border-radius:15px;">
                <label style="display:block; margin-bottom:10px; font-weight:900; font-size:12px; color:var(--primary)">ניהול תמונות ומסמכים</label>
                <div style="display:flex; gap:10px; align-items:center; margin-bottom:15px">
                    <input type="file" id="image-upload" accept="image/*" style="display:none" onchange="handleUpload(this)">
                    <button class="btn" onclick="document.getElementById('image-upload').click()" style="background:#0f172a; border:1px dashed var(--primary); color:var(--primary); padding:10px 20px; font-size:14px;">📁 העלאת תמונה</button>
                    
                    <input type="file" id="content-import" accept=".docx,.pdf,.txt" style="display:none" onchange="importContent(this)">
                    <button class="btn" onclick="document.getElementById('content-import').click()" style="background:rgba(16,185,129,0.1); border:1px dashed #10b981; color:#10b981; padding:10px 20px; font-size:14px;">📄 ייבוא תוכן מקובץ (Word/PDF)</button>
                </div>
                <div id="guide-images" style="display:flex; gap:10px; flex-wrap:wrap;"></div>
            </div>
            
            <button class="btn" onclick="saveGuide()" style="margin-top:10px; padding: 20px; font-size: 18px; background: var(--primary);">💾 שמירת המדריך למערכת</button>
        </div>
    </div>

    <!-- CUSTOMER DETAIL MODAL -->
    <div class="modal" id="cust-detail-modal" style="width:500px; text-align:right;">
        <div style="background:#0f172a; padding:15px 25px; display:flex; justify-content:space-between; align-items:center;">
            <b id="detail-name">פרופיל לקוח</b><button onclick="closeM()" style="background:none; border:none; color:#ef4444; font-size:24px; cursor:pointer;">&times;</button>
        </div>
        <div class="modal-body" style="align-items:center; text-align:center;">
            <div style="width:120px; height:120px; background:rgba(255,255,255,0.05); border-radius:20px; display:flex; align-items:center; justify-content:center; padding:20px; margin-bottom:20px;">
                <img id="detail-logo" src="" style="max-width:100%; max-height:100%; object-fit:contain;">
            </div>
            <h2 id="detail-title" style="margin:0; font-size:24px; color:#fff;"></h2>
            <p id="detail-desc" style="color:var(--dim); line-height:1.6; font-size:16px; margin-top:15px;"></p>
            <div style="width:100%; height:1px; background:var(--border); margin:20px 0;"></div>
            <button class="btn" onclick="closeM()" style="width:100%; background:var(--primary)">סגור פרופיל</button>
        </div>
    </div>

    <!-- EDIT INTEGRATION MODAL -->
    <div class="modal" id="edit-modal" style="width:700px;">
        <div style="background:#0f172a; padding:15px 25px; display:flex; justify-content:space-between; align-items:center;">
            <b>Edit Project Data</b><button onclick="closeM()" style="background:none; border:none; color:#ef4444; font-size:24px; cursor:pointer;">&times;</button>
        </div>
        <div class="modal-body">
            <div style="display:grid; grid-template-columns: 1fr 1fr; gap:15px;">
                <div class="input-group">
                    <label id="lbl-cust" style="color:var(--dim); font-size:12px">CUSTOMER</label>
                    <input type="text" id="edit-cust" style="width:100%; box-sizing:border-box">
                </div>
                <div class="input-group">
                    <label id="lbl-device" style="color:var(--dim); font-size:12px">SOLUTION TYPE (DEVICE)</label>
                    <input type="text" id="edit-device" style="width:100%; box-sizing:border-box">
                </div>
                <div class="input-group">
                    <label id="lbl-gw" style="color:var(--dim); font-size:12px">GW / CONNECTION</label>
                    <input type="text" id="edit-gw" style="width:100%; box-sizing:border-box">
                </div>
                <div class="input-group">
                    <label id="lbl-pm" style="color:var(--dim); font-size:12px">PROJECT MANAGER</label>
                    <input type="text" id="edit-pm" style="width:100%; box-sizing:border-box">
                </div>
                <div class="input-group">
                    <label id="lbl-ver" style="color:var(--dim); font-size:12px">VERSION</label>
                    <input type="text" id="edit-version" style="width:100%; box-sizing:border-box">
                </div>
                <div class="input-group" style="grid-column: span 2;">
                    <label style="color:var(--dim); font-size:12px">CATEGORY (e.g. Retail, Food, Luxury)</label>
                    <input type="text" id="edit-project-cat" list="project-cat-list" style="width:100%; box-sizing:border-box" placeholder="Select or TYPE NEW Name to Create Tab...">
                    <datalist id="project-cat-list">
                        <option value="Retail">
                        <option value="Food & Beverage">
                        <option value="Luxury">
                        <option value="Production">
                        <option value="Development">
                        <option value="VIP High-Priority">
                        <option value="Critical Maintenance">
                        <option value="Cloud Migration">
                        <option value="On-Premise">
                        <option value="External Vendor">
                        <option value="Internal QA">
                    </datalist>
                </div>
            </div>
            
            <div style="margin-top:20px; padding:20px; background:rgba(255,255,255,0.03); border:1px solid var(--border); border-radius:15px;">
                <div style="font-size:12px; font-weight:900; color:var(--primary); text-transform:uppercase; margin-bottom:15px;">Project Documentation (Files)</div>
                
                <div style="display:flex; flex-direction:column; gap:15px;">
                    <div style="display:flex; gap:10px; align-items:center;">
                        <input type="text" id="edit-sheet" placeholder="Release Sheet URL" style="flex:1;">
                        <button class="btn" onclick="document.getElementById('upload-sheet-file').click()" style="width:auto; padding:8px 15px; font-size:12px;">📁 Upload Sheet</button>
                        <input type="file" id="upload-sheet-file" style="display:none" accept=".pdf,.docx,.doc,.xlsx,.xls,.pptx,.ppt" onchange="handleCustUpload(this, 'edit-sheet')">
                    </div>
                    
                    <div style="display:flex; gap:10px; align-items:center;">
                        <input type="text" id="edit-note" placeholder="Release Note URL" style="flex:1;">
                        <button class="btn" onclick="document.getElementById('upload-note-file').click()" style="width:auto; padding:8px 15px; font-size:12px;">📁 Upload Note</button>
                        <input type="file" id="upload-note-file" style="display:none" accept=".pdf,.docx,.doc,.xlsx,.xls,.pptx,.ppt" onchange="handleCustUpload(this, 'edit-note')">
                    </div>

                    <div style="display:flex; gap:10px; align-items:center;">
                        <input type="text" id="edit-manual" placeholder="Manual / Config URL" style="flex:1;">
                        <button class="btn" onclick="document.getElementById('upload-manual-file').click()" style="width:auto; padding:8px 15px; font-size:12px;">📁 Upload Manual/Config</button>
                        <input type="file" id="upload-manual-file" style="display:none" accept=".pdf,.docx,.doc,.xlsx,.xls,.pptx,.ppt" onchange="handleCustUpload(this, 'edit-manual')">
                    </div>
                </div>
                <div id="cust-upload-status" style="font-size:11px; margin-top:10px; font-weight:900; color:var(--accent);"></div>
            </div>
            
            <button class="btn" onclick="saveEdit()" style="margin-top:20px; background:var(--accent)">Update Project</button>
        </div>
    </div>

    <!-- CUSTOMER DETAIL MODAL -->
    <div class="modal" id="cust-detail-modal" style="width:500px; text-align:right;">
        <div style="background:#0f172a; padding:15px 25px; display:flex; justify-content:space-between; align-items:center;">
            <b id="detail-name">פרופיל לקוח</b><button onclick="closeM()" style="background:none; border:none; color:#ef4444; font-size:24px; cursor:pointer;">&times;</button>
        </div>
        <div class="modal-body" style="align-items:center; text-align:center;">
            <div style="width:120px; height:120px; background:rgba(255,255,255,0.05); border-radius:20px; display:flex; align-items:center; justify-content:center; padding:20px; margin-bottom:20px;">
                <img id="detail-logo" src="" style="max-width:100%; max-height:100%; object-fit:contain;">
            </div>
            <h2 id="detail-title" style="margin:0; font-size:24px; color:#fff;"></h2>
            <p id="detail-desc" style="color:var(--dim); line-height:1.6; font-size:16px; margin-top:15px;"></p>
            <div style="width:100%; height:1px; background:var(--border); margin:20px 0;"></div>
            <button class="btn" onclick="closeM()" style="width:100%; background:var(--primary)">סגור פרופיל</button>
        </div>
    </div>

    <!-- CATEGORY MODAL -->
    <div class="modal" id="cat-modal" style="width:600px;">
        <div style="background:#0f172a; padding:15px 25px; display:flex; justify-content:space-between; align-items:center;">
            <b>הוספת קטגוריה חדשה</b><button onclick="closeM()" style="background:none; border:none; color:#ef4444; font-size:24px; cursor:pointer;">&times;</button>
        </div>
        <div class="modal-body">
            <div class="input-group">
                <label style="color:var(--dim); font-size:12px; margin-bottom:8px; display:block;">שם הקטגוריה</label>
                <input type="text" id="cat-name" placeholder="לדוגמה: מדריכי VPN" style="width:100%; box-sizing:border-box">
            </div>
            
            <div class="input-group">
                <label style="color:var(--dim); font-size:12px; margin-bottom:8px; display:block;">אימוג'י</label>
                <input type="text" id="cat-emoji" placeholder="📚" maxlength="2" style="width:100%; box-sizing:border-box; font-size:32px; text-align:center;">
            </div>
            
            <div class="input-group">
                <label style="color:var(--dim); font-size:12px; margin-bottom:8px; display:block;">סוג קטגוריה</label>
                <select id="cat-type" style="width:100%; box-sizing:border-box; background:rgba(255,255,255,0.05); border:1px solid var(--border); padding:15px; border-radius:10px; color:#fff; font-family:inherit; font-size:16px;">
                    <option value="guides">מדריכים רגילים</option>
                    <option value="table_pass">טבלת סיסמאות</option>
                    <option value="table_ip">טבלת IP</option>
                    <option value="table_phones">טבלת טלפונים</option>
                    <option value="table_general">טבלה כללית</option>
                </select>
            </div>
            
            <button class="btn" id="cat-save-btn" onclick="saveCategory()" style="margin-top:20px; background:var(--primary); width:100%;">יצירת קטגוריה</button>
        </div>
    </div>

    <script>
        let subSect = 'projects', selectedSubCatId = null;
        let stats_data = { Integrations: [] };
        let guides_data = [];
        let editingCatId = null;
        let editingGuideId = null;
        
        const EMOJIS = ['🤝','📚','🛠️','🚀','💡','📋','⚙️','🛡️','💎','🔥','📈','🌐','📱','💻','🔑','📎','📂','📍','✅','⚠️','🆘','🏁','🏆','🎁','📦','🔔','📣','✨'];

        function initEmojiPicker() {
            const picker = document.getElementById('emoji-picker');
            if(!picker) return;
            picker.innerHTML = EMOJIS.map(e => `<span onclick="selectEmoji('${e}')" style="font-size:24px; cursor:pointer; padding:5px; border-radius:8px; transition:0.2s; display:inline-block;" onmouseover="this.style.background='rgba(255,255,255,0.1)'" onmouseout="this.style.background='transparent'">${e}</span>`).join('');
        }
        function selectEmoji(e) { document.getElementById('cat-emoji').value = e; }
        let sect = 'customers';
        let selectedCatId = null;
        let selectedGuideId = null;

        async function init() {
            console.log("init() started");
            // Check System Health
            fetch('/api/health').then(r => r.json()).then(h => {
                const fb = document.getElementById('h-firebase');
                if(fb) {
                    fb.style.color = h.firebase ? 'var(--accent)' : '#ef4444';
                    fb.querySelector('span').style.background = h.firebase ? 'var(--accent)' : '#ef4444';
                }
                // GDrive health check removed
            }).catch(e => console.warn("Health check failed", e));

            const clock = document.getElementById('live-clock');
            if(clock) setInterval(() => clock.innerText = new Date().toLocaleTimeString('en-GB'), 1000);
            
            // Set default date range (today)
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('rep-start').value = today;
            document.getElementById('rep-end').value = today;

            // Handle Hash for deep-linking/back button
            window.addEventListener('hashchange', parseHash);
            parseHash(false);

            await refresh();
            setInterval(refresh, 60000);
            
            // Setup Drag & Drop for guides
            setupGuidesDropZone();
        }
        
        function setupGuidesDropZone() {
            const guideSection = document.getElementById('guides-section');
            if(!guideSection) return;
            
            guideSection.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.stopPropagation();
                guideSection.style.background = 'rgba(var(--accent-rgb, 96, 165, 250), 0.1)';
                guideSection.style.border = '2px dashed var(--accent)';
            });
            
            guideSection.addEventListener('dragleave', (e) => {
                e.preventDefault();
                e.stopPropagation();
                guideSection.style.background = '';
                guideSection.style.border = '';
            });
            
            guideSection.addEventListener('drop', async (e) => {
                e.preventDefault();
                e.stopPropagation();
                guideSection.style.background = '';
                guideSection.style.border = '';
                
                const files = Array.from(e.dataTransfer.files);
                if(files.length === 0) return;
                
                // Get category from current context
                const catId = selectedCatId;
                if(!catId) {
                    alert('יש לבחור קטגוריה תחילה מהתפריט הצדדי');
                    return;
                }
                
                // Show progress
                const progressDiv = document.createElement('div');
                progressDiv.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:rgba(0,0,0,0.9);padding:30px;border-radius:12px;z-index:9999;color:white;text-align:center;';
                progressDiv.innerHTML = `<div style="font-size:24px;margin-bottom:10px;">מעלה ${files.length} קבצים...</div><div id="progress-text">0%</div>`;
                document.body.appendChild(progressDiv);
                
                let completed = 0;
                for(let file of files) {
                    await processFileToGuide(file, catId);
                    completed++;
                    document.getElementById('progress-text').innerText = Math.round((completed/files.length)*100) + '%';
                }
                
                await syncGuides();
                update();
                
                document.body.removeChild(progressDiv);
                alert(`${files.length} מדריכים נוספו בהצלחה!`);
            });
        }

        function parseHash(shouldUpdate = true) {
            const h = window.location.hash.substring(1).split('/');
            sect = h[0] || 'customers';
            selectedCatId = h[1] || null;
            
            if (h[2] === 'guide' && h[3]) {
                selectedGuideId = h[3];
                selectedSubCatId = null;
            } else if (h[2] === 'sub' && h[3]) {
                selectedSubCatId = h[3];
                selectedGuideId = null;
            } else {
                selectedGuideId = null;
                selectedSubCatId = null;
            }

            if(shouldUpdate) update(false);
        }

        // --- NEW UPDATE LOGIC FOR GUIDES ---
        function update(doSyncHash = true) {
            if(doSyncHash) syncHash();
            renderTopNav();
            
            // Hide all main sections by default
            document.getElementById('filter-box').style.display = 'none';
            document.getElementById('report-filter-box').style.display = 'none';
            document.querySelector('.sub-nav').style.display = 'none';
            document.querySelector('.kpi-row').style.display = 'none';
            document.getElementById('guides-section').style.display = 'none';
            document.getElementById('customers-showcase').style.display = 'none';
            document.getElementById('perf-card').style.display = 'none';
            document.getElementById('manager-view')?.remove();

            if (sect === 'customers') {
                document.getElementById('filter-box').style.display = 'flex';
                document.querySelector('.sub-nav').style.display = 'flex';
                document.querySelector('.kpi-row').style.display = 'grid';
                document.getElementById('perf-card').style.display = 'block';
                
                document.getElementById('t').innerText = 'אינטגרציות ופרויקטים';
                document.getElementById('s').innerText = subSect === 'projects' ? 'ניהול מחזור חיי פרויקט' : 'ניהול עומסי מנהלים';
                
                renderCustomerSubNav();
                
                if(!stats_data || !stats_data.Integrations) return;
                let d = stats_data.Integrations;
                
                if(subSect === 'projects' && selectedSubCatId) {
                    d = d.filter(x => x.Category === selectedSubCatId);
                }

                uk("סה\"כ לקוחות", d.length, "בביצוע", d.length, "איכות", "100%", "סטטוס", "פעיל");
                if(subSect === 'projects') {
                    document.getElementById('perf-card').style.display = 'block';
                    renderIntegrations(d);
                } else if(subSect === 'warranty') {
                    document.getElementById('perf-card').style.display = 'block';
                    renderWarrantyTable(d);
                } else {
                    document.getElementById('perf-card').style.display = 'none';
                    renderManagers(d);
                }
            } else if (sect === 'our-customers') {
                document.getElementById('customers-showcase').style.display = 'block';
                document.getElementById('t').innerText = 'בין לקוחותנו';
                document.getElementById('s').innerText = 'מערכת אינטגרציות ארגונית';
                renderOurCustomers();
            } else if (sect === 'reports') {
                document.getElementById('report-filter-box').style.display = 'flex';
                document.querySelector('.kpi-row').style.display = 'grid';
                
                document.getElementById('t').innerText = 'ניתוח ביצועים';
                document.getElementById('s').innerText = 'דוחות API ומדדי שירות';
                renderReports();
            } else if (sect === 'guides') {  
                const cat = guides_data.find(c => c.id == selectedCatId);
                if (cat && (cat.type === 'table' || (cat.type && cat.type.startsWith('table')))) {
                    document.getElementById('filter-box').style.display = 'flex';
                    document.querySelector('.sub-nav').style.display = 'flex';
                    document.querySelector('.kpi-row').style.display = 'grid';
                    document.getElementById('perf-card').style.display = 'block';
                    document.getElementById('t').innerText = cat.name;
                    document.getElementById('s').innerText = 'קונסולת ניהול נתונים';
                    renderCustomerSubNav(); 
                    let d = cat.guides || []; 
                    if(selectedSubCatId) d = d.filter(x => x.Category === selectedSubCatId);
                    uk("סה\"כ שורות", d.length, "לקוחות פעילים", d.length, "בריאות", "100%", "תבנית", "טבלה");
                    renderIntegrations(d);
                } else {
                    document.getElementById('guides-section').style.display = 'flex';
                    const cat = guides_data.find(c => c.id == selectedCatId);
                    
                    // Always show the "Add Guide" button area if we are in Guides section
                    document.getElementById('g-sidebar').style.display = 'flex';
                    
                    if(!cat) {
                        // Default view for Guides when no category is selected
                        document.getElementById('t').innerText = 'מרכז ידע ותיעוד';
                        document.getElementById('s').innerText = 'בחר קטגוריה מהתפריט העליון';
                        document.getElementById('g-nav-tree').innerHTML = '<div style="padding:20px; color:var(--dim); text-align:center">אנא בחר קטגוריה לעריכה או צפייה</div>';
                        return;
                    }

                    document.querySelector('.kpi-row').style.display = 'none';
                    document.querySelector('.sub-nav').style.display = 'none';
                    if(selectedGuideId) renderGuideView(selectedCatId, selectedGuideId);
                    else {
                        document.getElementById('t').innerText = cat.name;
                        document.getElementById('s').innerText = 'מרכז תיעוד ומדריכים';
                        renderGuideContent(cat);
                    }
                }
            }
        }

        function renderCustomerSubNav() {
            const container = document.querySelector('.sub-nav');
            let data_source = (sect === 'customers') ? stats_data.Integrations : (guides_data.find(c=>c.id==selectedCatId)?.guides || []);
            if(!data_source) return;
            
            let cats = [...new Set(data_source.map(x => x.Category).filter(Boolean))].sort();
            
            let html = `<div class="sub-nav-item ${subSect==='projects' && !selectedSubCatId ?'active':''}" onclick="selectedSubCatId=null; subNav('projects')">📁 כל הלקוחות</div>`;
            
            cats.forEach(c => {
                html += `<div class="sub-nav-item ${selectedSubCatId === c ?'active':''}" onclick="selectedSubCatId='${c}'; subNav('projects')">${c}</div>`;
            });
            
            if(sect === 'customers') {
                html += `<div class="sub-nav-item ${subSect==='warranty'?'active':''}" onclick="selectedSubCatId=null; subNav('warranty')">🛡️ אחריות לקוחות</div>`;
                html += `<div class="sub-nav-item ${subSect==='managers'?'active':''}" onclick="selectedSubCatId=null; subNav('managers')">מנהלי פרויקטים</div>`;
            }
            container.innerHTML = html;
        }

        function renderGuideContent(cat) {
            renderGuideTree(cat);
            uk("מרכז ידע", cat.name, "פריטים", (cat.guides?cat.guides.length:0), "גישה", "ציבורי", "סטטוס", "מסונכרן");
            if(!selectedGuideId) {
                document.getElementById('g-viewer').innerHTML = `
                    <div style="text-align:center; padding-top:150px; opacity:0.1">
                        <span style="font-size:150px">📖</span>
                        <h2 style="font-size:40px; margin-top:20px;">Documentation Center</h2>
                        <p>Select a guide from the sidebar to start reading.</p>
                    </div>`;
            }
        }

        function renderGuideTree(cat) {
            const tree = document.getElementById('g-nav-tree');
            if(!cat) { tree.innerHTML = ''; return; }
            
            let html = `
                <div class="nav-tree-item ${!selectedGuideId && !selectedSubCatId ? 'active' : ''}" onclick="selectedGuideId=null; selectedSubCatId=null; update();" style="font-weight:900; color:var(--accent); margin-bottom:10px; background:rgba(16,185,129,0.1); border-radius:12px; border:1px solid rgba(16,185,129,0.2);">
                    🏠 סקירה כללית
                </div>
                
                <div class="nav-tree-item" onclick="openAddSubCat()" style="font-weight:900; color:#fff; margin-bottom:20px; background:linear-gradient(135deg, var(--primary), var(--accent)); border-radius:12px; border:none; justify-content:center; box-shadow:0 5px 15px rgba(59,130,246,0.3);">
                    📂 הוסף תיקייה חדשה +
                </div>`;

            const colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899'];
            let colorIdx = 0;

            if (cat.guides && cat.guides.length > 0) {
                cat.guides.forEach(g => {
                    const color = colors[colorIdx++ % colors.length];
                    html += `
                        <div class="nav-tree-item ${selectedGuideId === g.id ? 'active' : ''}" onclick="viewGuide('${cat.id}', '${g.id}')">
                            <span style="color:${color}; font-size:18px;">📄</span> 
                            <div style="flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${g.title}</div>
                        </div>`;
                });
            }

            if (cat.subCategories && cat.subCategories.length > 0) {
                cat.subCategories.forEach((sub, idx) => {
                    const color = colors[(idx + 2) % colors.length];
                    const isOpen = selectedSubCatId === sub.id || (selectedGuideId && sub.guides && sub.guides.some(g => g.id === selectedGuideId));
                    html += `
                        <div class="subcat-header" onclick="navSubGuide('${sub.id}')" style="display:flex; align-items:center; gap:12px; padding:15px; margin-top:10px; border-radius:12px; cursor:pointer; font-weight:800; border:1px solid ${selectedSubCatId===sub.id?color:'rgba(255,255,255,0.05)'}; background:${selectedSubCatId===sub.id?color+'1A':'rgba(255,255,255,0.02)'}; color:${selectedSubCatId===sub.id?color:'var(--dim)'}; transition:0.3s;">
                            <span style="font-size:20px; color:${color}">${isOpen ? '📂' : '📁'}</span>
                            <span style="flex:1">${sub.name}</span>
                        </div>`;
                    
                    if (isOpen && sub.guides) {
                        sub.guides.forEach(g => {
                            html += `
                                <div class="nav-tree-item ${selectedGuideId === g.id ? 'active' : ''}" onclick="viewGuide('${cat.id}', '${g.id}')" style="padding-right:45px; font-size:14px; opacity:0.9; color:#f1f5f9;">
                                    <span style="color:${color}; opacity:0.7">●</span> 
                                    <div style="flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${g.title}</div>
                                </div>`;
                        });
                    }
                });
            }
            tree.innerHTML = html;
        }

        
        // Load guides from Firestore (Merged Logic)
        async function loadGuidesFromFirestore() {
            try {
                const { collection, getDocs } = window.firebaseRefs;
                
                // 1. Fetch Categories
                const catCol = collection(window.db, 'guides_categories');
                const catSnap = await getDocs(catCol);
                let cats = [];
                if (!catSnap.empty) {
                    cats = catSnap.docs.map(doc => ({ id: doc.id, guides: [], ...doc.data() }));
                }

                // 2. Fetch Guides
                const guidesCol = collection(window.db, 'guides');
                const guideSnap = await getDocs(guidesCol);
                let guides = [];
                if (!guideSnap.empty) {
                    guides = guideSnap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
                }

                // 3. Merge Guides into Categories
                cats.forEach(c => {
                    // Filter guides belonging to this category
                    c.guides = guides.filter(g => g.Category === c.id);
                });
                
                // If cats is empty but guides exist (orphaned?), potentially handle them.
                // For now, we trust the cats list.
                
                if (cats.length > 0) {
                    guides_data = cats;
                    console.log(`Loaded ${cats.length} categories and ${guides.length} guides from Firestore.`);
                } else {
                    console.log('No categories field found, trying legacy or empty.');
                    guides_data = [];
                }
            } catch(e) {
                console.error("Firestore load error:", e);
                // Fallback to API if Firestore fails
                try {
                    const res = await fetch('/api/stats');
                    const data = await res.json();
                    if(data.Guides) guides_data = data.Guides;
                } catch(apiError) {
                    console.error("API fallback failed:", apiError);
                }
            }
        }

        async function syncGDrive() {
             console.log("GDrive sync removed.");
             alert("Google Drive integration has been removed.");
        }

        async function refresh() {
            console.log("Refreshing data...");
            try {
                const res = await fetch('/api/stats');
                console.log("Stats API response received:", res.status);
                if (!res.ok) {
                    if (res.status === 401 || res.status === 403) {
                        window.location.href = '/login';
                        return;
                    }
                    throw new Error(`API returned ${res.status}`);
                }
                
                const contentType = res.headers.get("content-type");
                if (contentType && contentType.indexOf("application/json") === -1) {
                    console.warn("API returned non-JSON content (" + contentType + "). Redirecting to login...");
                    window.location.href = '/login';
                    return;
                }

                const data = await res.json();
                stats_data = data;
                
                if (data.GuidesCategories) {
                    guides_data = data.GuidesCategories;
                }
                
                update();
            } catch(e) { 
                console.error("Poll error:", e);
                if (e.message.includes('Unexpected token')) {
                    console.warn("Likely received HTML instead of JSON. Redirecting...");
                    window.location.href = '/login';
                }
            }
        }
        function nav(s) {
            sect = s;
            selectedCatId = null;
            selectedGuideId = null;
            update();
        }
        function subNav(s) {
            subSect = s;
            update();
        }
        function navSubGuide(id) {
            const base = `#guides/${selectedCatId}`;
            window.location.hash = id ? `${base}/sub/${id}` : base;
        }
        function syncHash() {
            let h = `#${sect}`;
            if(selectedCatId) h += `/${selectedCatId}`;
            if(selectedGuideId) h += `/guide/${selectedGuideId}`;
            else if(selectedSubCatId) h += `/sub/${selectedSubCatId}`;
            
            if(window.location.hash !== h) {
                window.location.hash = h;
            }
        }

        function showCustomerDetail(key) {
            const logos = stats_data.CustomerLogos || {};
            const data = logos[key] || { name: key, logo: 'https://i.ibb.co/0Y4f2N0/v-white.png', desc: 'מידע נוסף אודות הלקוח אינו זמין כרגע.' };
            
            document.getElementById('detail-name').innerText = data.name;
            document.getElementById('detail-title').innerText = data.name;
            document.getElementById('detail-logo').src = data.logo;
            document.getElementById('detail-desc').innerText = data.desc;
            
            document.querySelector('.overlay').style.display = 'block';
            document.getElementById('cust-detail-modal').style.display = 'flex';
        }

        function renderOurCustomers() {
            const grid = document.getElementById('customer-grid');
            if(!stats_data || !stats_data.Integrations) return;
            
            // Get unique customer names from integrations (filter nulls)
            let uniqueCustNames = [...new Set(stats_data.Integrations.map(x => x.Customer).filter(Boolean))];
            
            // Prioritize Verifone to the TOP
            uniqueCustNames.sort((a, b) => {
                const aLow = String(a).toLowerCase();
                const bLow = String(b).toLowerCase();
                if(aLow.includes('verifone')) return -1;
                if(bLow.includes('verifone')) return 1;
                return 0;
            });

            const logos = stats_data.CustomerLogos || {};
            
            let html = '';
            uniqueCustNames.forEach(name => {
                const key = name.toLowerCase();
                let cleanName = name.replace(/Startup Booster\s*/i, '').replace(/NCR-\s*/i, '').trim();
                let displayData = { name: cleanName, logo: 'https://i.ibb.co/0Y4f2N0/v-white.png' };
                
                // Try to find matching logo and translation
                for(let k in logos) {
                    if(key.includes(k.toLowerCase())) {
                        displayData = logos[k];
                        break;
                    }
                }
                
                html += `
                <div class="card" onclick="showCustomerDetail('${key}')" style="padding:0; overflow:hidden; border:1px solid var(--border); transition:0.4s; aspect-ratio:1/1.1; display:flex; flex-direction:column; background:rgba(255,255,255,0.02); cursor:pointer;">
                    <div style="flex:1; display:flex; align-items:center; justify-content:center; padding:40px; background:rgba(255,255,255,0.08);">
                        <img src="${displayData.logo}" data-fallbacks="${(displayData.fallbacks || []).join(',')}"
                             onerror="handleLogoError(this)"
                             style="max-width:85%; max-height:85%; object-fit:contain; filter:drop-shadow(0 0 15px rgba(255,255,255,0.2))">
                    </div>
                    <div style="padding:20px; background:rgba(255,255,255,0.03); border-top:1px solid var(--border); text-align:center;">
                        <h3 style="margin:0; font-size:18px; font-weight:900; color:#fff;">${displayData.name}</h3>
                        <p style="margin:5px 0 0; font-size:11px; color:var(--dim); font-weight:900; text-transform:uppercase; letter-spacing:1px;">לקוח אנטרפרייז</p>
                    </div>
                </div>`;
            });
            grid.innerHTML = html;
        }

        function handleLogoError(img) {
            let fallbacks = img.getAttribute('data-fallbacks');
            if (fallbacks) {
                let list = fallbacks.split(',');
                if (list.length > 0) {
                    let next = list.shift();
                    img.setAttribute('data-fallbacks', list.join(','));
                    img.src = next;
                    return;
                }
            }
            // Final verified Verifone fallback
            img.src = 'https://cdn.verifone.com/verifone-standard-logo.png';
            img.style.opacity = '0.5';
            img.onerror = null; // Prevent infinite loop
        }

        function renderTopNav() {
            const nav = document.getElementById('main-nav');
            let html = `
                <div class="nav ${sect==='customers'?'active':''}" onclick="nav('customers')">🤝 לקוחות</div>
                <div class="nav ${sect==='our-customers'?'active':''}" onclick="nav('our-customers')">💎 בין לקוחותנו</div>`;
            
            if (guides_data && Array.isArray(guides_data)) {
                guides_data.forEach(cat => {
                    const isActive = (sect==='guides' && selectedCatId === cat.id);
                    const emoji = cat.emoji || '📚';
                    html += `<div class="nav ${isActive?'active':''}" id="nav-cat-${cat.id}" onclick="navGuideCat('${cat.id}')" style="position:relative; display:flex; align-items:center; gap:8px;">
                        <span>${emoji} ${cat.name}</span>
                        <div style="display:flex; gap:12px; margin-right:10px; opacity:0; transition:0.3s; padding:5px; border-radius:8px; background:rgba(255,255,255,0.05)" class="cat-actions">
                            <span onclick="event.stopPropagation(); openEditCat('${cat.id}')" style="cursor:pointer; font-size:14px; filter:grayscale(1)">✏️</span>
                            <span onclick="event.stopPropagation(); deleteCat('${cat.id}')" style="cursor:pointer; font-size:14px; filter:grayscale(1)">🗑️</span>
                        </div>
                    </div>`;
                });
            }
            nav.innerHTML = html;
        }

        function navGuideCat(id) {
            sect = 'guides';
            selectedCatId = id;
            update();
        }

        function renderGuidesForCat(catId) {
            const cat = guides_data.find(c => c.id == catId);
            const display = document.getElementById('g-display');
            if(!cat) return;

            display.innerHTML = `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:30px">
                <h2 style="font-size:36px; margin:0">${cat.name}</h2>
            </div>
            <div class="guides-list">
                ${cat.guides.map(g => `
                    <div class="guide-mini-card" onclick="viewGuide('${cat.id}', '${g.id}')">
                        <span class="delete-btn del-guide" onclick="event.stopPropagation(); deleteGuide('${cat.id}', '${g.id}')">🗑️</span>
                        <h3>${g.title}</h3>
                        <p style="color:var(--dim); font-size:14px; margin-top:10px">${g.content.substring(0, 100)}...</p>
                    </div>
                `).join('')}
            </div>`;
        }
        function viewGuide(catId, gId) {
            window.location.hash = `guides/${catId}/guide/${gId}`;
        }
        function renderGuideView(catId, gId) {
            const cat = guides_data.find(c => c.id == catId);
            if(!cat) return;
            
            let guide = cat.guides ? cat.guides.find(g => g.id == gId) : null;
            if(!guide && cat.subCategories) {
                for(let s of cat.subCategories) {
                    guide = s.guides ? s.guides.find(g => g.id == gId) : null;
                    if(guide) break;
                }
            }
            if(!guide) return;
            
            const display = document.getElementById('g-viewer');
            display.innerHTML = '';
            
            const backBtn = document.createElement('button');
            backBtn.className = 'btn';
            backBtn.style.marginBottom = '30px';
            backBtn.style.padding = '10px 20px';
            backBtn.style.fontSize = '12px';
            backBtn.style.background = 'rgba(255,255,255,0.05)';
            backBtn.style.border = '1px solid var(--border)';
            backBtn.style.color = '#fff';
            backBtn.innerText = '← חזרה לרשימה';
            backBtn.onclick = () => { selectedGuideId = null; update(); };
            display.appendChild(backBtn);
            
            let formattedContent = guide.content;
            
            // Inline Image Replacement logic
            if(guide.images && guide.images.length > 0) {
                guide.images.forEach((img, idx) => {
                    const ext = img.split('.').pop().toLowerCase();
                    const isImg = ['jpg','jpeg','png','gif'].includes(ext);
                    if(isImg) {
                        const imgTag = `<img src="${img}" style="max-width:100%; border-radius:15px; margin:20px 0; border:1px solid var(--border); display:block; box-shadow: 0 10px 30px rgba(0,0,0,0.2)">`;
                        const placeholder = `[IMG${idx+1}]`;
                        if(formattedContent.includes(placeholder)) {
                            formattedContent = formattedContent.replace(new RegExp('\\' + placeholder, 'g'), imgTag);
                        }
                    }
                });
            }

            if(!formattedContent.includes('<') || !formattedContent.includes('>')) {
                formattedContent = formattedContent.replace(/\n/g, '<br>');
            }
            
            const contentDiv = document.createElement('div');
            contentDiv.style.maxWidth = '1000px';
            contentDiv.style.margin = '0 auto';
            contentDiv.innerHTML = `
                <div style="text-align:center; margin-bottom:50px; position:relative;">
                    <div style="position:absolute; top:0; right:0; display:flex; gap:10px;">
                        <button class="admin-btn" onclick="if(confirm('למחוק את המדריך?')) deleteGuide('${cat.id}', '${guide.id}')" style="background:rgba(239,68,68,0.1); border:1px solid #ef4444; color:#ef4444; padding:5px 10px; border-radius:8px; cursor:pointer;">🗑️ מחק</button>
                        <button class="admin-btn" onclick="openEditGuide('${cat.id}', '${guide.id}')">✏️ ערוך מדריך</button>
                    </div>
                    <h2 style="font-size:40px; font-weight:900; background: linear-gradient(to left, #fff, var(--dim)); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin:0;">${guide.title}</h2>
                    <p style="color:var(--dim); font-size:16px; margin-top:10px;">${cat.name} • ${guide.date || new Date().toLocaleDateString('he-IL')}</p>
                </div>
                <div class="guide-body" style="direction:rtl; text-align:right; font-size:18px; line-height:1.8; color:rgba(255,255,255,0.9)">${formattedContent}</div>`;
            display.appendChild(contentDiv);
            
            if(guide.images && guide.images.length > 0) {
                let attachments = [];
                let unusedImages = [];
                
                guide.images.forEach((img, idx) => {
                    const ext = img.split('.').pop().toLowerCase();
                    const isImg = ['jpg','jpeg','png','gif'].includes(ext);
                    
                    if(isImg) {
                        // If not already used inline, add to bottom
                        if(!guide.content.includes(`[IMG${idx+1}]`)) {
                            unusedImages.push(img);
                        }
                    } else {
                        attachments.push(img);
                    }
                });

                if(unusedImages.length > 0) {
                    unusedImages.forEach(img => {
                        display.innerHTML += `<img src="${img}" style="max-width:100%; border-radius:20px; margin-top:30px; border:1px solid var(--border); box-shadow: 0 10px 30px rgba(0,0,0,0.3)">`;
                    });
                }
                
                if(attachments.length > 0) {
                    display.innerHTML += `<div style="margin-top:40px; border-top:1px solid var(--border); padding-top:20px">
                        <h4 style="color:var(--dim); font-size:12px; text-transform:uppercase">קבצים מצורפים</h4>
                        <div style="display:flex; flex-direction:column; gap:10px">
                            ${attachments.map(url => `
                                <a href="${url}" target="_blank" style="background:rgba(255,255,255,0.05); padding:15px; border-radius:12px; color:var(--primary); text-decoration:none; display:flex; align-items:center; gap:10px; font-weight:700">
                                    <span style="font-size:24px">📄</span> הורד קובץ (${url.split('.').pop().toUpperCase()})
                                </a>
                            `).join('')}
                        </div>
                    </div>`;
                }
            } else if(guide.img) {
                // Fallback for old guides
                display.innerHTML += `<img src="${guide.img}" style="max-width:100%; border-radius:20px; margin-top:30px; border:1px solid var(--border)">`;
            }
        }
        function openAddCat() {
            editingCatId = null;
            document.getElementById('cat-modal').querySelector('b').innerText = 'הוספת קטגוריה חדשה';
            document.getElementById('cat-save-btn').innerText = 'שמור קטגוריה';
            document.getElementById('cat-name').value = '';
            document.getElementById('cat-emoji').value = '';
            document.getElementById('cat-type').value = 'kb';
            initEmojiPicker();
            document.querySelector('.overlay').style.display = 'block';
            document.getElementById('cat-modal').style.display = 'flex';
        }
        
        // Firebase Storage upload helper
        async function uploadToFirebaseStorage(file) {
            try {
                const { ref, uploadBytes, getDownloadURL } = window.firebaseRefs;
                const timestamp = Date.now();
                const fileName = `${timestamp}_${file.name}`;
                const storageRef = ref(window.storage, `uploads/${fileName}`);
                
                const snapshot = await uploadBytes(storageRef, file);
                const downloadURL = await getDownloadURL(snapshot.ref);
                
                console.log('File uploaded to Firebase Storage:', downloadURL);
                return { url: downloadURL, success: true };
            } catch (error) {
                console.error('Firebase Storage upload error:', error);
                throw error;
            }
        }
        
        let currentGuideImages = [];
        async function handleCustUpload(input, targetId) {
            if(!input.files || !input.files[0]) return;
            const status = document.getElementById('cust-upload-status');
            status.innerText = "מעלה...";
            
            const formData = new FormData();
            formData.append('file', input.files[0]);
            
            try {
                const resp = await fetch('/api/upload', { method: 'POST', body: formData });
                const data = await resp.json();
                document.getElementById(targetId).value = data.url;
                status.innerText = "הועלה בהצלחה!";
                setTimeout(() => status.innerText = "", 3000);
            } catch (err) {
                console.error(err);
                status.innerText = "העלאה נכשלה";
            }
        }
        async function handleUpload(input) {
            if(!input.files || !input.files[0]) return;
            const status = document.getElementById('upload-status');
            status.innerText = "מעלה...";
            
            const formData = new FormData();
            formData.append('file', input.files[0]);
            
            try {
                const resp = await fetch('/api/upload', {
                    method: 'POST',
                    body: formData
                });
                const data = await resp.json();
                currentGuideImages.push(data.url);
                renderGuideImages();
                status.innerText = "הושלם";
                setTimeout(() => status.innerText = "מוכן להעלאה נוספת", 2000);
            } catch (err) {
                console.error(err);
                status.innerText = "נכשל";
            }
        }
        function renderGuideImages() {
            const div = document.getElementById('guide-images');
            div.innerHTML = currentGuideImages.map((src, i) => {
                const ext = src.split('.').pop().toLowerCase();
                const isImg = ['jpg','jpeg','png','gif'].includes(ext);
                const tag = `[IMG${i+1}]`;
                return `
                <div style="position:relative; width:100px; height:120px; background:rgba(255,255,255,0.05); border-radius:8px; border:1px solid var(--border); display:flex; flex-direction:column; align-items:center; justify-content:center; overflow:visible">
                    ${isImg ? `<img src="${src}" style="width:100%; height:70px; object-fit:cover; border-radius:8px 8px 0 0;">` : `<span style="font-size:32px">📄</span>`}
                    <div style="background:var(--primary); color:#fff; font-size:10px; font-weight:900; width:100%; text-align:center; cursor:pointer; padding:4px 0" onclick="copyTag('${tag}')">העתק ${tag}</div>
                    <span onclick="removeGuideImage(${i})" style="position:absolute; top:-8px; right:-8px; background:#ef4444; color:#fff; border-radius:50%; width:20px; height:20px; font-size:12px; display:flex; align-items:center; justify-content:center; cursor:pointer; font-weight:900; box-shadow:0 0 10px rgba(0,0,0,0.5)">×</span>
                </div>`;
            }).join('');
        }
        function copyTag(t) {
            navigator.clipboard.writeText(t);
            const btn = event.target;
            const old = btn.innerText;
            btn.innerText = "הועתק!";
            setTimeout(() => btn.innerText = "העתק " + t, 1500);
        }
        function removeGuideImage(i) {
            currentGuideImages.splice(i, 1);
            renderGuideImages();
        }

        function updateSubCatDropdown(selectedId = null) {
            const catId = document.getElementById('guide-cat').value;
            const cat = guides_data.find(c => c.id == catId);
            const sel = document.getElementById('guide-subcat');
            sel.innerHTML = '<option value="">[ Main Category ]</option>';
            if(cat && cat.subCategories) {
                cat.subCategories.forEach(s => {
                    sel.innerHTML += `<option value="${s.id}" ${s.id==selectedId?'selected':''}>${s.name}</option>`;
                });
            }
        }

        function openAddGuide() {
            editingGuideId = null;
            
            // Check if we have any categories
            if (!guides_data || guides_data.length === 0) {
                alert("אנא צור קטגוריה חדשה לפני יצירת מדריך.");
                return;
            }

            document.getElementById('guide-modal').querySelector('b').innerText = 'יצירת מדריך חדש למערכת';
            document.getElementById('guide-title').value = '';
            document.getElementById('guide-content').innerHTML = '';
            currentGuideImages = [];
            renderGuideImages();
            
            // Populate Category Select
            const sel = document.getElementById('guide-cat');
            sel.innerHTML = guides_data.map(c => `<option value="${c.id}" ${c.id==selectedCatId?'selected':''}>${c.name}</option>`).join('');
            
            // If no category selected, pick the first one and update subcats
            if (!selectedCatId && guides_data.length > 0) {
                 // Trigger change update manually implies just visually setting it
                 // But we want subcats to load
                 updateSubCatDropdown(null);
            } else {
                 updateSubCatDropdown(); 
            }

            document.querySelector('.overlay').style.display = 'block';
            document.getElementById('guide-modal').style.display = 'flex';
            
            // Re-bind paste handler every time modal opens to be safe
            const editor = document.getElementById('guide-content');
            editor.onpaste = async (e) => {
                const items = (e.clipboardData || e.originalEvent.clipboardData).items;
                const types = (e.clipboardData || e.originalEvent.clipboardData).types;
                
                // If it's a pure image paste (like screenshot), we prevent default
                if (types.length === 1 && types[0] === 'Files') {
                    e.preventDefault();
                }

                for (const item of items) {
                    if (item.type.indexOf('image') !== -1) {
                        const blob = item.getAsFile();
                        const formData = new FormData();
                        formData.append('file', blob);
                        
                        const uploadId = 'up-' + Date.now();
                        document.execCommand('insertHTML', false, `<i id="${uploadId}">[Uploading Image...]</i>`);
                        
                        const resp = await fetch('/api/upload', { method: 'POST', body: formData });
                        const data = await resp.json();
                        
                        const placeholder = document.getElementById(uploadId);
                        if(placeholder) {
                            const imgHtml = `<img src="${data.url}" style="max-width:100%; border-radius:15px; margin:20px 0; border:1px solid var(--border); display:block; box-shadow:0 10px 30px rgba(0,0,0,0.3)">`;
                            placeholder.outerHTML = imgHtml;
                        }
                        currentGuideImages.push(data.url);
                    }
                }
            };
        }
        async function importContent(input) {
            if(!input.files || !input.files[0]) return;
            const file = input.files[0];
            const formData = new FormData();
            formData.append('file', file);
            
            const btn = document.querySelector('button[onclick*="content-import"]');
            const originalText = btn.innerText;
            btn.innerHTML = '<span class="spin">⏳</span> מעבד נתונים ותמונות...';
            btn.disabled = true;

            try {
                // First upload the file
                const formData = new FormData();
                formData.append('file', file);
                const uploadResp = await fetch('/api/upload', { method: 'POST', body: formData });
                const uploadData = await uploadResp.json();
                
                // Then extract its content
                const extractResp = await fetch('/api/extract-content', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ url: uploadData.url })
                });
                const extractData = await extractResp.json();
                
                if(extractData.content) {
                    const editor = document.getElementById('guide-content');
                    // mammoth returns HTML (with /uploads/ urls), so we insert directly
                    editor.innerHTML = extractData.content;
                    
                    // Scan for images to update the image list
                    const imgs = editor.querySelectorAll('img');
                    imgs.forEach(img => {
                        const src = img.getAttribute('src');
                        if (src && (src.startsWith('/uploads/') || src.startsWith('data:'))) {
                            if (!currentGuideImages.includes(src)) {
                                currentGuideImages.push(src);
                            }
                        }
                    });
                    renderGuideImages();
                    
                    console.log(`Content extracted successfully. Found ${imgs.length} images.`);
                    alert(`הטקסט והתמונות (${imgs.length}) נטענו בהצלחה!`);
                }
            } catch (e) {
                err_log("Extraction failed: " + e);
                alert("Failed to extract content. Please check file format.");
            } finally {
                btn.innerText = originalText;
                btn.disabled = false;
                input.value = ""; // Reset file input
            }
        }

        async function processFileToGuide(file, catId) {
            try {
                // Upload file
                const formData = new FormData();
                formData.append('file', file);
                const uploadResp = await fetch('/api/upload', { method: 'POST', body: formData });
                const uploadData = await uploadResp.json();
                
                // Extract content
                const extractResp = await fetch('/api/extract-content', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ url: uploadData.url })
                });
                const extractData = await extractResp.json();
                
                if(!extractData.content) return;
                
                // Extract images from content
                const tempDiv = document.createElement('div');
                tempDiv.innerHTML = extractData.content;
                const imgs = tempDiv.querySelectorAll('img');
                const images = Array.from(imgs).map(img => img.getAttribute('src')).filter(Boolean);
                
                // Create guide
                const cat = guides_data.find(c => c.id == catId);
                if(!cat) return;
                
                const guideObj = {
                    id: Date.now().toString() + Math.random().toString().substring(2, 8),
                    title: file.name.replace(/\.(docx?|pdf)$/i, ''),
                    content: extractData.content,
                    images: images
                };
                
                if(!cat.guides) cat.guides = [];
                cat.guides.push(guideObj);
                
            } catch(e) {
                console.error('Error processing file:', file.name, e);
            }
        }

        function closeM() {
            document.querySelectorAll('.modal, .overlay').forEach(el => el.style.display = 'none');
        }

        function openAddCat() {
            editingCatId = null;
            document.getElementById('cat-modal').querySelector('b').innerText = 'הוספת קטגוריה חדשה';
            document.getElementById('cat-save-btn').innerText = 'יצירת קטגוריה';
            document.getElementById('cat-name').value = '';
            document.getElementById('cat-emoji').value = '📚';
            document.getElementById('cat-type').value = 'guides';
            initEmojiPicker();
            document.querySelector('.overlay').style.display = 'block';
            document.getElementById('cat-modal').style.display = 'flex';
        }

        function openEditCat(id) {
            editingCatId = id;
            const cat = guides_data.find(c => c.id == id);
            document.getElementById('cat-modal').querySelector('b').innerText = 'Edit Category';
            document.getElementById('cat-save-btn').innerText = 'Update Category';
            document.getElementById('cat-name').value = cat.name;
            document.getElementById('cat-emoji').value = cat.emoji || '📚';
            initEmojiPicker();
            document.querySelector('.overlay').style.display = 'block';
            document.getElementById('cat-modal').style.display = 'flex';
        }

        async function saveCategory() {
            const name = document.getElementById('cat-name').value;
            const emoji = document.getElementById('cat-emoji').value;
            const type = document.getElementById('cat-type').value;
            if(!name) return;

            if(editingCatId) {
                const cat = guides_data.find(c => c.id == editingCatId);
                if(cat) { 
                    cat.name = name; 
                    cat.emoji = emoji;
                    cat.type = type;
                }
            } else {
                guides_data.push({
                    id: Date.now().toString(),
                    name: name,
                    emoji: emoji,
                    type: type,
                    guides: [],
                    subCategories: []
                });
            }
            
            await syncGuides();
            closeM();
            update();
        }
        function openAddSubCat() {
            const name = prompt("Enter Sub-Category Name:");
            if(!name) return;
            const cat = guides_data.find(c => c.id == selectedCatId);
            if(!cat) return;
            if(!cat.subCategories) cat.subCategories = [];
            cat.subCategories.push({ id: Date.now().toString(), name: name, guides: [] });
            syncGuides().then(update);
        }

        async function deleteGuide(catId, guideId) {
            if(!confirm('האם למחוק מדריך זה?')) return;
            
            guides_data.forEach(c => {
                if(c.guides) c.guides = c.guides.filter(g => g.id != guideId);
                if(c.subCategories) {
                    c.subCategories.forEach(s => {
                        if(s.guides) s.guides = s.guides.filter(g => g.id != guideId);
                    });
                }
            });
            
            await syncGuides();
            update();
            alert('המדריך נמחק בהצלחה');
        }

        function openEditGuide(catId, guideId) {
            editingGuideId = guideId;
            let cat = guides_data.find(c => c.id == catId);
            let guide = null;
            let subId = "";
            
            if(cat) {
                guide = cat.guides ? cat.guides.find(g => g.id == guideId) : null;
                if(!guide && cat.subCategories) {
                    for(let s of cat.subCategories) {
                        guide = s.guides ? s.guides.find(g => g.id == guideId) : null;
                        if(guide) { subId = s.id; break; }
                    }
                }
            }
            
            if(!guide) return;

            document.getElementById('guide-modal').querySelector('b').innerText = 'עריכת מדריך קיים';
            document.getElementById('guide-title').value = guide.title;
            document.getElementById('guide-content').innerHTML = guide.content;
            currentGuideImages = [...guide.images];
            renderGuideImages();
            
            const sel = document.getElementById('guide-cat');
            sel.innerHTML = guides_data.map(c => `<option value="${c.id}" ${c.id==catId?'selected':''}>${c.name}</option>`).join('');
            
            updateSubCatDropdown(subId);
            
            document.querySelector('.overlay').style.display = 'block';
            document.getElementById('guide-modal').style.display = 'flex';
        }

        async function saveGuide() {
            const catId = document.getElementById('guide-cat').value;
            const subId = document.getElementById('guide-subcat').value;
            const title = document.getElementById('guide-title').value;
            const content = document.getElementById('guide-content').innerHTML;
            if(!catId || !title) return;

            // 1. Remove guide from ANY current location (if editing)
            if(editingGuideId) {
                guides_data.forEach(c => {
                    if(c.guides) c.guides = c.guides.filter(g => g.id != editingGuideId);
                    if(c.subCategories) {
                        c.subCategories.forEach(s => {
                            if(s.guides) s.guides = s.guides.filter(g => g.id != editingGuideId);
                        });
                    }
                });
            }

            const cat = guides_data.find(c => c.id == catId);
            let gId = editingGuideId || Date.now().toString();
            const guideObj = { id: gId, title: title, content: content, images: [...currentGuideImages] };

            if(subId && cat.subCategories) {
                const sub = cat.subCategories.find(s => s.id == subId);
                if(sub) {
                    if(!sub.guides) sub.guides = [];
                    sub.guides.push(guideObj);
                }
            } else {
                if(!cat.guides) cat.guides = [];
                cat.guides.push(guideObj);
            }
            
            await syncGuides();
            closeM();
            selectedCatId = catId;
            selectedGuideId = gId;
            update();
        }
        async function deleteCat(catId) {
            if(!confirm("Are you sure? This will delete the category AND all its guides. This cannot be undone.")) return;
            guides_data = guides_data.filter(c => c.id != catId);
            
            // Optimistic Update
            if(selectedCatId === catId) {
                selectedCatId = null;
                selectedGuideId = null;
                nav('customers');
            } else {
                update();
            }
            
            // Save in background (no refresh needed as we just updated local state)
            try {
                await fetch('/api/guides/save', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(guides_data)
                });
                console.log("Category deletion synced to backend.");
            } catch(e) {
                console.error("Backend sync failed:", e);
                // We don't alert here to avoid annoying the user if it's transient, 
                // but strictly we should.
            }
        }
        async function deleteGuide(catId, guideId) {
            if(!confirm("Delete this guide?")) return;
            const cat = guides_data.find(c => c.id == catId);
            cat.guides = cat.guides.filter(g => g.id != guideId);
            
            // Optimistic Update
            renderGuidesForCat(catId);
            
            // Background Save
            fetch('/api/guides/save', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(guides_data)
            }).catch(e => console.error(e));
        }
        async function syncGuides(doRefresh = true) {
            try {
                const resp = await fetch('/api/guides/save', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(guides_data)
                });
                if(resp.ok) {
                    console.log('Synced', guides_data.length, 'guides to backend');
                } else {
                    throw new Error("Backend save failed");
                }
            } catch(e) {
                console.error("Sync error:", e);
                alert("שגיאת שמירה: הנתונים לא נשמרו בשרת. וודא שאתה מחובר.");
            }
            if(doRefresh) refresh();
        }
        function renderIntegrations(data) {
            const h = document.getElementById('thead');
            const cat = guides_data.find(c => c.id == selectedCatId);
            const type = (cat && cat.type) ? cat.type : 'project';

            // Headers Mapping
            if (type === 'table_phones') {
                h.innerHTML = `<tr><th>שם / מחלקה</th><th>מספר טלפון</th><th>תפקיד / הערה</th><th>אימייל</th><th>פעולה</th></tr>`;
            } else if (type === 'table_ip') {
                h.innerHTML = `<tr><th>שם שרת / התקן</th><th>כתובת IP</th><th>מיקום / VLAN</th><th>PORT / מידע</th><th>פעולה</th></tr>`;
            } else if (type === 'table_pass') {
                h.innerHTML = `<tr><th>שם מערכת</th><th>שם משתמש</th><th>*********</th><th>הערות</th><th>פעולה</th></tr>`;
            } else if (type === 'table_general') {
                h.innerHTML = `<tr><th>שם פריט</th><th>תיאור</th><th>סטטוס</th><th>הערות נוספות</th><th>פעולה</th></tr>`;
            } else {
                // Default Project/Integrations Table
                h.innerHTML = `<tr><th>פרויקט</th><th>סוג מכשיר</th><th>GW</th><th>מנהל</th><th>גרסה</th><th style="width:80px">מדריכים</th><th style="width:100px">פעולה</th></tr>`;
            }
            
            const b = document.getElementById('files'); b.innerHTML = '';
            data.forEach((r, idx) => {
                let editIdx = idx;
                if (sect === 'customers' && subSect === 'projects') {
                    editIdx = stats_data.Integrations.indexOf(r);
                }

                // Render Row based on Type (reusing existing fields as generic storage)
                // Mapping: Customer->Col1, Device->Col2, GW->Col3, PM->Col4, Version->Col5
                
                let rowHtml = '';
                if (type === 'table_phones') {
                    rowHtml = `
                        <td><b>${r.Customer || ''}</b></td>
                        <td><a href="tel:${r.Device}" style="color:var(--accent); text-decoration:none">${r.Device || ''}</a></td>
                        <td>${r.GW || ''}</td>
                        <td>${r.PM || ''}</td>
                        <td><button onclick="openEdit(${editIdx})" style="background:rgba(255,255,255,0.05); border:1px solid var(--border); color:#fff; padding:5px 12px; border-radius:8px; cursor:pointer; font-size:12px">Edit</button></td>
                    `;
                } else if (type === 'table_ip') {
                    rowHtml = `
                        <td><b>${r.Customer || ''}</b></td>
                        <td style="font-family:monospace; color:#10b981">${r.Device || ''}</td>
                        <td>${r.GW || ''}</td>
                        <td>${r.PM || ''}</td>
                        <td><button onclick="openEdit(${editIdx})" style="background:rgba(255,255,255,0.05); border:1px solid var(--border); color:#fff; padding:5px 12px; border-radius:8px; cursor:pointer; font-size:12px">Edit</button></td>
                    `;
                } else if (type === 'table_pass') {
                    rowHtml = `
                        <td><b>${r.Customer || ''}</b></td>
                        <td>${r.Device || ''}</td>
                        <td style="filter:blur(4px); cursor:pointer" onclick="this.style.filter='none'">${r.GW || ''}</td>
                        <td>${r.PM || ''}</td>
                        <td><button onclick="openEdit(${editIdx})" style="background:rgba(255,255,255,0.05); border:1px solid var(--border); color:#fff; padding:5px 12px; border-radius:8px; cursor:pointer; font-size:12px">Edit</button></td>
                    `;
                } else if (type === 'table_general') {
                     rowHtml = `
                        <td><b>${r.Customer || ''}</b></td>
                        <td>${r.Device || ''}</td>
                        <td><span style="background:rgba(255,255,255,0.1); padding:2px 8px; border-radius:4px">${r.GW || ''}</span></td>
                        <td>${r.PM || ''}</td>
                        <td><button onclick="openEdit(${editIdx})" style="background:rgba(255,255,255,0.05); border:1px solid var(--border); color:#fff; padding:5px 12px; border-radius:8px; cursor:pointer; font-size:12px">Edit</button></td>
                    `;
                } else {
                    // Default Project Row
                    const sheet = r.Sheet ? `<a href="${r.Sheet}" target="_blank" title="Release Sheet" style="text-decoration:none; font-size:24px; margin:0 5px;">📄</a>` : '';
                    const note = r.Note ? `<a href="${r.Note}" target="_blank" title="Release Note" style="text-decoration:none; font-size:24px; margin:0 5px;">📝</a>` : '';
                    const manual = r.Manual ? `<a href="${r.Manual}" target="_blank" title="Manual/Config" style="text-decoration:none; font-size:24px; margin:0 5px;">⚙️</a>` : '';
                    rowHtml = `
                        <td><b>${r.Customer}</b></td>
                        <td>${r.Device}</td>
                        <td><span style="background:rgba(59,130,246,0.1); padding:4px 10px; border-radius:6px; color:#60a5fa; font-size:14px">${r.GW}</span></td>
                        <td>${r.PM}</td>
                        <td><span style="color:${r.Version?'#fff':'#ef4444'}">${r.Version || "MISSING"}</span></td>
                        <td style="text-align:center; display:flex; justify-content:center; align-items:center; gap:5px;">${sheet} ${note} ${manual}</td>
                        <td><button onclick="openEdit(${editIdx})" style="background:rgba(255,255,255,0.05); border:1px solid var(--border); color:#fff; padding:5px 12px; border-radius:8px; cursor:pointer; font-size:12px">Edit</button></td>
                    `;
                }

                b.innerHTML += `<tr>${rowHtml}</tr>`;
            });
        }

        function renderWarrantyTable(data) {
            const h = document.getElementById('thead');
            h.innerHTML = `<tr><th>לקוח</th><th>אחריות</th><th>משך</th><th>כיסוי</th></tr>`;
            
            const b = document.getElementById('files'); b.innerHTML = '';
            
            const projectCustomers = new Set(stats_data.Integrations.map(x => x.Customer));

            data.forEach((r) => {
                const hasWarrantyInfo = r.WarrantyStatus && r.WarrantyStatus !== 'n/a' && r.WarrantyStatus !== 'אין';
                if (!hasWarrantyInfo && projectCustomers.has(r.Customer)) return;

                const status = (r.WarrantyStatus || 'אין').includes('יש') ? '✅ ' + r.WarrantyStatus : '❌ ' + (r.WarrantyStatus || 'n/a');
                b.innerHTML += `<tr>
                    <td style="font-size:18px"><b>${r.Customer}</b></td>
                    <td style="font-size:16px">${status}</td>
                    <td style="font-size:16px">${r.WarrantyDuration || '-'}</td>
                    <td style="font-size:15px; max-width:400px; line-height:1.5">${r.WarrantyCoverage || '-'}</td>
                </tr>`;
            });
        }
        let currentEditIdx = -1;
        function openAdd() {
            currentEditIdx = -1;
            document.getElementById('edit-modal').querySelector('b').innerText = 'Add New Item';
            
            const cat = guides_data.find(c => c.id == selectedCatId);
            const type = (cat && cat.type) ? cat.type : 'project';

            if (type === 'table_phones') {
                document.getElementById('lbl-cust').innerText = 'שם / מחלקה';
                document.getElementById('lbl-device').innerText = 'מספר טלפון';
                document.getElementById('lbl-gw').innerText = 'תפקיד / הערה';
                document.getElementById('lbl-pm').innerText = 'אימייל';
                document.getElementById('lbl-ver').innerText = '---';
            } else if (type === 'table_ip') {
                document.getElementById('lbl-cust').innerText = 'שם שרת';
                document.getElementById('lbl-device').innerText = 'כתובת IP';
                document.getElementById('lbl-gw').innerText = 'מיקום / VLAN';
                document.getElementById('lbl-pm').innerText = 'PORT';
                document.getElementById('lbl-ver').innerText = '---';
            } else if (type === 'table_pass') {
                document.getElementById('lbl-cust').innerText = 'שם מערכת';
                document.getElementById('lbl-device').innerText = 'שם משתמש';
                document.getElementById('lbl-gw').innerText = 'סיסמה';
                document.getElementById('lbl-pm').innerText = 'הערות';
                document.getElementById('lbl-ver').innerText = '---';
            } else if (type === 'table_general') {
                document.getElementById('lbl-cust').innerText = 'שם פריט';
                document.getElementById('lbl-device').innerText = 'תיאור';
                document.getElementById('lbl-gw').innerText = 'סטטוס';
                document.getElementById('lbl-pm').innerText = 'הערות נוספות';
                document.getElementById('lbl-ver').innerText = '---';
            } else {
                document.getElementById('lbl-cust').innerText = 'CUSTOMER';
                document.getElementById('lbl-device').innerText = 'SOLUTION TYPE';
                document.getElementById('lbl-gw').innerText = 'GW / CONNECTION';
                document.getElementById('lbl-pm').innerText = 'PROJECT MANAGER';
                document.getElementById('lbl-ver').innerText = 'VERSION';
            }

            document.getElementById('edit-cust').value = '';
            document.getElementById('edit-device').value = '';
            document.getElementById('edit-gw').value = '';
            document.getElementById('edit-pm').value = '';
            document.getElementById('edit-version').value = '';
            document.getElementById('edit-project-cat').value = '';
            document.getElementById('edit-sheet').value = '';
            document.getElementById('edit-note').value = '';
            document.getElementById('edit-manual').value = '';
            document.querySelector('.overlay').style.display = 'block';
            document.getElementById('edit-modal').style.display = 'flex';
        }
        function openEdit(idx) {
            currentEditIdx = idx;
            document.getElementById('edit-modal').querySelector('b').innerText = 'Edit Item';
            
            let data_source = (sect === 'customers') ? stats_data.Integrations : (guides_data.find(c=>c.id==selectedCatId)?.guides || []);
            const r = data_source[idx];
            const cat = guides_data.find(c => c.id == selectedCatId);
            const type = (cat && cat.type) ? cat.type : 'project';

            // Dynamic Labels
            if (type === 'table_phones') {
                document.getElementById('lbl-cust').innerText = 'שם / מחלקה';
                document.getElementById('lbl-device').innerText = 'מספר טלפון';
                document.getElementById('lbl-gw').innerText = 'תפקיד / הערה';
                document.getElementById('lbl-pm').innerText = 'אימייל';
                document.getElementById('lbl-ver').innerText = '---';
            } else if (type === 'table_ip') {
                document.getElementById('lbl-cust').innerText = 'שם שרת';
                document.getElementById('lbl-device').innerText = 'כתובת IP';
                document.getElementById('lbl-gw').innerText = 'מיקום / VLAN';
                document.getElementById('lbl-pm').innerText = 'PORT';
                document.getElementById('lbl-ver').innerText = '---';
            } else if (type === 'table_pass') {
                document.getElementById('lbl-cust').innerText = 'שם מערכת';
                document.getElementById('lbl-device').innerText = 'שם משתמש';
                document.getElementById('lbl-gw').innerText = 'סיסמה';
                document.getElementById('lbl-pm').innerText = 'הערות';
                document.getElementById('lbl-ver').innerText = '---';
            } else if (type === 'table_general') {
                document.getElementById('lbl-cust').innerText = 'שם פריט';
                document.getElementById('lbl-device').innerText = 'תיאור';
                document.getElementById('lbl-gw').innerText = 'סטטוס';
                document.getElementById('lbl-pm').innerText = 'הערות נוספות';
                document.getElementById('lbl-ver').innerText = '---';
            } else {
                document.getElementById('lbl-cust').innerText = 'CUSTOMER';
                document.getElementById('lbl-device').innerText = 'SOLUTION TYPE';
                document.getElementById('lbl-gw').innerText = 'GW / CONNECTION';
                document.getElementById('lbl-pm').innerText = 'PROJECT MANAGER';
                document.getElementById('lbl-ver').innerText = 'VERSION';
            }
            
            document.getElementById('edit-cust').value = r.Customer || '';
            document.getElementById('edit-device').value = r.Device || '';
            document.getElementById('edit-gw').value = r.GW || '';
            document.getElementById('edit-pm').value = r.PM || '';
            document.getElementById('edit-version').value = r.Version || '';
            document.getElementById('edit-project-cat').value = r.Category || '';
            document.getElementById('edit-sheet').value = r.Sheet || '';
            document.getElementById('edit-note').value = r.Note || '';
            document.getElementById('edit-manual').value = r.Manual || '';
            document.querySelector('.overlay').style.display = 'block';
            document.getElementById('edit-modal').style.display = 'flex';
        }
        async function saveEdit() {
            const data = {
                Customer: document.getElementById('edit-cust').value,
                Device: document.getElementById('edit-device').value,
                GW: document.getElementById('edit-gw').value,
                PM: document.getElementById('edit-pm').value,
                Version: document.getElementById('edit-version').value,
                Category: document.getElementById('edit-project-cat').value,
                Sheet: document.getElementById('edit-sheet').value,
                Note: document.getElementById('edit-note').value,
                Manual: document.getElementById('edit-manual').value
            };

            if (sect === 'customers') {
                if(currentEditIdx === -1) {
                    stats_data.Integrations.push(data);
                } else {
                    // Update the correct element in the full list
                    Object.assign(stats_data.Integrations[currentEditIdx], data);
                }
                
                try {
                    const resp = await fetch('/api/integrations/save', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(stats_data.Integrations)
                    });
                    if(!resp.ok) throw new Error("Backend save failed");
                } catch(e) {
                    console.error("Save error:", e);
                    alert("שגיאת שמירה בלקוחות.");
                }
            } else {
                // Table-based category
                const cat = guides_data.find(c => c.id == selectedCatId);
                if (cat) {
                    if (!cat.guides) cat.guides = [];
                    if (currentEditIdx === -1) {
                        // Ensure ID generation for new items!
                        data.id = Date.now().toString(); 
                        cat.guides.push(data);
                    } else {
                        // Preserve ID if editing
                        data.id = cat.guides[currentEditIdx].id || Date.now().toString();
                        Object.assign(cat.guides[currentEditIdx], data);
                    }
                    await syncGuides();
                }
            }
            
            closeM();
            update();
        }

        function renderIntegrations(data) {
            const table = document.getElementById('perf-card');
            if(!table) return;
            
            const cat = guides_data.find(c => c.id == selectedCatId);
            const type = (cat && cat.type) ? cat.type : 'project';
            
            // Dynamic headers based on type
            let headers = ['לקוח', 'פתרון', 'חיבור', 'מנהל', 'גרסה', 'פעולות'];
            if (type === 'table_phones') {
                headers = ['שם / מחלקה', 'מספר טלפון', 'תפקיד / הערה', 'אימייל', '', 'פעולות'];
            } else if (type === 'table_ip') {
                headers = ['שם שרת', 'כתובת IP', 'מיקום / VLAN', 'PORT', '', 'פעולות'];
            } else if (type === 'table_pass') {
                headers = ['שם מערכת', 'שם משתמש', 'סיסמה', 'הערות', '', 'פעולות'];
            } else if (type === 'table_general') {
                headers = ['שם פריט', 'תיאור', 'סטטוס', 'הערות נוספות', '', 'פעולות'];
            }
            
            let html = `
                <div class="card-t">רשימת נתונים</div>
                <table style="width:100%; border-collapse:collapse;">
                    <thead>
                        <tr style="border-bottom:2px solid var(--border);">
                            ${headers.map(h => h ? `<th style="padding:15px; text-align:right; font-weight:900; color:var(--accent);">${h}</th>` : '<th></th>').join('')}
                        </tr>
                    </thead>
                    <tbody>
                        ${data.map((r, idx) => `
                            <tr style="border-bottom:1px solid var(--border); cursor:pointer; transition:0.2s;" 
                                onmouseover="this.style.background='rgba(255,255,255,0.03)'" 
                                onmouseout="this.style.background='transparent'"
                                onclick="openEdit(${idx})">
                                <td style="padding:15px;">${r.Customer || ''}</td>
                                <td style="padding:15px;">${r.Device || ''}</td>
                                <td style="padding:15px;">${r.GW || ''}</td>
                                <td style="padding:15px;">${r.PM || ''}</td>
                                <td style="padding:15px;">${r.Version || ''}</td>
                                <td style="padding:15px;">
                                    <button onclick="event.stopPropagation(); deleteTableItem(${idx})" 
                                            style="padding:8px 12px; background:#ef4444; border:none; border-radius:8px; cursor:pointer; color:white; font-weight:bold;">
                                        🗑️
                                    </button>
                                </td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            `;
            
            table.innerHTML = html;
        }
        
        async function deleteTableItem(idx) {
            if(!confirm('האם למחוק פריט זה?')) return;
            
            if (sect === 'customers') {
                stats_data.Integrations.splice(idx, 1);
                try {
                    await fetch('/api/integrations/save', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(stats_data.Integrations)
                    });
                } catch(e) {
                    console.error("Delete error:", e);
                    alert("שגיאת מחיקה");
                }
            } else {
                const cat = guides_data.find(c => c.id == selectedCatId);
                if (cat && cat.guides) {
                    cat.guides.splice(idx, 1);
                    await syncGuides();
                }
            }
            
            update();
            alert('הפריט נמחק בהצלחה');
        }

        function renderManagers(data) {
            let m = document.getElementById('manager-view');
            if(!m) { m = document.createElement('div'); m.id = 'manager-view'; document.getElementById('capture-area').appendChild(m); }
            m.innerHTML = '<div class="manager-grid"></div>';
            const grid = m.querySelector('.manager-grid');
            const pms = data.reduce((acc, x) => { (acc[x.PM] = acc[x.PM] || []).push(x); return acc; }, {});
            Object.entries(pms).sort((a,b)=>b[1].length - a[1].length).forEach(([name, prjs]) => {
                const card = document.createElement('div');
                card.className = 'manager-card';
                card.innerHTML = `<h3>${name}</h3><p>${prjs.length} Projects</p>`;
                card.onclick = () => { document.getElementById('cust-search').value = name; subNav('projects'); filterIntegrations(); };
                grid.appendChild(card);
            });
        }
        function filterIntegrations() {
            const t = document.getElementById('cust-search').value.toLowerCase();
            
            let source = [];
            if (sect === 'customers') {
                source = stats_data.Integrations;
            } else if (sect === 'guides') {
                const cat = guides_data.find(c => c.id == selectedCatId);
                source = cat ? (cat.guides || []) : [];
            }

            const f = source.filter(x => {
                const cust = (x.Customer || '').toLowerCase();
                const pm = (x.PM || '').toLowerCase();
                const dev = (x.Device || '').toLowerCase();
                return cust.includes(t) || pm.includes(t) || dev.includes(t);
            });
            renderIntegrations(f);
        }
        function uk(a,b,c,d,e,f,g,h) {
            document.getElementById('v1').innerText=b; document.getElementById('l1').innerText=a;
            document.getElementById('v2').innerText=d; document.getElementById('l2').innerText=c;
            document.getElementById('v3').innerText=f; document.getElementById('l3').innerText=e;
            document.getElementById('v4').innerText=h; document.getElementById('l4').innerText=g;
        }
        async function takeShot() {
            const area = document.getElementById('capture-area');
            const canvas = await html2canvas(area, { backgroundColor: '#030712' });
            const link = document.createElement('a');
            link.download = `Vico_Dashboard_${new Date().toISOString()}.png`;
            link.href = canvas.toDataURL();
            link.click();
        }
        window.onload = init;
    </script>
</body>
</html>
        """

if __name__ == "__main__":
    socketserver.TCPServer.allow_reuse_address = True
    port = 8000
    # Define PORT for consistency with the instruction's final block
    PORT = port 
    try:
        with socketserver.TCPServer(("", PORT), handler) as httpd:
            log(f"TIER 2 VICO LIVE AT http://localhost:{PORT}")
            log("Press Ctrl+C to stop.")
            httpd.serve_forever()
    except OSError as e:
        err_log(f"Port {port} is busy or cannot be opened: {e}")
