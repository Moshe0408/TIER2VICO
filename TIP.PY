# -*- coding: utf-8 -*-
import win32com.client as win32
import os
import pytz
import datetime
import csv
import random
import shutil
import tempfile
import requests

# ===== נתונים =====

# שימוש בנתיב אבסולוטי לתמונה
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
IMAGE_FILENAME = os.path.join(SCRIPT_DIR, "customer_service_tip.png")
RECIPIENTS_CSV = os.path.join(SCRIPT_DIR, "recipients.csv")

# רשימת טיפים (לבקשת המשתמש - רענון 2026)
ALL_TIPS = [
    "הלקוח אומר 'לא נגעתי בכלום'? תתרגם ל'ניסיתי לתקן את זה עם מזלג ועכשיו הכל מקצר'. 🍴⚡",
    "מסוף לא מגיב? אולי הוא צריך רגע לחשוב על המשמעות של 'חיוב מוצלח'. 🤔💳",
    "המדפסת לא מוציאה נייר? תבדוק אם שמת גליל או שאתה סומך על כח המחשבה. 🧠📜",
    "תקלה בקופה ביום חמישי ב-19:55? זה לא באג, זה היקום אומר לך 'לך הביתה'. 🏠🌌",
    "הלקוח שואל אם אפשר לעקוף את הקוד הסודי? תציע לו לעקוף את חוקי הפיזיקה. 🛸🛡️",
    "נורית אדומה מהבהבת? זה אומר שהמסוף נכנס למצב 'לחץ סוציאלי'. 🔴😱",
    "המסך מגע לא עובד? תנקה את החומוס מהאצבע, זה עוזר פלאים. 🧆☝️",
    "אין תקשורת? המסוף והראוטר בדיוק באמצע ריב זוגי. תן להם דקה. 📶💔",
    "הלקוח אומר 'זה עבד אתמול'? גם אני הייתי צעיר ויפה אתמול. דברים משתנים. 👴✨",
    "ה-P400 עושה רעש של מכונת כביסה? הוא פשוט מנסה לכבס את הכסף של הלקוח. 🧼💸",
    "המקלדת בסינית? כנראה שהחתול של הלקוח ניסה להזמין משלוח מאלי-אקספרס. 🐈🧧",
    "הלקוח שואל 'למה זה לוקח זמן?' - תענה: 'אנחנו מחכים שהביטים יסיימו את הקפה'. ☕💾",
    "מסוף תלוש מהקיר? זה לא 'נתק', זה 'שאיפה לחופש'. 🕊️🔌",
    "הלקוח אומר 'זה תמיד עבד ככה' על תקלה ברורה? הוא כנראה בשלב ההכחשה. 🙈📈",
    "שגיאת תקשורת 001? המסוף פשוט ביישן ומפחד לדבר עם שבא. 🫣📡",
    "המסך של הקופה שחור? אולי היא פשוט משחקת 'מחבואים' עם הנתונים. 🎭🖥️",
    "הלקוח צועק? תזכור: זה לא אתה, זה הטון של הלקוח שלא מסתנכרן עם המציאות. 🧘‍♂️🎶",
    "הברקוד לא נסרק? תנסה לשיר לו, אולי הוא צריך קצת מוטיבציה. 🎤🔦",
    "שכחו את הסיסמה? תציע '1234' - המזל תמיד עוזר למי שמעז. 🔢🍀",
    "המסוף כותב 'ננעל עקב חבלה'? תגיד ללקוח שזה לא שק אגרוף, זה מסוף אשראי. 🔨😱",
    "ריח של שרוף? אל דאגה, אלו רק הרעיונות המבריקים של הצוות שעובדים קשה. 🔥🧠",
    "הלקוח אומר 'אני לא טכנולוגי' - זה הקוד ל'תעשה הכל במקומי ואני פשוט אסתכל'. 🛋️🛠️",
    "המסוף מוציא 'דף לבן'? הוא פשוט רוצה שתכתוב עליו שיר אהבה לתמיכה. 📝❤️",
    "תקלה בערב חג? זה הדרך של הגורל לבדוק כמה סבלנות נשארה לך אחרי הקניות. 🍎🍯",
    "הלקוח שואל 'זה יעבוד לנצח?' - תענה: 'עד העדכון הבא, או עד שיפול עליו קפה'. ☕📅",
    "העכבר לא זז? תבדוק אם הוא לא מת מרעב (או שפשוט הכבל בחוץ). 🐭🔌",
    "שגיאה לא מזוהה? זה פשוט המחשב שמנסה להיות מסתורי ומעניין. 🕵️‍♂️💻",
    "המערכת איטית? היא לא איטית, היא 'במחשבה עמוקה על הצעד הבא'. 🐢💭",
    "הלקוח אומר 'ניסיתי הכל' - חוץ מהדבר האחד שבאמת פותר את הבעיה. 🏆🔧",
    "ניסית לכבות ולהדליק? הקסם הכי ישן בספר, ועדיין עובד יותר מכל תואר ב-MIT. 🔄🪄",
    "הלקוח מאיים לעזוב? תזכור: בחוץ יש תקלות הרבה יותר גרועות, אנחנו לפחות נחמדים. 😊🛡️",
    "המסך של ה-V400 נהיה ורוד? הוא פשוט מאוהב בשירות המצוין שלך. 💖📟",
    "הלקוח אומר 'זה קרה פתאום' - בדיוק אחרי שהוא שפך עליו גלון של מים רותחים. 💧🔥",
    "תקלה גורפת? זה לא 'משבר', זה 'מסיבת עדכונים לא מתוכננת'. 🎊📉",
    "סיימת שיחה בהצלחה? תן לעצמך טפיחה על השכם, הרגע הצלת את העולם (או לפחות קנייה בסופר). 🌍🛒",
]

# רשימת GIFs משתנים ומשעשעים (Verified CDN Links - Buzzfeed/Imgur)
ALL_GIFS = [
    "https://img.buzzfeed.com/buzzfeed-static/static/2020-02/13/13/asset/68dc30da732b/anigif_sub-buzz-1034-1581602081-1.gif",
    "https://img.buzzfeed.com/buzzfeed-static/static/2020-02/13/13/asset/9b0705fe7d02/anigif_sub-buzz-363-1581602092-1.gif",
    "https://img.buzzfeed.com/buzzfeed-static/static/2020-02/13/15/asset/27f6d020c948/anigif_sub-buzz-2723-1581609514-1.gif",
    "https://i.imgur.com/6w2SOpz.gif",
    "https://i.imgur.com/RezqZWZ.gif",
    "https://i.imgur.com/xwm7z96.gif",
    "https://i.imgur.com/MzFCtct.gif",
    "https://i.imgur.com/9mwloRL.gif",
    "https://img.buzzfeed.com/buzzfeed-static/static/2020-02/13/14/asset/5bb055a630d4/anigif_sub-buzz-2590-1581602994-1.gif"
]

# ===== לוגיקה =====

def get_daily_greeting():
    israel_tz = pytz.timezone('Asia/Jerusalem')
    now = datetime.datetime.now(israel_tz)
    weekday = now.weekday()  # Monday=0, Sunday=6
    
    # מיפוי ימים לברכות
    greetings = {
        6: "✨ בוקר אור! שבוע חדש, אנרגיות חדשות!",      # Sunday
        0: "💪 יום שני – כבר נכנסנו לקצב!",              # Monday
        1: "🚀 יום שלישי – פעמיים כי טוב!",             # Tuesday
        2: "🌟 יום רביעי – אמצע השבוע, אנחנו על הגל!",    # Wednesday
        3: "🎉 יום חמישי הגיע! הסופ\"ש מעבר לפינה!",      # Thursday
    }
    # 4=Friday, 5=Saturday (נסנן אותם בלוגיקה הראשית)
    
    return greetings.get(weekday, "יום נפלא!")

def get_tip_and_gif():
    """
    בוחר טיפ ו-GIF המבוססים על היום בשנה כדי להבטיח שכל יום יהיה תוכן שונה.
    """
    israel_tz = pytz.timezone('Asia/Jerusalem')
    now = datetime.datetime.now(israel_tz)
    day_of_year = now.timetuple().tm_yday
    
    tip = ALL_TIPS[day_of_year % len(ALL_TIPS)] if ALL_TIPS else "אין טיפים זמינים."
    gif = ALL_GIFS[day_of_year % len(ALL_GIFS)] if ALL_GIFS else ""
    
    return tip, gif

def read_recipients():
    """קריאת הנמענים מקובץ ה-CSV המקומי"""
    emails = []
    if not os.path.exists(RECIPIENTS_CSV):
        print(f"FAILED: Recipients file not found: {RECIPIENTS_CSV}")
        return emails
        
    encodings = ['utf-8', 'utf-8-sig', 'windows-1255', 'latin-1']
    content = None
    
    for enc in encodings:
        try:
            with open(RECIPIENTS_CSV, encoding=enc) as csvfile:
                # בדיקה אם הקובץ ריק או לא תקין
                content = csvfile.read()
                if content:
                    # חזרה לתחילת הקובץ לקריאה עם DictReader
                    csvfile.seek(0)
                    reader = csv.DictReader(csvfile)
                    for row in reader:
                        email = row.get("email", "").strip()
                        if email and "@" in email:
                            emails.append(email)
                    if emails:
                        print(f"✅ הקובץ נקרא בהצלחה בקידוד: {enc}")
                        break
        except Exception:
            continue
            
    if not emails and content:
        print("ALERT: No valid emails found in file. Check for 'email' column.")
    elif not emails:
        print("FAILED: Could not read file with supported encodings.")
        
    return emails

def send_daily_email():
    israel_tz = pytz.timezone('Asia/Jerusalem')
    now = datetime.datetime.now(israel_tz)
    current_weekday = now.weekday() 
    
    # 4 = Friday, 5 = Saturday
    if current_weekday in [4, 5]:
        print(f"📅 היום {now.strftime('%A')} (סופ\"ש) - המערכת נחה.")
        return
 

    recipients_emails = read_recipients()
    if not recipients_emails:
        print("FAILED: No recipients found (file missing or empty).")
        return

    # בחירת תוכן
    greeting = get_daily_greeting()
    tip_text, selected_gif_url = get_tip_and_gif()
    
    # Download GIF to attach as CID (Most reliable for Outlook)
    temp_gif_path = None
    if selected_gif_url:
        try:
            print(f"🎬 מוריד GIF להטמעה: {selected_gif_url}")
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8'
            }
            # Add Referer based on domain
            if "buzzfeed" in selected_gif_url:
                headers['Referer'] = 'https://www.buzzfeed.com/'
            elif "imgur" in selected_gif_url:
                headers['Referer'] = 'https://imgur.com/'
            r = requests.get(selected_gif_url, stream=True, timeout=20, headers=headers)
            print(f"   [D] סטטוס הורדה: {r.status_code}")
            
            # בדיקה אם התוכן קצר מדי (כנראה placeholder של שגיאה)
            content_size = int(r.headers.get('Content-Length', 0))
            if r.status_code == 200 and content_size > 5000:
                temp_dir = tempfile.gettempdir()
                temp_gif_path = os.path.join(temp_dir, f"tip_gif_{now.strftime('%H%M%S')}.gif")
                with open(temp_gif_path, 'wb') as f:
                    shutil.copyfileobj(r.raw, f)
                print(f"   [V] GIF נשמר זמנית ב: {temp_gif_path} (גודל: {content_size} bytes)")
            else:
                print(f"   [!] אזהרה: הקובץ קטן מדי או שגיאת שרת ({r.status_code})")
        except Exception as e:
            print(f"   [!] אזהרה: לא ניתן היה להוריד את ה-GIF: {e}")

    # עיצוב HTML יוקרתי, ממורכז ונקי
    html_body = f"""
    <!DOCTYPE html>
    <html lang="he" dir="rtl">
    <head>
    <meta charset="UTF-8">
    <style>
        body {{
            margin: 0;
            padding: 0;
            background-color: #f4f7f9;
            font-family: Arial, Helvetica, sans-serif;
            direction: rtl;
        }}
        .email-container {{
            max-width: 600px;
            margin: 20px auto;
            background-color: #ffffff;
            border-radius: 12px;
            overflow: hidden;
            border: 1px solid #d1d9e0;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }}
        .header {{
            background-color: #0052cc;
            color: #ffffff;
            padding: 25px 20px;
            text-align: center;
        }}
        .header h1 {{
            margin: 0;
            font-size: 24px;
            font-weight: bold;
        }}
        .content {{
            padding: 35px 25px;
            text-align: center;
        }}
        .tip-box {{
            background-color: #f0f7ff;
            padding: 25px;
            border-radius: 12px;
            margin-bottom: 25px;
        }}
        .tip-text {{
            font-size: 22px;
            color: #1a1f23;
            line-height: 1.6;
            margin: 0;
            font-weight: bold;
        }}
        .funny-gif {{
            max-width: 100%;
            height: auto;
            border-radius: 12px;
            margin: 15px 0;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }}
        .greeting {{
            font-size: 20px;
            color: #0052cc;
            margin-top: 15px;
            font-weight: bold;
        }}
        .footer {{
            padding: 15px;
            text-align: center;
            background-color: #fcfdfe;
            color: #6a737d;
            font-size: 13px;
            border-top: 1px solid #e1e4e8;
        }}
        .verifone-link {{
            color: #0052cc;
            text-decoration: none;
            font-weight: bold;
        }}
    </style>
    </head>
    <body>
        <div class="email-container">
            <div class="header">
                <h1>💡 הטיפ היומי שלך</h1>
            </div>
            
            <div class="content">
                <div class="tip-box">
                    <p class="tip-text">
                        {tip_text}
                    </p>
                </div>

                {f'<img src="cid:daily_visual" alt="Customer Service Fun" class="funny-gif">' if temp_gif_path else ''}

                <div class="greeting">
                    {greeting}
                </div>
            </div>

            <div class="footer">
               של משה נשלח באופן אוטומטי • <b>גרסה 2.3</b>
                <br>
                <span class="verifone-link">Verifone Team</span> • {now.year}
            </div>
        </div>
    </body>
    </html>
    """


    # השולח הוא גם הנמען הראשי (כדי לקבל את ה"עותק" שנשלח וגם שהנמענים יראו ממי הגיע)
    # הנמענים האחרים יהיו ב-BCC
    sender_email = "moshei1@verifone.com" 
    
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        
        # למי שולחים? לעצמנו, והשאר ב-BCC מוסתר
        mail.To = sender_email
        mail.BCC = ";".join(recipients_emails)
        
        mail.Subject = f"הטיפ היומי שלך (V2.3) 💡 - {now.strftime('%d/%m/%Y %H:%M:%S')}"
        
        # הטמעת ה-GIF כ-CID
        if temp_gif_path and os.path.exists(temp_gif_path):
            attachment = mail.Attachments.Add(temp_gif_path)
            # Property ID for CID embedding
            attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", "daily_visual")
        
        mail.HTMLBody = html_body
        mail.Send()
        print(f"SUCCESS: Sent to {len(recipients_emails)} recipients (BCC), and to: {sender_email}")
        
    except Exception as e:
        print(f"FAILED: Error sending email: {e}")
    finally:
        # ניקוי קובץ זמני
        if temp_gif_path and os.path.exists(temp_gif_path):
            try: os.remove(temp_gif_path)
            except: pass

if __name__ == "__main__":
    print("--- מתחיל הפצת טיפ יומי ---")
    send_daily_email()
