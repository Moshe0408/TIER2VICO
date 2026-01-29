import sys, io
sys.stdout.reconfigure(encoding='utf-8')
import pythoncom
import xml.etree.ElementTree as ET
import pandas as pd
import threading
import requests
import json
import time
import os
from datetime import datetime, timedelta
from tkinter import messagebox, Tk, Text, END, Scrollbar, VERTICAL, RIGHT, Y, LEFT, BOTH, Frame, Label
from tkinter import ttk
import win32com.client as win32
import zipfile
import tempfile
from datetime import datetime

# פרטי התחברות
USER = "MosheI1"
PASSWORD = "Verifone!!2026"

def merge_ready_to_settle_files(ready_records, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    grouped = defaultdict(list)

    # קיבוץ לפי בית עסק
    for r in ready_records:
        merchant = r["account"]
        grouped[merchant].append(r["csv_file"])

    merged_files = []

    for merchant, file_list in grouped.items():
        frames = []
        for f in file_list:
            frames.append(pd.read_csv(f))

        merged_df = pd.concat(frames, ignore_index=True)
        file_path = os.path.join(output_dir, f"{merchant}_ReadyToSettle.xlsx")
        merged_df.to_excel(file_path, index=False)
        merged_files.append(file_path)

    return merged_files

    
# ================== רשימת חשבונות ==================
ENTITY_IDS = [
    "PISR00010633","PISR00010634","PISR00010555","PISR00010635",
    "PISR00010668","PISR00011014","PISR00028200","PISR00011831",
    "PISR00011015","PISR00011013","PISR00013665","PISR00018750",
]

# ----------------- רשימת סניפים מלאה -----------------
# ----------------- רשימת סניפים מלאה -----------------

try:
    BRANCH_DB = {
        "STISR0119426": {"branch_name": "שלי תא- בן יהודה-1", "business_code": "PISR00010555"},
        "STISR0119427": {"branch_name": "שלי ירושלים- אגרון-2", "business_code": "PISR00010555"},
        "STISR0119428": {"branch_name": "שלי גבעתיים- סירקין-3", "business_code": "PISR00010555"},
        "STISR0119429": {"branch_name": "שלי חיפה- כרמל-4", "business_code": "PISR00010555"},
        "STISR0119430": {"branch_name": "שלי ירושלים- יובל-5", "business_code": "PISR00010555"},
        "STISR0119431": {"branch_name": "שלי חולון- קוגל-6", "business_code": "PISR00010555"},
        "STISR0119432": {"branch_name": "שלי תא- ארלוזורוב-7", "business_code": "PISR00010555"},
        "STISR0119433": {"branch_name": "שלי נתניה- ויצמן-9", "business_code": "PISR00010555"},
        "STISR0119434": {"branch_name": "שלי תא- נורדאו-11", "business_code": "PISR00010555"},
        "STISR0119435": {"branch_name": "יש בני ברק- ירושלים-12", "business_code": "PISR00010634"},
        "STISR0119905": {"branch_name": "דיל בית שמש- העליה-13", "business_code": "PISR00010633"},
        "STISR0127242": {"branch_name": "שופרסל דיל ברנע אשקלון-14", "business_code": "PISR00010633"},
        "STISR0119912": {"branch_name": "יש פת- רוטשילד-15", "business_code": "PISR00010634"},
        "STISR0119916": {"branch_name": "מחסן אונליין מישוראדומים-16", "business_code": "PISR00010633"},
        "STISR0119917": {"branch_name": "שלי חיפה- חורב-17", "business_code": "PISR00010555"},
        "STISR0119918": {"branch_name": "דיל חולון-קרן היסוד-18", "business_code": "PISR00010633"},
        "STISR0119919": {"branch_name": "שלי חיפה- זיו-19", "business_code": "PISR00010555"},
        "STISR0119921": {"branch_name": "שלי רחובות- הרצל-20", "business_code": "PISR00010555"},
        "STISR0119922": {"branch_name": "יש ירושלים מרים-פארן-21", "business_code": "PISR00010634"},
        "STISR0119923": {"branch_name": "שלי נתניה- סמילנסקי-22", "business_code": "PISR00010555"},
        "STISR0119925": {"branch_name": "שלי מיתר-23", "business_code": "PISR00010555"},
        "STISR0119926": {"branch_name": "שלי אשדוד- הנביאים-24", "business_code": "PISR00010555"},
        "STISR0131035": {"branch_name": "אקספרס הכרמל כפר סבא-25", "business_code": "PISR00010635"},
        "STISR0121575": {"branch_name": "בי דראגסטורס בלוך גבעתיים-26", "business_code": "PISR00010668"},
        "STISR0119927": {"branch_name": "שלי רעננה- אחוזה-27", "business_code": "PISR00010555"},
        "STISR0119280": {"branch_name": "שלי תא- צהל-28", "business_code": "PISR00010555"},
        "STISR0121215": {"branch_name": "אקספרס חריש האודם-29", "business_code": "PISR00010635"},
        "STISR0120255": {"branch_name": "שלי רמת גן- קריניצי-30", "business_code": "PISR00010555"},
        "STISR0120089": {"branch_name": "שלי רמת גן- קסם-32", "business_code": "PISR00010555"},
        "STISR0120309": {"branch_name": "שלי חיפה- סטלה-33", "business_code": "PISR00010555"},
        "STISR0120090": {"branch_name": "שלי רמת השרון- סוקולוב-34", "business_code": "PISR00010555"},
        "STISR0120310": {"branch_name": "דיל אקסטרה באר-שבע ולפסו-35", "business_code": "PISR00011014"},
        "STISR0120311": {"branch_name": "שלי כ\"ס- רוטשילד-36", "business_code": "PISR00010555"},
        "STISR0120299": {"branch_name": "שלי ב\"ש- יעלים-37", "business_code": "PISR00010555"},
        "STISR0120312": {"branch_name": "שלי חיפה- דניה-38", "business_code": "PISR00010555"},
        "STISR0121942": {"branch_name": "שלי תא- ברזיל-39", "business_code": "PISR00010555"},
        "STISR0120313": {"branch_name": "שלי ב\"ש- עומר-40", "business_code": "PISR00010555"},
        "STISR0120252": {"branch_name": "בי דראגסטורס דזינגוף סנטר-41", "business_code": "PISR00010668"},
        "STISR0121138": {"branch_name": "שלי ירושלים- ניות-42", "business_code": "PISR00010555"},
        "STISR0120300": {"branch_name": "שלי הוד השרון- הבנים-43", "business_code": "PISR00010555"},
        "STISR0121423": {"branch_name": "דיל ירושלים- תלפיות-45", "business_code": "PISR00010633"},
        "STISR0122154": {"branch_name": "דיל רמלה- היצירה-46", "business_code": "PISR00010633"},
        "STISR0121943": {"branch_name": "דיל ב\"ש- הר בוקר-49", "business_code": "PISR00010633"},
        "STISR0120800": {"branch_name": "שלי מבשרת ציון-50", "business_code": "PISR00010555"},
        "STISR0134340": {"branch_name": "אקספרס צמרות הרצליה-51", "business_code": "PISR00010635"},
        "STISR0120789": {"branch_name": "אקספרס חריש הרימון-56", "business_code": "PISR00010635"},
        "STISR0120560": {"branch_name": "שלי רמת השרון- אוסישקין-57", "business_code": "PISR00010555"},
        "STISR0120460": {"branch_name": "אקספרס צור משה-61", "business_code": "PISR00010635"},
        "STISR0120301": {"branch_name": "יש חסד רכסים-62", "business_code": "PISR00010634"},
        "STISR0121424": {"branch_name": "דיל ב\"ש- צ'ורלי-65", "business_code": "PISR00010633"},
        "STISR0121955": {"branch_name": "שלי גבעת שמואל-בן גוריון-67", "business_code": "PISR00010555"},
        "STISR0120315": {"branch_name": "דיל כורדני-68", "business_code": "PISR00010633"},
        "STISR0120535": {"branch_name": "שלי ת\"א- רמת אביב ב-69", "business_code": "PISR00010555"},
        "STISR0121139": {"branch_name": "שלי רמת גן- מרום נווה-70", "business_code": "PISR00010555"},
        "STISR0119391": {"branch_name": "דיל אקסטרה תל חנן נשר-71", "business_code": "PISR00011014"},
        "STISR0120088": {"branch_name": "שלי מודיעין רעות-72", "business_code": "PISR00010555"},
        "STISR0121425": {"branch_name": "יש חסד אשדוד-73", "business_code": "PISR00010634"},
        "STISR0121238": {"branch_name": "דיל נתניה- הדרים-74", "business_code": "PISR00010633"},
        "STISR0121239": {"branch_name": "דיל אריאל-76", "business_code": "PISR00010633"},
        "STISR0121230": {"branch_name": "דיל חריש-77", "business_code": "PISR00010633"},
        "STISR0120665": {"branch_name": "יש הר נוף ירושלים-שאולזו-78", "business_code": "PISR00010634"},
        "STISR0120556": {"branch_name": "אקספרס חריש-79", "business_code": "PISR00010635"},
        "STISR0120493": {"branch_name": "שלי ראשל\"צ- נווה הדרים-80", "business_code": "PISR00010555"},
        "STISR0120823": {"branch_name": "יש חסד מודיעין עלית דרום-81", "business_code": "PISR00010634"},
        "STISR0121216": {"branch_name": "אקספרס חולון-82", "business_code": "PISR00010635"},
        "STISR0121227": {"branch_name": "דיל ב\"ש- שאול המלך-83", "business_code": "PISR00010633"},
        "STISR0120389": {"branch_name": "אקספרס בית חשמונאי-84", "business_code": "PISR00010635"},
        "STISR0120494": {"branch_name": "דיל רעננה- החרושת-87", "business_code": "PISR00010633"},
        "STISR0120495": {"branch_name": "דיל עזריאלי - רמלה-89", "business_code": "PISR00010633"},
        "STISR0120395": {"branch_name": "אקספרס גב ים באר שבע-90", "business_code": "PISR00010635"},
        "STISR0121242": {"branch_name": "דיל עמק חפר- אזור תעשיה-91", "business_code": "PISR00010633"},
        "STISR0121112": {"branch_name": "דיל אשדוד- שבט לוי-92", "business_code": "PISR00010633"},
        "STISR0121561": {"branch_name": "שלי יהוד- סביונים-93", "business_code": "PISR00010555"},
        "STISR0121130": {"branch_name": "דיל אור עקיבא קניון אורו-95", "business_code": "PISR00010633"},
          "STISR0121513": {"branch_name": "דיל חדרה- הפרדס-96", "business_code": "PISR00010633"},
          "STISR0120092": {"branch_name": "דיל קרית שמונה- מתחם ביג-97", "business_code": "PISR00010633"},
          "STISR0121120": {"branch_name": "דיל - אלונים-98", "business_code": "PISR00010633"},
          "STISR0120086": {"branch_name": "שלי הוד השרון- ק מרגלית-101", "business_code": "PISR00010555"},
          "STISR0121224": {"branch_name": "שלי פרדסיה- הנשיא-102", "business_code": "PISR00010555"},
          "STISR0121140": {"branch_name": "שלי שוהם- מרכז מסחרי-103", "business_code": "PISR00010555"},
          "STISR0121243": {"branch_name": "שלי פ\"ת- גד מכנס-104", "business_code": "PISR00010555"},
          "STISR0120302": {"branch_name": "דיל פ\"ת- אליעזר פרדימן-105", "business_code": "PISR00010633"},
          "STISR0119393": {"branch_name": "דיל אקסטרה נס ציונה הפטיש-106", "business_code": "PISR00011014"},
          "STISR0120496": {"branch_name": "דיל רחובות- קניון-107", "business_code": "PISR00010633"},
          "STISR0121121": {"branch_name": "שלי ראש העין- ז'בוטינסקי-109", "business_code": "PISR00010555"},
          "STISR0133716": {"branch_name": "דיל הפארק באר שבע-110", "business_code": "PISR00010633"},
          "STISR0120303": {"branch_name": "יש חסד טבריה עלית-לב האג-111", "business_code": "PISR00010634"},
          "STISR0148545": {"branch_name": "קרית מוצקין BE-112", "business_code": "PISR00010668"},
          "STISR0121223": {"branch_name": "דיל ראשל\"צ- גולדה מאיר-113", "business_code": "PISR00010633"},
          "STISR0121234": {"branch_name": "דיל ירושלים- גילה-114", "business_code": "PISR00010633"},
          "STISR0120468": {"branch_name": "אקספרס דיזינגוף-115", "business_code": "PISR00010635"},
          "STISR0120822": {"branch_name": "שלי גבעתיים- מכתש-116", "business_code": "PISR00010555"},
          "STISR0120536": {"branch_name": "שלי תל אביב-איכילוב-117", "business_code": "PISR00010555"},
          "STISR0120314": {"branch_name": "דיל אילת הסתת-118", "business_code": "PISR00010633"},
          "STISR0120537": {"branch_name": "דיל מודיעין- סנטר-119", "business_code": "PISR00010633"},
          "STISR0119394": {"branch_name": "דיל אקסטרה חולון-המרכבה-120", "business_code": "PISR00011014"},
          "STISR0121236": {"branch_name": "דיל נהריה- לוחמי הגטאות-121", "business_code": "PISR00010633"},
          "STISR0121122": {"branch_name": "דיל יבנה- ברוש דרך הים-122", "business_code": "PISR00010633"},
          "STISR0121221": {"branch_name": "דיל חולון- גולדה-123", "business_code": "PISR00010633"},
          "STISR0121237": {"branch_name": "דיל בת ים- אורט ישראל-124", "business_code": "PISR00010633"},
          "STISR0119284": {"branch_name": "דיל הרצליה- המנופים-125", "business_code": "PISR00010633"},
          "STISR0121190": {"branch_name": "אקספרס טירת הכרמל-127", "business_code": "PISR00010635"},
          "STISR0119395": {"branch_name": "דיל טירת הכרמל- נחום חת-128", "business_code": "PISR00010633"},
          "STISR0121229": {"branch_name": "דיל אקסטרה נתניה- המלאכה-129", "business_code": "PISR00011014"},
          "STISR0121131": {"branch_name": "דיל נתניה- קלאוזנר עמליה-130", "business_code": "PISR00010633"},
          "STISR0125326": {"branch_name": "אקספרס סוקולוב הוד השרון-131", "business_code": "PISR00010635"},
          "STISR0119396": {"branch_name": "דיל ת\"א- השלום-132", "business_code": "PISR00010633"},
          "STISR0121916": {"branch_name": "יש חסד ביתר עילית- הר\"ן-133", "business_code": "PISR00010634"},
          "STISR0120257": {"branch_name": "דיל מודיעין- ישפרו סנטר-134", "business_code": "PISR00010633"},
          "STISR0121504": {"branch_name": "דיל אשקלון- פאור סנטר-135", "business_code": "PISR00010633"},
          "STISR0148501": {"branch_name": "שופרסל סטוק אופקים-136", "business_code": "PISR00028200"},
          "STISR0120304": {"branch_name": "דיל אקסטרה ב\"ש-סולל בונה-138", "business_code": "PISR00011014"},
          "STISR0121562": {"branch_name": "דיל צפת- דובק ויצמן-139", "business_code": "PISR00010633"},
          "STISR0125327": {"branch_name": "אקספרס תל חי כפר סבא-140", "business_code": "PISR00010635"},
          "STISR0121945": {"branch_name": "דיל יוקנעם- שדרות רבין-141", "business_code": "PISR00010633"},
          "STISR0120546": {"branch_name": "דיל זכרון- המייסדים-142", "business_code": "PISR00010633"},
          "STISR0120462": {"branch_name": "אקספרס חיפה מוריה-143", "business_code": "PISR00010635"},
          "STISR0121225": {"branch_name": "דיל שבירו כפר סבא-144", "business_code": "PISR00010633"},
          "STISR0121686": {"branch_name": "בי דראגסטורס נתיבות-145", "business_code": "PISR00010668"},
          "STISR0120538": {"branch_name": "שלי גבעתיים- הסתדרות-146", "business_code": "PISR00010555"},
          "STISR0120641": {"branch_name": "אקספרס זכרון-147", "business_code": "PISR00010635"},
          "STISR0121416": {"branch_name": "אקספרס מלצר רחובות-148", "business_code": "PISR00010635"},
          "STISR0120640": {"branch_name": "אקספרס רודנסקי תל אביב-150", "business_code": "PISR00010635"},
          "STISR0120087": {"branch_name": "דיל רעננה- קניון רננים-151", "business_code": "PISR00010633"},
          "STISR0120258": {"branch_name": "דיל ערד- ישפרו המנוף-152", "business_code": "PISR00010633"},
          "STISR0120091": {"branch_name": "דיל אופקים- ז'בוטינסקי-153", "business_code": "PISR00010633"},
          "STISR0121563": {"branch_name": "אקספרס קריית אונו-154", "business_code": "PISR00010635"},
          "STISR0121241": {"branch_name": "דיל צור יגאל- מרכז מסחרי-155", "business_code": "PISR00010633"},
          "STISR0121616": {"branch_name": "באר יעקב C&C סניף-156", "business_code": "PISR00011831"},
          "STISR0121617": {"branch_name": "cash & carry חיפה-157", "business_code": "PISR00011831"},
          "STISR0120259": {"branch_name": "דיל רמת הנשיא חיפה-159", "business_code": "PISR00010633"},
          "STISR0120305": {"branch_name": "שלי רמת גן- הירדן-160", "business_code": "PISR00010555"},
          "STISR0127471": {"branch_name": "שופרסל cash & carry אשדוד-161", "business_code": "PISR00011831"},
          "STISR0121249": {"branch_name": "דיל שדרות- הפלדה-163", "business_code": "PISR00010633"},
          "STISR0121564": {"branch_name": "שלי רמת גן- ביאליק-164", "business_code": "PISR00010555"},
          "STISR0123539": {"branch_name": "בי דראגסטורס המרכבה חולון-165", "business_code": "PISR00010668"},
          "STISR0120316": {"branch_name": "דיל אקסטרה אשדוד- הבושם-166", "business_code": "PISR00011014"},
          "STISR0119397": {"branch_name": "דיל ראשל\"צ- הלח\"י-167", "business_code": "PISR00010633"},
          "STISR0121141": {"branch_name": "שלי אשדוד- כינרת-168", "business_code": "PISR00010555"},
          "STISR0121426": {"branch_name": "אקספרס ב\"ש אפרים-169", "business_code": "PISR00010635"},
          "STISR0143006": {"branch_name": "חנות עובדים מרלוג ראשלצ-170", "business_code": "PISR00010633"},
          "STISR0121219": {"branch_name": "אקספרס אוסטשינסקי כ.סבא-171", "business_code": "PISR00010635"},
          "STISR0148546": {"branch_name": "be באר יעקב-172", "business_code": "PISR00010668"},
          "STISR0126934": {"branch_name": "דיל כפר יונה-173", "business_code": "PISR00010633"},
          "STISR0120497": {"branch_name": "דיל ערד- קניון-174", "business_code": "PISR00010633"},
          "STISR0121150": {"branch_name": "אקספרס טבריה-175", "business_code": "PISR00010635"},
          "STISR0121593": {"branch_name": "שערי רווחה כדורי-176", "business_code": "PISR00011015"},
          "STISR0121250": {"branch_name": "דיל מצפה רמון- בן גוריון-177", "business_code": "PISR00010633"},
          "STISR0121151": {"branch_name": "סגולה פתח תקווה Be-178", "business_code": "PISR00010668"},
          "STISR0120463": {"branch_name": "אקספרס שד ירושלים ר\"ג-179", "business_code": "PISR00010635"},
          "STISR0121222": {"branch_name": "דיל ק.מוצקין- בן גוריון-180", "business_code": "PISR00010633"},
          "STISR0120498": {"branch_name": "דיל אשקלון- רמז-181", "business_code": "PISR00010633"},
          "STISR0121123": {"branch_name": "דיל דימונה- גולדה מאיר-182", "business_code": "PISR00010633"},
          "STISR0121505": {"branch_name": "דיל אשדוד- הרצל-183", "business_code": "PISR00010633"},
          "STISR0121124": {"branch_name": "דיל ראשל\"צ- זבוטינסקי-184", "business_code": "PISR00010633"},
          "STISR0120790": {"branch_name": "אקספרס התשבי-186", "business_code": "PISR00010635"},
          "STISR0121565": {"branch_name": "דיל מעלה אדומים-187", "business_code": "PISR00010633"},
          "STISR0121506": {"branch_name": "דיל ירוחם- צבי בורנשטיין-188", "business_code": "PISR00010633"},
          "STISR0121142": {"branch_name": "דיל פרדס חנה-189", "business_code": "PISR00010633"},
          "STISR0121956": {"branch_name": "דיל חצור- הגלילית-190", "business_code": "PISR00010633"},
          "STISR0119400": {"branch_name": "שלי ראשל\"צ- מזרח-193", "business_code": "PISR00010555"},
          "STISR0121113": {"branch_name": "דיל קריון- דרך עכו-195", "business_code": "PISR00010633"},
          "STISR0120266": {"branch_name": "שלי-אבן גבירול-196", "business_code": "PISR00010555"},
          "STISR0119398": {"branch_name": "דיל בת ים- בלפור-199", "business_code": "PISR00010633"},
          "STISR0121205": {"branch_name": "בי דראגסטורס בריגה נתניה-201", "business_code": "PISR00010668"},
          "STISR0121125": {"branch_name": "דיל רמלה לוד- הצופית-203", "business_code": "PISR00010633"},
          "STISR0121114": {"branch_name": "דיל יהוד- אלטלף-205", "business_code": "PISR00010633"},
          "STISR0119285": {"branch_name": "יש אור יהודה-אליהו סעדו-207", "business_code": "PISR00010634"},
          "STISR0120791": {"branch_name": "דיל אילת נחל אורה-208", "business_code": "PISR00010633"},
          "STISR0121126": {"branch_name": "דיל אשקלון- רוטשילד-209", "business_code": "PISR00010633"},
          "STISR0119399": {"branch_name": "יוניברס ראשל\"צ- שמוטקין-210", "business_code": "PISR00010633"},
          "STISR0120261": {"branch_name": "דיל כרמיאל- אזור התעשיה-211", "business_code": "PISR00010633"},
          "STISR0121950": {"branch_name": "שלי חיפה- ורדיה-212", "business_code": "PISR00010555"},
          "STISR0120539": {"branch_name": "דיל קרית גת- מלכי ישראל-214", "business_code": "PISR00010633"},
          "STISR0120499": {"branch_name": "דיל בית שאן-ירושלים הביר-215", "business_code": "PISR00010633"},
          "STISR0129611": {"branch_name": "שופרסל שלי הוד השרון הנשיאים-216", "business_code": "PISR00010555"},
          "STISR0121233": {"branch_name": "אקספרס הדישון ירושלים-217", "business_code": "PISR00010635"},
          "STISR0121132": {"branch_name": "דיל עפולה- יצחק רבין-218", "business_code": "PISR00010633"},
          "STISR0121918": {"branch_name": "יש חסד בני ברק- עזרא-219", "business_code": "PISR00010634"},
          "STISR0121133": {"branch_name": "דיל רמות ב\"ש-222", "business_code": "PISR00010633"},
          "STISR0127334": {"branch_name": "שופרסל אקספרס מודיעין-223", "business_code": "PISR00010635"},
          "STISR0120642": {"branch_name": "אקספרס ראשל\"צ-נווה דקלים-224", "business_code": "PISR00010635"},
          "STISR0121127": {"branch_name": "דיל כרמיאל-226", "business_code": "PISR00010633"},
          "STISR0120540": {"branch_name": "דיל-איקאה קרית אתא-227", "business_code": "PISR00010633"},
          "STISR0120306": {"branch_name": "שלי בנימינה - הגביע-228", "business_code": "PISR00010555"},
          "STISR0121134": {"branch_name": "דיל ת\"א- יגאל אלון-229", "business_code": "PISR00010633"},
          "STISR0120051": {"branch_name": "שלי כ\"ס- ויצמן-230", "business_code": "PISR00010555"},
          "STISR0121220": {"branch_name": "אקספרס רעננה- רבקה גרובר-231", "business_code": "PISR00010635"},
          "STISR0120792": {"branch_name": "שלי צפת- מ.מסחר רמת רזי-232", "business_code": "PISR00010555"},
          "STISR0133344": {"branch_name": "בי דראגסטורס נחל פרת -233", "business_code": "PISR00010668"},
          "STISR0120317": {"branch_name": "יש חסד רסידו ב.שמש-234", "business_code": "PISR00010634"},
          "STISR0121218": {"branch_name": "אקספרס אילת שחמון-236", "business_code": "PISR00010635"},
          "STISR0121685": {"branch_name": "דיל גוש עציון-237", "business_code": "PISR00010633"},
          "STISR0121240": {"branch_name": "דיל באר יעקב-238", "business_code": "PISR00010633"},
          "STISR0120643": {"branch_name": "אקספרס יפו- הדרור-239", "business_code": "PISR00010635"},
          "STISR0121135": {"branch_name": "דיל רהט-240", "business_code": "PISR00010633"},
          "STISR0120318": {"branch_name": "דיל בני ברק איילון-241", "business_code": "PISR00010633"},
          "STISR0123321": {"branch_name": "בי דראגסטורס ויוה חדרה-242", "business_code": "PISR00010668"},
          "STISR0120561": {"branch_name": "רחובות Be-243", "business_code": "PISR00010668"},
          "STISR0121128": {"branch_name": "דיל גלילות רמת השרון-244", "business_code": "PISR00010633"},
          "STISR0120474": {"branch_name": "דיל ירושלים- פסגת זאב-245", "business_code": "PISR00010633"},
          "STISR0120644": {"branch_name": "שלי חולון- רבינוביץ-247", "business_code": "PISR00010555"},
          "STISR0121919": {"branch_name": "דיל ב\"ש- נווה מנחם-248", "business_code": "PISR00010633"},
          "STISR0120541": {"branch_name": "דיל נתניה- פולג-249", "business_code": "PISR00010633"},
          "STISR0120475": {"branch_name": "דיל ק.שמונה מצודות-250", "business_code": "PISR00010633"},
          "STISR0121514": {"branch_name": "דיל עפולה- כורש-251", "business_code": "PISR00010633"},
          "STISR0121626": {"branch_name": "בי דאטגסטורס עין שמר-252", "business_code": "PISR00010668"},
          "STISR0132850": {"branch_name": "דיל כרמיגת-254", "business_code": "PISR00010633"},
          "STISR0121115": {"branch_name": "דיל כפר תבור-255", "business_code": "PISR00010633"},
          "STISR0121422": {"branch_name": "נס ציונה ישפרו Be-256", "business_code": "PISR00010668"},
          "STISR0121507": {"branch_name": "דיל מרקט כפר סבא-257", "business_code": "PISR00010633"},
          "STISR0119903": {"branch_name": "דיל אדמירליטי חיפה- איטי-258", "business_code": "PISR00010633"},
          "STISR0121516": {"branch_name": "דיל גזית כ\"ס- ויצמן-259", "business_code": "PISR00010633"},
          "STISR0120115": {"branch_name": "דיל ירושלים-קניון תלפיות-260", "business_code": "PISR00010633"},
          "STISR0121920": {"branch_name": "דיל באר טוביה-261", "business_code": "PISR00010633"},
          "STISR0121129": {"branch_name": "שלי מזכרת בתיה- בגין-262", "business_code": "PISR00010555"},
          "STISR0121231": {"branch_name": "אקספרס מקס ברוד ת\"א-263", "business_code": "PISR00010635"},
          "STISR0121023": {"branch_name": "מרכז שילוח כ\"ס הצומת-264", "business_code": "PISR00010633"},
          "STISR0120562": {"branch_name": "אקספרס חולון - זלמן ארן-265", "business_code": "PISR00010635"},
          "STISR0120262": {"branch_name": "שלי התמרים אילת-266", "business_code": "PISR00010555"},
          "STISR0121116": {"branch_name": "דיל קצרין- חרמון-267", "business_code": "PISR00010633"},
          "STISR0121921": {"branch_name": "דיל אקסטרה פ\"ת- סגולה-269", "business_code": "PISR00011014"},
          "STISR0120563": {"branch_name": "חולון-הפלד Be-270", "business_code": "PISR00010668"},
          "STISR0121508": {"branch_name": "שלי רחובות- יעקובי-271", "business_code": "PISR00010555"},
          "STISR0121519": {"branch_name": "שלי גבעתיים- ויצמן-272", "business_code": "PISR00010555"},
          "STISR0121196": {"branch_name": "אקספרס ת\"א- פלורנטין-274", "business_code": "PISR00010635"},
          "STISR0120263": {"branch_name": "דיל אמט\"ל רחובות-276", "business_code": "PISR00010633"},
          "STISR0120396": {"branch_name": "אקספרס.בני בנימין הרצליה-277", "business_code": "PISR00010635"},
          "STISR0120547": {"branch_name": "שלי פ\"ת -העצמאות-278", "business_code": "PISR00010555"},
          "STISR0121198": {"branch_name": "אקספרס-ת\"א דרויאנוב-279", "business_code": "PISR00010635"},
          "STISR0121946": {"branch_name": "דיל חדרה- צה\"ל-280", "business_code": "PISR00010633"},
          "STISR0121935": {"branch_name": "דיל מגדל העמק- המדע-281", "business_code": "PISR00010633"},
          "STISR0120307": {"branch_name": "דיל צומת גבעת מרדכי-282", "business_code": "PISR00010633"},
          "STISR0121367": {"branch_name": "דיל קרית עקרון- ביל\"ו-283", "business_code": "PISR00010633"},
          "STISR0120308": {"branch_name": "שלי תל מונד- הדקל-284", "business_code": "PISR00010555"},
          "STISR0120390": {"branch_name": "אקספרס סוקולוב פ\"ת-285", "business_code": "PISR00010635"},
          "STISR0120666": {"branch_name": "שלי גבעתיים-שביט-287", "business_code": "PISR00010555"},
          "STISR0121566": {"branch_name": "יש רחובות- סירני-288", "business_code": "PISR00010634"},
          "STISR0121136": {"branch_name": "דיל פ\"ת- יכין סנטר-290", "business_code": "PISR00010633"},
          "STISR0122156": {"branch_name": "מחסן אונליין יבנה הירמוך-292", "business_code": "PISR00010633"},
          "STISR0120391": {"branch_name": "אקספרס צור הדסה-293", "business_code": "PISR00010635"},
          "STISR0121957": {"branch_name": "מחסן אונליין פ\"ת בזל-294", "business_code": "PISR00010633"},
          "STISR0121117": {"branch_name": "יש חסד בני ברק-שלמה המלך-295", "business_code": "PISR00010634"},
          "STISR0120564": {"branch_name": "אקספרס פ\"ת רחל המשוררת-296", "business_code": "PISR00010635"},
          "STISR0121947": {"branch_name": "דיל חדרה- ארבע האגודות-297", "business_code": "PISR00010633"},
          "STISR0120392": {"branch_name": "אמות באר שבע Be-298", "business_code": "PISR00010668"},
          "STISR0120565": {"branch_name": "אקספרס הוד השרון-299", "business_code": "PISR00010635"},
          "STISR0119401": {"branch_name": "יש חסד בארות יצחק- פאוור-300", "business_code": "PISR00010634"},
          "STISR0135166": {"branch_name": "שופרסל אקספרס אחי מאיר-301", "business_code": "PISR00010635"},
          "STISR0121948": {"branch_name": "שלי חיפה- קלר-302", "business_code": "PISR00010555"},
          "STISR0121143": {"branch_name": "שלי טבעון- יהודה הנשיא-303", "business_code": "PISR00010555"},
          "STISR0121144": {"branch_name": "אקספרס קריית חיים-304", "business_code": "PISR00010635"},
          "STISR0120469": {"branch_name": "אקספרס ת\"א- יהודה המכבי-305", "business_code": "PISR00010635"},
          "STISR0119957": {"branch_name": "יש נוה שאנן חיפה- חניתה-306", "business_code": "PISR00010634"},
          "STISR0121145": {"branch_name": "יש קרית אתא- זבולון-307", "business_code": "PISR00010634"},
          "STISR0121188": {"branch_name": "שלי קרית אתא- איינשטין-308", "business_code": "PISR00010555"},
          "STISR0121370": {"branch_name": "תל השומר Be-309", "business_code": "PISR00010668"},
          "STISR0120542": {"branch_name": "שלי קרית חיים- דגניה-310", "business_code": "PISR00010555"},
          "STISR0120397": {"branch_name": "אקספרס רחובות- בוסתנאי-311", "business_code": "PISR00010635"},
          "STISR0121137": {"branch_name": "דיל חיפה- קרית אליעזר-312", "business_code": "PISR00010633"},
          "STISR0120667": {"branch_name": "שלי קרית חיים-אח\"י אילת-313", "business_code": "PISR00010555"},
          "STISR0120543": {"branch_name": "יש  חיפה- צרפת-314", "business_code": "PISR00010634"},
          "STISR0121111": {"branch_name": "אקספרס כפר נטר-315", "business_code": "PISR00010635"},
          "STISR0121417": {"branch_name": "שלי קרית מוצקין-שד ויצמן-316", "business_code": "PISR00010555"},
          "STISR0121375": {"branch_name": "אקספרס גבעת עדה-317", "business_code": "PISR00010635"},
          "STISR0120566": {"branch_name": "אקספרס חולון - וולפסון-318", "business_code": "PISR00010635"},
          "STISR0121382": {"branch_name": "שלי עכו- שפירא-319", "business_code": "PISR00010555"},
          "STISR0143538": {"branch_name": "דיל אשדוד האגס-320", "business_code": "PISR00010633"},
          "STISR0120393": {"branch_name": "גן העיר ת\"א Be-322", "business_code": "PISR00010668"},
          "STISR0120668": {"branch_name": "שלי רסקו נצרת עילית-הדקל-323", "business_code": "PISR00010555"},
          "STISR0119286": {"branch_name": "אקספרס סביון- השקמה-324", "business_code": "PISR00010635"},
          "STISR0120548": {"branch_name": "אקספרס ששת הימים נס ציונ-325", "business_code": "PISR00010635"},
          "STISR0121146": {"branch_name": "שלי נהריה- הגעתון-326", "business_code": "PISR00010555"},
          "STISR0121578": {"branch_name": "אקספרס חיפה- ראול ולנברג-327", "business_code": "PISR00010635"},
          "STISR0121376": {"branch_name": "אקספרס האילנות-328", "business_code": "PISR00010635"},
          "STISR0120549": {"branch_name": "יש יוקנעם עילית-היובלים-329", "business_code": "PISR00010634"},
          "STISR0121512": {"branch_name": "שלי חיפה- נתיב חן-330", "business_code": "PISR00010555"},
          "STISR0121567": {"branch_name": "אקספרס ת\"א- תוספתא-331", "business_code": "PISR00010635"},
          "STISR0120793": {"branch_name": "שלי חיפה- אורן-333", "business_code": "PISR00010555"},
          "STISR0120567": {"branch_name": "אקספרס פ\"ת - שפירא-334", "business_code": "PISR00010635"},
          "STISR0121594": {"branch_name": "אקספרס כ\"ס- בן גוריון-335", "business_code": "PISR00010635"},
          "STISR0120550": {"branch_name": "שלי חיפה- רמות ספיר-336", "business_code": "PISR00010555"},
          "STISR0122157": {"branch_name": "מחסן און ליין כנות-337", "business_code": "PISR00010633"},
          "STISR0121595": {"branch_name": "אקספרס כוכב הצפון-338", "business_code": "PISR00010635"},
          "STISR0120116": {"branch_name": "דיל טבריה- העמקים-339", "business_code": "PISR00010633"},
          "STISR0121596": {"branch_name": "אקספרס ת\"א- זריצקי-340", "business_code": "PISR00010635"},
          "STISR0120669": {"branch_name": "אקספרס רמת גן- הרוא\"ה-341", "business_code": "PISR00010635"},
          "STISR0120500": {"branch_name": "דיל כרמיאל- קניון לב-342", "business_code": "PISR00010633"},
          "STISR0120501": {"branch_name": "שלי ראש העין- ברקן-343", "business_code": "PISR00010555"},
          "STISR0120264": {"branch_name": "דיל מעלות- דרך האלוף עוז-344", "business_code": "PISR00010633"},
          "STISR0120645": {"branch_name": "מידטאון ת\"א Be-345", "business_code": "PISR00010668"},
          "STISR0120117": {"branch_name": "גרין קרית אונו-ירמיהו-346", "business_code": "PISR00011013"},
          "STISR0121206": {"branch_name": "אקספרס גדרה- דרך הפרחים-347", "business_code": "PISR00010635"},
          "STISR0121373": {"branch_name": "אקספרס פ\"ת- קק\"ל-348", "business_code": "PISR00010635"},
          "STISR0121147": {"branch_name": "יש חסד צפת- כנען-349", "business_code": "PISR00010634"},
          "STISR0120544": {"branch_name": "דיל גבעת אולגה-הרב ניסים-350", "business_code": "PISR00010633"},
          "STISR0139479": {"branch_name": "בי טייבה-352", "business_code": "PISR00010668"},
          "STISR0121199": {"branch_name": "אקספרס חיפה יותם-353", "business_code": "PISR00010635"},
          "STISR0121597": {"branch_name": "אקספרס ת\"א- קינג ג'ורג-354", "business_code": "PISR00010635"},
          "STISR0120118": {"branch_name": "גרין ת\"א- אהרון בקר-355", "business_code": "PISR00011013"},
          "STISR0119112": {"branch_name": "שלי צורן קדימה- לב השרון-357", "business_code": "PISR00010555"},
          "STISR0121118": {"branch_name": "דיל חיפה- גרנד קניון-359", "business_code": "PISR00010633"},
          "STISR0121580": {"branch_name": "אקספרס מאירוביץ ראשל\"צ-360", "business_code": "PISR00010635"},
          "STISR0121509": {"branch_name": "דיל מגדל העמק- א ת דרומי-361", "business_code": "PISR00010633"},
          "STISR0120794": {"branch_name": "אקספרס קרן היסוד ירושלים-362", "business_code": "PISR00010635"},
          "STISR0136378": {"branch_name": "יש חסד בית וגן-364", "business_code": "PISR00010634"},
          "STISR0120260": {"branch_name": "שלי פרדס חנה כרכור- קדמה-365", "business_code": "PISR00010555"},
          "STISR0120551": {"branch_name": "שלי כפר ורדים- מרכז מסחר-366", "business_code": "PISR00010555"},
          "STISR0121618": {"branch_name": "שרי ישראל  ירושלים Be-367", "business_code": "PISR00010668"},
          "STISR0121619": {"branch_name": "אקספרס-נשר מרגנית-368", "business_code": "PISR00010635"},
          "STISR0120795": {"branch_name": "שלי שוהם- תרשיש-369", "business_code": "PISR00010555"},
          "STISR0120502": {"branch_name": "שלי ברניצקי ראשון-371", "business_code": "PISR00010555"},
          "STISR0120545": {"branch_name": "אקספרס בת ים- חשמונאים-372", "business_code": "PISR00010635"},
          "STISR0120670": {"branch_name": "שלי הרצליה- הבנים-374", "business_code": "PISR00010555"},
          "STISR0121620": {"branch_name": "אקספרס כ\"ס- הגליל-375", "business_code": "PISR00010635"},
          "STISR0120552": {"branch_name": "אקספרס רעננה- משה דיין-376", "business_code": "PISR00010635"},
          "STISR0120796": {"branch_name": "אקספרס ירושלים- הדסה-377", "business_code": "PISR00010635"},
          "STISR0122155": {"branch_name": "אקספרס עמק רפאים ירושלים-378", "business_code": "PISR00010635"},
          "STISR0121598": {"branch_name": "אקספרס גבעתיים- כצנלסון-379", "business_code": "PISR00010635"},
          "STISR0120399": {"branch_name": "אקספרס ת\"א- מרמורק-380", "business_code": "PISR00010635"},
          "STISR0121520": {"branch_name": "אקספרס פ\"ת- היבנר-382", "business_code": "PISR00010635"},
          "STISR0120671": {"branch_name": "אקספרס קרית ביאליק-384", "business_code": "PISR00010635"},
          "STISR0120672": {"branch_name": "אקספרס מאיר יערי ת\"א-385", "business_code": "PISR00010635"},
          "STISR0121207": {"branch_name": "אקספרס רמב\"ם חיפה-387", "business_code": "PISR00010635"},
          "STISR0121371": {"branch_name": "אקספרס סביוני דניה-388", "business_code": "PISR00010635"},
          "STISR0120265": {"branch_name": "שלי בת חפר-389", "business_code": "PISR00010555"},
          "STISR0119402": {"branch_name": "אקספרס פ\"ת מרכז רום-390", "business_code": "PISR00010635"},
          "STISR0121200": {"branch_name": "אקספרס אלנקווה-391", "business_code": "PISR00010635"},
          "STISR0121958": {"branch_name": "בי דראגסטורס שילת-392", "business_code": "PISR00010668"},
          "STISR0121687": {"branch_name": "אקספרס להבים באר שבע-393", "business_code": "PISR00010635"},
          "STISR0119436": {"branch_name": "אקספרס כפר סבא הגליל-394", "business_code": "PISR00010635"},
          "STISR0121418": {"branch_name": "אקס רחובות לוין אפשטיין-395", "business_code": "PISR00010635"},
          "STISR0129023": {"branch_name": "שופרסל דיל בקדר נתניה-396", "business_code": "PISR00010633"},
          "STISR0120993": {"branch_name": "אקספרס עפולה פארק-397", "business_code": "PISR00010635"},
          "STISR0137393": {"branch_name": "אקספרס אגמים-398", "business_code": "PISR00010635"},
          "STISR0120553": {"branch_name": "דיל עכו-399", "business_code": "PISR00010633"},
          "STISR0121627": {"branch_name": "אקספרס שבזי ראש העין-400", "business_code": "PISR00010635"},
          "STISR0146612": {"branch_name": "שופרסל סטוק חולון-435", "business_code": "PISR00028200"},
          "STISR0121419": {"branch_name": "אייס מול אילת Be-437", "business_code": "PISR00010668"},
          "STISR0121949": {"branch_name": "גרין הרצליה- הנגב-438", "business_code": "PISR00011013"},
          "STISR0121510": {"branch_name": "גרין רחובות- יעקב-440", "business_code": "PISR00011013"},
          "STISR0120248": {"branch_name": "גרין ת\"א- שינקין-441", "business_code": "PISR00011013"},
          "STISR0121208": {"branch_name": "אקספרס באר שבע אביסרור-444", "business_code": "PISR00010635"},
          "STISR0121372": {"branch_name": "אקספרס רחובות יעקב-445", "business_code": "PISR00010635"},
          "STISR0121959": {"branch_name": "חנות גדרון-חולון-446", "business_code": "PISR00010633"},
          "STISR0121960": {"branch_name": "חנות גדרון פתח תקוה-447", "business_code": "PISR00010633"},
          "STISR0121621": {"branch_name": "שערי רווחה-448", "business_code": "PISR00011015"},
          "STISR0121961": {"branch_name": "חנות גדרון צומת הנגב-449", "business_code": "PISR00010633"},
          "STISR0131712": {"branch_name": "בי דראגסטור ככר המושבה הוד השרון-476", "business_code": "PISR00010668"},
          "STISR0126808": {"branch_name": "אקספרס הציפורים מודיעין-477", "business_code": "PISR00010635"},
          "STISR0125386": {"branch_name": "שופרסל אקספרס גדרה-478", "business_code": "PISR00010635"},
          "STISR0122110": {"branch_name": "אקספרס בר אילן-479", "business_code": "PISR00010635"},
          "STISR0123540": {"branch_name": "אקספרס מנדלה מוכר ספרים-481", "business_code": "PISR00010635"},
          "STISR0133345": {"branch_name": "בי דראגסטורס נווה דורון רמלה-482", "business_code": "PISR00010668"},
          "STISR0121615": {"branch_name": "בלה אשדוד-485", "business_code": "PISR00010668"},
          "STISR0140004": {"branch_name": "אקספרס מורשת מודיעין-489", "business_code": "PISR00010635"},
          "STISR0121152": {"branch_name": "יש חסד כנפי נשרים-496", "business_code": "PISR00010634"},
          "STISR0134932": {"branch_name": "יש חסד עפולה-499", "business_code": "PISR00010634"},
          "STISR0151558": {"branch_name": "אקספרס רסקו נצרת-511", "business_code": "PISR00010635"},
          "STISR0152030": {"branch_name": "כפר קרע Be-518", "business_code": "PISR00010668"},
          "STISR0136731": {"branch_name": "אקספרס סביון-579", "business_code": "PISR00010635"},
          "STISR0140005": {"branch_name": "יש חסד אשדוד-592", "business_code": "PISR00010634"},
          "STISR0140007": {"branch_name": "יש חסד באר טוביה-593", "business_code": "PISR00010634"},
          "STISR0142789": {"branch_name": "אקספרס בת חפר-595", "business_code": "PISR00010635"},
          "STISR0137699": {"branch_name": "אקספרס אבן גבירול -596", "business_code": "PISR00010635"},
          "STISR0139480": {"branch_name": "אקספרס כפר ורדים-597", "business_code": "PISR00010635"},
          "STISR0138242": {"branch_name": "אקספרס הירדן-598", "business_code": "PISR00010635"},
          "STISR0138243": {"branch_name": "אקספרס ההסתדרות-599", "business_code": "PISR00010635"},
          "STISR0121148": {"branch_name": "פרדס חנה Be-601", "business_code": "PISR00010668"},
          "STISR0121599": {"branch_name": "אשקלון מרינה Be-602", "business_code": "PISR00010668"},
          "STISR0121428": {"branch_name": "יש חסד בית שמש- מ. מסחרי-606", "business_code": "PISR00010634"},
          "STISR0121228": {"branch_name": "יש חסד מודיעין עלי-ק.ספר-607", "business_code": "PISR00010634"},
          "STISR0121119": {"branch_name": "יש חסד ירושלים-גבעת שאול-608", "business_code": "PISR00010634"},
          "STISR0121511": {"branch_name": "יש חסד אלעד-ניסים גאון-609", "business_code": "PISR00010634"},
          "STISR0121427": {"branch_name": "יש חסד ירושלים- רמת שלמה-610", "business_code": "PISR00010634"},
          "STISR0121429": {"branch_name": "יש חסד בני ברק- השומר-611", "business_code": "PISR00010634"},
          "STISR0120797": {"branch_name": "אשדוד Be-612", "business_code": "PISR00010668"},
          "STISR0121210": {"branch_name": "נהריה Be-613", "business_code": "PISR00010668"},
          "STISR0120646": {"branch_name": "נתניה Be-614", "business_code": "PISR00010668"},
          "STISR0119366": {"branch_name": "כפר סבא Be-615", "business_code": "PISR00010668"},
          "STISR0121377": {"branch_name": "תחנה מרכזית ת\"א Be-616", "business_code": "PISR00010668"},
          "STISR0120568": {"branch_name": "נצרת Be-617", "business_code": "PISR00010668"},
          "STISR0121202": {"branch_name": "ראשל'צ Be-618", "business_code": "PISR00010668"},
          "STISR0121420": {"branch_name": "אילת Be-620", "business_code": "PISR00010668"},
          "STISR0120798": {"branch_name": "סינמה סיטי נתניה Be-621", "business_code": "PISR00010668"},
          "STISR0119403": {"branch_name": "סביונים Be-623", "business_code": "PISR00010668"},
          "STISR0120400": {"branch_name": "ולפסון Be-631", "business_code": "PISR00010668"},
          "STISR0120994": {"branch_name": "נשר תל-חנן Be-633", "business_code": "PISR00010668"},
          "STISR0120673": {"branch_name": "בת ים Be-634", "business_code": "PISR00010668"},
          "STISR0121211": {"branch_name": "ראש-פינה Be-635", "business_code": "PISR00010668"},
          "STISR0121203": {"branch_name": "רעננה Be-637", "business_code": "PISR00010668"},
          "STISR0120249": {"branch_name": "משמר השרון Be-638", "business_code": "PISR00010668"},
          "STISR0120647": {"branch_name": "אריאל Be-639", "business_code": "PISR00010668"},
          "STISR0121378": {"branch_name": "דלית אל כרמל Be-640", "business_code": "PISR00010668"},
          "STISR0120648": {"branch_name": "גבעתיים Be-641", "business_code": "PISR00010668"},
          "STISR0121622": {"branch_name": "אלונים-טבעון Be-642", "business_code": "PISR00010668"},
          "STISR0120401": {"branch_name": "סירקין פתח תקוה Be-643", "business_code": "PISR00010668"},
          "STISR0120402": {"branch_name": "גבעת שמואל Be-645", "business_code": "PISR00010668"},
          "STISR0121421": {"branch_name": "G  באר-שבע Be-646", "business_code": "PISR00010668"},
          "STISR0121430": {"branch_name": "חולון Be-647", "business_code": "PISR00010668"},
          "STISR0121379": {"branch_name": "מאיר-כפר סבא Be-648", "business_code": "PISR00010668"},
          "STISR0121381": {"branch_name": "עכו Be-649", "business_code": "PISR00010668"},
          "STISR0120554": {"branch_name": "חיפה Be-650", "business_code": "PISR00010668"},
          "STISR0120298": {"branch_name": "דרורים Be-651", "business_code": "PISR00010668"},
          "STISR0121568": {"branch_name": "מודיעין Be-652", "business_code": "PISR00010668"},
          "STISR0121582": {"branch_name": "בני-ברק Be-653", "business_code": "PISR00010668"},
          "STISR0120649": {"branch_name": "אם המושבות Be-654", "business_code": "PISR00010668"},
          "STISR0121569": {"branch_name": "רחובות מזרח Be-655", "business_code": "PISR00010668"},
          "STISR0120503": {"branch_name": "בן יהודה Be-656", "business_code": "PISR00010668"},
          "STISR0121212": {"branch_name": "קריית שמונה Be-657", "business_code": "PISR00010668"},
          "STISR0121588": {"branch_name": "קריית אתא Be-658", "business_code": "PISR00010668"},
          "STISR0121213": {"branch_name": "הדר Be-659", "business_code": "PISR00010668"},
          "STISR0120555": {"branch_name": "רמת השרון Be-660", "business_code": "PISR00010668"},
          "STISR0120465": {"branch_name": "ב.חולים אשדוד Be-662", "business_code": "PISR00010668"},
          "STISR0120569": {"branch_name": "פלורנטין Be-663", "business_code": "PISR00010668"},
          "STISR0121600": {"branch_name": "שערי צדק Be-664", "business_code": "PISR00010668"},
          "STISR0120799": {"branch_name": "חשמונאים Be-665", "business_code": "PISR00010668"},
          "STISR0121614": {"branch_name": "מודיעין עזריאלי Be-666", "business_code": "PISR00010668"},
          "STISR0121591": {"branch_name": "מדיקל הלל  חדרה Be-667", "business_code": "PISR00010668"},
          "STISR0120466": {"branch_name": "קריית השרון נתניה Be-668", "business_code": "PISR00010668"},
          "STISR0121214": {"branch_name": "איכילוב Be-670", "business_code": "PISR00010668"},
          "STISR0121570": {"branch_name": "אשדוד סיטי Be-672", "business_code": "PISR00010668"},
          "STISR0120470": {"branch_name": "פנורמה חיפה Be-673", "business_code": "PISR00010668"},
          "STISR0120674": {"branch_name": "תל ברוך Be-674", "business_code": "PISR00010668"},
          "STISR0120650": {"branch_name": "שרונים Be-675", "business_code": "PISR00010668"},
          "STISR0120297": {"branch_name": "יקנעם Be-676", "business_code": "PISR00010668"},
          "STISR0121592": {"branch_name": "הרצליה Be-677", "business_code": "PISR00010668"},
          "STISR0121204": {"branch_name": "עפולה Be-678", "business_code": "PISR00010668"},
          "STISR0121623": {"branch_name": "כרכור Be-679", "business_code": "PISR00010668"},
          "STISR0121601": {"branch_name": "שוהם Be-681", "business_code": "PISR00010668"},
          "STISR0119437": {"branch_name": "כפר סבא הירוקה Be-682", "business_code": "PISR00010668"},
          "STISR0121571": {"branch_name": "אשקלון  דוידי Be-683", "business_code": "PISR00010668"},
          "STISR0121624": {"branch_name": "שדרות Be-684", "business_code": "PISR00010668"},
          "STISR0121602": {"branch_name": "אזורי חן Be-685", "business_code": "PISR00010668"},
          "STISR0120467": {"branch_name": "מדיקל-אזור Be-686", "business_code": "PISR00010668"},
          "STISR0120471": {"branch_name": "מדיקל-רחובות Be-687", "business_code": "PISR00010668"},
          "STISR0121607": {"branch_name": "מדיקל-בלינסון Be-688", "business_code": "PISR00010668"},
          "STISR0121613": {"branch_name": "מדיקל גדרה  Be-689", "business_code": "PISR00010668"},
          "STISR0121232": {"branch_name": "אשדוד מדיקל Be-691", "business_code": "PISR00010668"},
          "STISR0120394": {"branch_name": "אסותא רמת החייל Be-695", "business_code": "PISR00010668"},
          "STISR0120570": {"branch_name": "אלעד Be-696", "business_code": "PISR00010668"},
          "STISR0121380": {"branch_name": "ירכא Be-697", "business_code": "PISR00010668"},
          "STISR0120675": {"branch_name": "כרמיאל Be-698", "business_code": "PISR00010668"},
          "STISR0131728": {"branch_name": "שופרסל אילת-709", "business_code": "PISR00010633"},
          "STISR0135076": {"branch_name": "בי דראגסטורס עמק סנטר עפולה-711", "business_code": "PISR00010668"},
          "STISR0139481": {"branch_name": "בי עין הים חדרה-712", "business_code": "PISR00010668"},
          "STISR0135785": {"branch_name": "בי דראגסטור אור עקיבא (אור הים)-713", "business_code": "PISR00010668"},
          "STISR0135786": {"branch_name": "בי דראגסטור סכנין-714", "business_code": "PISR00010668"},
          "STISR0138894": {"branch_name": "BE אגמים אשקלון-715", "business_code": "PISR00010668"},
          "STISR0141459": {"branch_name": "be mall 7-716", "business_code": "PISR00010668"},
          "STISR0146614": {"branch_name": "שופרסל סטוק עכו-719", "business_code": "PISR00028200"},
          "STISR0146615": {"branch_name": "שופרסל סטוק יבנה-720", "business_code": "PISR00028200"},
          "STISR0146616": {"branch_name": "שופרסל סטוק עפולה-721", "business_code": "PISR00028200"},
          "STISR0146617": {"branch_name": "שופרסל סטוק קריית חיים-722", "business_code": "PISR00028200"},
          "STISR0146618": {"branch_name": "שופרסל סטוק באר שבע-723", "business_code": "PISR00028200"},
          "STISR0146619": {"branch_name": "שופרסל סטוק רמת גן-724", "business_code": "PISR00028200"},
          "STISR0146620": {"branch_name": "שופרסל סטוק פתח תקווה-725", "business_code": "PISR00028200"},
          "STISR0146621": {"branch_name": "שופרסל סטוק נס ציונה-726", "business_code": "PISR00028200"},
          "STISR0146622": {"branch_name": "שופרסל סטוק תל אביב-727", "business_code": "PISR00028200"},
          "STISR0146623": {"branch_name": "שופרסל סטוק טירת הכרמל-728", "business_code": "PISR00028200"},
          "STISR0146624": {"branch_name": "שופרסל סטוק פרדס חנה-729", "business_code": "PISR00028200"},
          "STISR0135086": {"branch_name": "מחסן מרלו\"ג קדימה-730", "business_code": "PISR00010633"},
          "STISR0146613": {"branch_name": "שופרסל סטוק מודיעין-731", "business_code": "PISR00028200"},
          "STISR0146625": {"branch_name": "שופרסל סטוק צ'ק פוסט-732", "business_code": "PISR00028200"},
          "STISR0144626": {"branch_name": "אקספרס אילת-733", "business_code": "PISR00010635"},
          "STISR0142178": {"branch_name": "גוד מרקט גאולה חיפה-734", "business_code": "PISR00013665"},
          "STISR0140078": {"branch_name": "CASH&CARRY באר שבע-737", "business_code": "PISR00011831"},
          "STISR0138800": {"branch_name": "אקספרס קריית אונו-740", "business_code": "PISR00010635"},
          "STISR0138801": {"branch_name": "אקספרס קריית אונו-741", "business_code": "PISR00010635"},
          "STISR0138534": {"branch_name": "אקספרס נווה אביבים-742", "business_code": "PISR00010635"},
          "STISR0139482": {"branch_name": "אקספרס קריית גת-743", "business_code": "PISR00010635"},
          "STISR0143585": {"branch_name": "אקספרס אשקלון נחל צאלים-744", "business_code": "PISR00010635"},
          "STISR0136401": {"branch_name": "שופרסל בע\"מ (נייד)-746", "business_code": "PISR00010633"},
          "STISR0136673": {"branch_name": "יש חסד אג\"מ (נייד)-747", "business_code": "PISR00010634"},
          "STISR0139483": {"branch_name": "בי  שבעת הכוכבים אילת-748", "business_code": "PISR00010668"},
          "STISR0135632": {"branch_name": "גוד מרקט שירת הים ירושלים-750", "business_code": "PISR00013665"},
          "STISR0152543": {"branch_name": "גוד מרקט בני ברק-751", "business_code": "PISR00013665"},
          "STISR0135077": {"branch_name": "שופרסל אקספרס דוכיפת כפר סבא-752", "business_code": "PISR00010635"},
          "STISR0135633": {"branch_name": "גוד מרקט רש\"י ירושלים-753", "business_code": "PISR00013665"},
          "STISR0150373": {"branch_name": "שופרסל שלי באר יעקב-756", "business_code": "PISR00010555"},
          "STISR0135167": {"branch_name": "שופרסל אקספרס קריית אונו-757", "business_code": "PISR00010635"},
          "STISR0151724": {"branch_name": "שופרסל שלי ראש העין-760", "business_code": "PISR00010555"},
          "STISR0129612": {"branch_name": "בי דראגסטורס אושיסקין רמת השרון-762", "business_code": "PISR00010668"},
          "STISR0131713": {"branch_name": "בי דראגסטור שפרעם-763", "business_code": "PISR00010668"},
          "STISR0131714": {"branch_name": "בי דראגסטור סטאר סנטר אשדוד-764", "business_code": "PISR00010668"},
          "STISR0136302": {"branch_name": "יש חסד בן איש חי-בית שמש-765", "business_code": "PISR00010634"},
          "STISR0130534": {"branch_name": "שופרסל אקספרס רמת גן-768", "business_code": "PISR00010635"},
          "STISR0131715": {"branch_name": "גוד מרקט בני ברק-771", "business_code": "PISR00013665"},
          "STISR0132317": {"branch_name": "גוט מרקט מאה שערים-772", "business_code": "PISR00013665"},
          "STISR0126901": {"branch_name": "גוד מרקט אהבת שלום בית שמש-773", "business_code": "PISR00013665"},
          "STISR0121149": {"branch_name": "דראגסטורס טירה Be-776", "business_code": "PISR00010668"},
          "STISR0134188": {"branch_name": "שופרסל אקספרס נרקיסים-777", "business_code": "PISR00010635"},
          "STISR0124215": {"branch_name": "גוד מרקט אהליאב ירושלים-778", "business_code": "PISR00013665"},
          "STISR0125388": {"branch_name": "שופרסל אקספרס יבנה-779", "business_code": "PISR00010635"},
          "STISR0126869": {"branch_name": "שופרסל אקספרס המעפיל רמת גן-780", "business_code": "PISR00010635"},
          "STISR0124857": {"branch_name": "בי דראגסטורס פרישמן תל אביב-781", "business_code": "PISR00010668"},
          "STISR0127135": {"branch_name": "בי דראגסטורס שיין ראשון לציון-782", "business_code": "PISR00010668"},
          "STISR0124858": {"branch_name": "בי דראגסטורס מודיעין-783", "business_code": "PISR00010668"},
          "STISR0124216": {"branch_name": "דראגסטורס כפר גנים פתח תקוה Be-784", "business_code": "PISR00010668"},
          "STISR0126870": {"branch_name": "בי דראגסטורס מעלה אדומים -785", "business_code": "PISR00010668"},
          "STISR0133346": {"branch_name": "בי דראגסטורס כרמי גת-786", "business_code": "PISR00010668"},
          "STISR0133347": {"branch_name": "בי דראגסטורס כיכר הגבורה-787", "business_code": "PISR00010668"},
          "STISR0129613": {"branch_name": "בי דראגסטורס אפרת-788", "business_code": "PISR00010668"},
          "STISR0130536": {"branch_name": "גוד מרקט אופקים-789", "business_code": "PISR00013665"},
          "STISR0133062": {"branch_name": "שופרסל דיל רמלה -790", "business_code": "PISR00010633"},
          "STISR0130384": {"branch_name": "שופרסל בע\"מ מכירת תווים-793", "business_code": "PISR00018750"},
          "STISR0130385": {"branch_name": "שופרסל קופות ראשיות-795", "business_code": "PISR00018750"},
          "STISR0130386": {"branch_name": "שופרסל שירותי קופה- מ. מזומן-798", "business_code": "PISR00018750"},
          "STISR0130387": {"branch_name": "בי דראגסטורס משיכת מזומן-800", "business_code": "PISR00018750"},
          "STISR0132782": {"branch_name": "בי דראגסטורס תרופות-801", "business_code": "PISR00018750"},
          "STISR0140087": {"branch_name": "מחסן אונליין BE לונג-824", "business_code": "PISR00010668"},
          "STISR0141108": {"branch_name": "גוד מרקט חריש-840", "business_code": "PISR00013665"},
          "STISR0133715": {"branch_name": "שופרסל אקספרס מטודלה תל אביב-842", "business_code": "PISR00010635"},
          "STISR0135634": {"branch_name": "גוד מרקט בית שמש-843", "business_code": "PISR00013665"},
          "STISR0127335": {"branch_name": "שופרסל אקספרס נגבה רמת גן-844", "business_code": "PISR00010635"},
          "STISR0126981": {"branch_name": "שופרסל אקספרס נחשון כפר סבא-845", "business_code": "PISR00010635"},
          "STISR0141189": {"branch_name": "שופרסל אקספרס אור ים-846", "business_code": "PISR00010635"},
          "STISR0140676": {"branch_name": "שופרסל דיל גבעת אלונים-854", "business_code": "PISR00010633"},
          "STISR0133717": {"branch_name": "שופרסל מחסן מרלו\"ג מודיעין-860", "business_code": "PISR00010633"},
          "STISR0122822": {"branch_name": "שופרסל בע\"מ  (משרד גביית כרטיסי אשראי)-920", "business_code": "PISR00010633"}
    }
    print("מילון BRANCH_DB נטען בהצלחה!")
except Exception as e:
    print("שגיאה בטעינת BRANCH_DB:", e)

# ================== פונקציה לסניף ברירת מחדל ==================
def get_default_site(account):
    """מחפש את הסניף הראשון ב-BRANCH_DB ששייך לחשבון (business_code)"""
    for site_id, info in BRANCH_DB.items():
        if info.get("business_code") == account:
            return site_id
            
    # אם לא נמצא ב-DB, נשתמש במיפוי ידני כגיבוי
    default_sites = {
        "PISR00010633": "STISR0119905",
        "PISR00010634": "STISR0119435",
        "PISR00010555": "STISR0119426",
        "PISR00010635": "STISR0131035",
        "PISR00010668": "STISR0121575",
        "PISR00011014": "STISR0120310",
        "PISR00028200": "STISR0148501",
        "PISR00011831": "STISR0121616",
        "PISR00011015": "STISR0121593",
        "PISR00011013": "STISR0121117",
        "PISR00013665": "STISR0142178",
        "PISR00018750": "STISR0130384",
    }
    return default_sites.get(account, "000000")

# ===== הגדרות תיקיות =====
LOG_DIR = r"C:\Users\moshei1\OneDrive - Verifone\Desktop\TIP\STFPNOW"
CSV_DIR = os.path.join(LOG_DIR, "csv")
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(CSV_DIR, exist_ok=True)

log_filename = os.path.join(LOG_DIR, f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
log_file = open(log_filename, "a", encoding="utf-8")

# ================== פונקציות לוג ==================
def log_write(text_widget, message):
    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S] ")
    full_msg = timestamp + message
    if text_widget:
        text_widget.insert("end", full_msg + "\n")
        text_widget.see("end")
    log_file.write(full_msg + "\n")
    log_file.flush()
    os.fsync(log_file.fileno())
    print(full_msg)


# ================== API ==================
def get_token(user_id, password, text_widget):
    url = "https://us.vfmerchantportal.com/gmp-web/rest/public/security/login"
    headers = {"Content-Type": "application/json"}
    payload = {"userId": user_id, "password": password}
    resp = requests.post(url, headers=headers, json=payload)
    if resp.status_code != 200:
        log_write(text_widget, f"❌ Login failed: {resp.status_code}")
        return None, None
    data = resp.json()
    log_write(text_widget, "✅ Token received")
    return data["bearerToken"], resp.cookies.get_dict()

def get_failed_batches(entity_id, token, cookies, text_widget):
    url = "https://us.vfmerchantportal.com/gmp-web/rest/settlement/execution"
    headers = {"Authorization": f"Bearer {token}"}

    # 📅 טווח יומיים אחורה: מה-יומיים לפני אתמול ועד אתמול
    today = datetime.today()
    date_to = (today - timedelta(days=1)).strftime("%Y%m%d")   # עד אתמול
    date_from = (today - timedelta(days=2)).strftime("%Y%m%d") # יומיים אחורה

    params = {
        "dateFrom": date_from,
        "dateTo": date_to,
        "entityId": entity_id,
        "entityType": "MERCHANT",
        "submissionFlag": "STFP_FAILED",
        "pageSize": 100
    }

    resp = requests.get(url, headers=headers, params=params, cookies=cookies)
    log_write(text_widget, f"[GET BATCHES] {resp.status_code} - {entity_id} ({date_from}-{date_to})")

    if resp.status_code != 200:
        return []

    batches = resp.json().get("list") or resp.json().get("results") or []
    result = []
    for b in batches:
        result.append({
            "site": b.get("entityId", "Unknown"),
            "account": entity_id,
            "totalValue": b.get("totalValue", 0),
            "shiftId": b.get("shiftId", 0)
        })

    # מניעת כפילויות לפי site
    unique_sites = {}
    for b in result:
        if b["site"] not in unique_sites:
            unique_sites[b["site"]] = b

    return list(unique_sites.values())



def generate_report_xml(account, token, date_from, date_to, cookies=None):
    url = "https://us.vfmerchantportal.com/gmp-web/rest/reportvalidation?reportId=4431"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {
        "merchantId": account,
        "dateFrom": date_from,
        "dateTo": date_to,
        "timeFrom": "000000",
        "timeTo": "235959",
        "reportId": "4431",
        "clientId": "israel",
        "_locale": "en",
        "__csvDelimiter": ",",
        "_rtl": False
    }
    resp = requests.post(url, headers=headers, json=payload, cookies=cookies)
    if resp.status_code != 200:
        raise RuntimeError(f"❌ Failed to generate report for account {account}")
    report_token = resp.text
    return download_report(report_token, token, cookies)

def download_report(report_token, token, cookies=None):
    url = "https://us.vfmerchantportal.com/report-viewer/rest/report/sync"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {"tokenId": report_token, "format": "XML"}
    resp = requests.post(url, headers=headers, json=payload, cookies=cookies)
    if resp.status_code != 200:
        raise RuntimeError("❌ Failed to download report XML")
    return resp.text

def parse_xml_for_errors(xml_text):
    """
    מפענחת XML ומחזירה רשומות עם סטטוס:
    - "Technical Error - Settlement"
    - "Ready to Settle"
    כולל DEBUG להצגת כל סטטוסים קיימים.
    """
    try:
        root = ET.fromstring(xml_text)
    except Exception as e:
        print("❌ XML Parse Error:", str(e))
        print(xml_text[:2000])  # הדפסת התחלה של ה-XML לבדיקה
        return []

    filtered_records = []
    all_statuses = set()  # נעקוב אחרי כל סוגי הסטטוסים שיש ב-XML

    for row in root.findall(".//row"):
        status = row.findtext("aggregated_status", default="")
        all_statuses.add(status.strip())

        if status.strip() in ["Technical Error - Settlement", "Ready to Settle"]:
            filtered_records.append({
                "transaction_id": row.findtext("transaction_id"),
                "terminal_id": row.findtext("terminal_id"),
                "merchant_id": row.findtext("merchant_id"),
                "branch_id": row.findtext("branch_id"),
                "abs_transaction_type": row.findtext("abs_transaction_type"),
                "transaction_type_weaver": row.findtext("transaction_type_weaver"),
                "date": row.findtext("date"),
                "time": row.findtext("time"),
                "uid": row.findtext("uid"),
                "arn": row.findtext("arn"),
                "batch_no": row.findtext("batch_no"),
                "card_number": row.findtext("card_number"),
                "expiration_date": row.findtext("expiration_date"),
                "original_amount": row.findtext("original_amount"),
                "original_currency": row.findtext("original_currency"),
                "currency_iso_code": row.findtext("currency_iso_code"),
                "entry_mode": row.findtext("entry_mode"),
                "issuer_name": row.findtext("issuer_name"),
                "abs_terminal_id": row.findtext("abs_terminal_id"),
                "auth_number": row.findtext("auth_number"),
                "credit_type": row.findtext("credit_type"),
                "installment_details": row.findtext("installment_details"),
                "brand": row.findtext("brand"),
                "acquirer_name": row.findtext("acquirer_name"),
                "channel": row.findtext("channel"),
                "cvv_result": row.findtext("cvv_result"),
                "offline_online_authorization_indication": row.findtext("offline_online_authorization_indication"),
                "captured_by_pos_indication": row.findtext("captured_by_pos_indication"),
                "status": row.findtext("status"),
                "additional_amounts_tip_cashback": row.findtext("additional_amounts_tip_cashback"),
                "store_branch": row.findtext("store_branch"),
                "credit_terms": row.findtext("credit_terms"),
                "payment_first_amount": row.findtext("payment_first_amount"),
                "payment_not_first_amount": row.findtext("payment_not_first_amount"),
                "installment_count": row.findtext("installment_count"),
                "authorization_source": row.findtext("authorization_source"),
                "ack_number": row.findtext("ack_number"),
                "processor_stan": row.findtext("processor_stan"),
                "terminal_lane": row.findtext("terminal_lane"),
                "processor_terminal_batch_no": row.findtext("processor_terminal_batch_no"),
                "processor_acquirer_id": row.findtext("processor_acquirer_id"),
                "date_auth": row.findtext("date_auth"),
                "time_auth": row.findtext("time_auth"),
                "transaction_status": row.findtext("transaction_status"),
                "informative_tran": row.findtext("informative_tran"),
                "bad_tran": row.findtext("bad_tran"),
                "id_ext": row.findtext("id_ext"),
                "merchant_name": row.findtext("merchant_name"),
                "branch_name": row.findtext("branch_name"),
                "shift_id": row.findtext("shift_id"),
                "aggregated_status": status,
                "merchant_group": row.findtext("merchant_group"),
                "status_settlement_code": row.findtext("status_settlement_code"),
                "datetime_settle": row.findtext("datetime_settle"),
                "tid_routing_code": row.findtext("tid_routing_code"),
            })

    # 🧠 DEBUG – נציג מה באמת יש ב־XML
    print("\n========== DEBUG: סטטוסים שנמצאו ב־XML ==========")
    for s in sorted(all_statuses):
        print(f" - {s}")
    print("=================================================\n")

    print(f"🔍 נמצאו {len(filtered_records)} רשומות רלוונטיות מתוך {len(all_statuses)} סטטוסים שונים.\n")
    return filtered_records



def save_xml_to_csv(xml_content, account_id, text_widget=None):
    """
    שומר את תוכן ה-XML כ-CSV.
    - מחזיר גם את רשומות השגיאות / Ready to Settle.
    """
    csv_dir = CSV_DIR
    os.makedirs(csv_dir, exist_ok=True)

    records = parse_xml_for_errors(xml_content)  # שימוש בפונקציה החדשה
    df = pd.DataFrame(records)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    csv_file = os.path.join(csv_dir, f"{account_id}_{timestamp}.csv")
    df.to_csv(csv_file, index=False, encoding="utf-8-sig")

    if text_widget:
        log_write(text_widget, f"✅ CSV saved for merchant {account_id}: {csv_file}")

    if not df.empty and text_widget:
        log_write(text_widget, f"⚠️ Found {len(df)} relevant records (Technical Error / Ready to Settle) for merchant {account_id}")

    return df, csv_file
    
def download_xml(site, account, shift_id, token, cookies, text_widget=None):
    """
    מוריד XML עבור סניף (site) וחשבון (account) לפי shift_id,
    שומר את ה-XML הגולמי לקובץ ומדפיס DEBUG של כל הסטטוסים.
    """
    try:
        today = datetime.today()
        date_to = (today - timedelta(days=1)).strftime("%Y%m%d")   # עד אתמול
        date_from = (today - timedelta(days=2)).strftime("%Y%m%d") # יומיים אחורה

        log_write(text_widget, f"[XML] Generating report for {account}, site {site}, shift {shift_id} ({date_from}-{date_to})")
        xml_content = generate_report_xml(account, token, date_from, date_to, cookies)

        if not xml_content or len(xml_content) < 100:
            log_write(text_widget, f"⚠️ Empty or invalid XML for account {account}")
            return None

        # 💾 שמירת ה-XML הגולמי
        xml_dir = os.path.join(LOG_DIR, "xml")
        os.makedirs(xml_dir, exist_ok=True)
        xml_file = os.path.join(xml_dir, f"{account}_{date_from}_{date_to}.xml")
        with open(xml_file, "w", encoding="utf-8") as f:
            f.write(xml_content)
        log_write(text_widget, f"✅ XML saved to file: {xml_file}")

        # 🔍 DEBUG – הצגת סטטוסים שנמצאו ב-XML
        root = ET.fromstring(xml_content)
        all_statuses = set()
        for row in root.findall(".//row"):
            status = row.findtext("aggregated_status", default="").strip()
            all_statuses.add(status)
        if text_widget:
            log_write(text_widget, f"🔍 Found aggregated_status values: {sorted(all_statuses)}")
        else:
            print(f"🔍 Found aggregated_status values for {account}: {sorted(all_statuses)}")

        log_write(text_widget, f"✅ XML downloaded successfully for {account}")
        return xml_content

    except Exception as e:
        log_write(text_widget, f"❌ Failed to download XML for account {account}: {str(e)}")
        return None

# ================== איסוף נתונים ==================
def collect_data_all_merchants(text_widget, progressbar, percent_label):
    all_records = []
    attachments = []
    checked_merchants = []

    try:
        token, cookies = get_token(USER, PASSWORD, text_widget)
        if not token:
            log_write(text_widget, "❌ Cannot proceed, login failed for main user.")
            return all_records, attachments, checked_merchants

        total_accounts = len(ENTITY_IDS)
        for idx, account in enumerate(ENTITY_IDS, start=1):
            try:
                log_write(text_widget, f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Processing account {account}")

                # ✅ מוסיפים לכל בתי העסק לרשימת בדיקה
                if account not in checked_merchants:
                    checked_merchants.append(account)

                # נבדוק אם יש batch בעייתי
                batches = get_failed_batches(account, token, cookies, text_widget)

                if batches:
                    first_batch = batches[0]
                    site = first_batch["site"]
                    shift_id = first_batch["shiftId"]
                else:
                    # אם אין batch, נבדוק לפי אתר/סניף ראשי או id ברירת מחדל
                    site = get_default_site(account)  # צריך פונקציה שתחזיר סניף ברירת מחדל
                    shift_id = "000000"  # shift dummy אם אין batch

                log_write(text_widget, f"Downloading XML for account {account}, site {site}, shift {shift_id}")
                xml_content = download_xml(site, account, shift_id, token, cookies, text_widget)
                if not xml_content:
                    continue

                df, csv_file = save_xml_to_csv(xml_content, account, text_widget)

                # זיהוי סטטוס כולל לקובץ - טיפול במצב של DataFrame ריק
                if df.empty or "aggregated_status" not in df.columns:
                    log_write(text_widget, f"ℹ️ No relevant transactions found for {account}")
                    continue

                if "Ready to Settle" in df["aggregated_status"].astype(str).unique():
                    agg_status = "Ready to Settle"
                elif "Technical Error - Settlement" in df["aggregated_status"].astype(str).unique():
                    agg_status = "Technical Error - Settlement"
                else:
                    agg_status = df["aggregated_status"].astype(str).unique()[0]

                # חילוץ שם סניף אם קיים בקובץ
                branch_name_from_file = ""
                if not df.empty and "branch_name" in df.columns:
                    possible_names = df["branch_name"].dropna().unique()
                    if len(possible_names) > 0:
                         branch_name_from_file = possible_names[0]

                # רישום אחד בלבד לבית עסק
                all_records.append({
                    "account": account,
                    "site": site,
                    "shiftId": shift_id,
                    "csv_file": csv_file,
                    "branch_name_from_file": branch_name_from_file,
                    "totalValue": df["original_amount"].astype(float).sum() if "original_amount" in df.columns else 0,
                    "aggregated_status": agg_status
                })


                attachments.append(csv_file)
                log_write(text_widget, f"✅ CSV saved for merchant {account}: {csv_file}")

                progress = int((idx / total_accounts) * 100)
                progressbar['value'] = progress
                percent_label.config(text=f"{progress}%")
                progressbar.update_idletasks()

            except Exception as e_acc:
                log_write(text_widget, f"❌ Fatal error on account {account}: {str(e_acc)}")

    except Exception as e:
        log_write(text_widget, f"❌ Fatal error: {str(e)}")
        messagebox.showerror("Error", f"Error: {str(e)}")
    return all_records, attachments, checked_merchants


# ================== שליחת מייל ==================
def send_stfp_emails(records, BRANCH_DB, output_dir, to_recipients, cc_recipients=None, checked_merchants=None):
    import pythoncom, win32com.client as win32, zipfile, os
    from datetime import datetime
    import time

    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch('outlook.application')
        namespace = outlook.GetNamespace("MAPI")
        namespace.Logon()

        os.makedirs(output_dir, exist_ok=True)

        # פילוח רשומות
        errors = [r for r in records if "Technical Error - Settlement" in str(r.get("aggregated_status", ""))]
        ready = [r for r in records if "Ready to Settle" in str(r.get("aggregated_status", ""))]
        
        run_date = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        # בניית רשימת בתי העסק שנבדקו עם V
        checked_merchants_local = checked_merchants or []
        merchants_list_html = "".join([f"<div style='margin: 2px 0;'>✅ {m}</div>" for m in checked_merchants_local])

        def build_html_table(data, title, color="#27ae60"):
            if not data:
                return ""
                
            html = f"""
            <div style='margin-top: 20px;'>
                <h3 style='color: {color}; border-bottom: 2px solid {color}; padding-bottom: 5px;'>{title}</h3>
                <table style='border-collapse: collapse; width: 100%; font-size: 11pt;'>
                    <tr style='background-color:{color}; color:white; font-weight:bold;'>
                        <th style='padding: 8px; border: 1px solid #ddd;'>SITE</th>
                        <th style='padding: 8px; border: 1px solid #ddd;'>ACCOUNT</th>
                        <th style='padding: 8px; border: 1px solid #ddd;'>שם סניף</th>
                        <th style='padding: 8px; border: 1px solid #ddd;'>מזהה משמרת</th>
                        <th style='padding: 8px; border: 1px solid #ddd;'>סכום</th>
                    </tr>
            """

            total_sum = 0
            for idx, r in enumerate(data):
                site = r.get("site", "")
                account = r.get("account", "")
                # נסיון שליפה מה-DB
                branch_info = BRANCH_DB.get(site, {})
                db_branch_name = branch_info.get("branch_name", "Unknown Branch") if isinstance(branch_info, dict) else "Unknown Branch"
                
                # אם לא נמצא ב-DB, ננסה מהקובץ
                file_branch_name = r.get("branch_name_from_file", "")
                
                if db_branch_name != "Unknown Branch":
                    branch_name = db_branch_name
                elif file_branch_name and str(file_branch_name).lower() != "nan":
                    branch_name = file_branch_name
                else:
                    branch_name = "Unknown Branch"

                shift = r.get("shiftId", "")
                
                # תיקון סכום - הסרת החילוק ב-100 לפי בקשת המשתמש
                try:
                    amount = float(r.get("totalValue", 0))
                except:
                    amount = 0
                    
                total_sum += amount
                row_color = "#f2f2f2" if idx % 2 == 0 else "#ffffff"
                
                html += f"""
                    <tr style='background-color:{row_color};'>
                        <td style='padding: 8px; border: 1px solid #ddd;'>{site}</td>
                        <td style='padding: 8px; border: 1px solid #ddd;'>{account}</td>
                        <td style='padding: 8px; border: 1px solid #ddd;'>{branch_name}</td>
                        <td style='padding: 8px; border: 1px solid #ddd;'>{shift}</td>
                        <td style='padding: 8px; border: 1px solid #ddd; font-weight:bold; color:#c0392b; direction: ltr; text-align: right;'>₪{amount:,.2f}</td>
                    </tr>
                """

            html += f"""
                    <tr style='background-color:#f9e79f; font-weight:bold;'>
                        <td colspan='4' style='padding: 8px; border: 1px solid #ddd;'>סה"כ לקטגוריה ({len(data)} רשומות)</td>
                        <td style='padding: 8px; border: 1px solid #ddd;'>₪{total_sum:,.2f}</td>
                    </tr>
                </table>
            </div>
            """
            return html

        # קביעת נושא המייל והכותרת לפי המצב
        if errors:
            subject = f"❌ שגיאה בשידור - Technical Error - שופרסל - {run_date}"
            title_text = "שגיאות בשידור - Technical Error - Settlement"
            title_color = "#c0392b"
        elif ready:
            subject = f"⚠️ עסקאות מוכנות לשידור - שופרסל - {run_date}"
            title_text = "עסקאות מוכנות לשידור בשופרסל"
            title_color = "#f39c12"
        else:
            subject = f"✅ כל בתי העסק תקינים בשופרסל - {run_date}"
            title_text = "כל בתי העסק תקינים בשופרסל"
            title_color = "#27ae60"

        mail = outlook.CreateItem(0)
        mail.Subject = subject

        # הוספת נמענים
        to_list = to_recipients.replace(';', ' ').replace(',', ' ').split() if isinstance(to_recipients, str) else to_recipients
        for addr in to_list:
            if addr:
                recip = mail.Recipients.Add(addr.strip())
                recip.Type = 1 # olTo
        
        if cc_recipients:
            cc_list = cc_recipients.replace(';', ' ').replace(',', ' ').split() if isinstance(cc_recipients, str) else cc_recipients
            for addr in cc_list:
                if addr:
                    recip = mail.Recipients.Add(addr.strip())
                    recip.Type = 2 # olCC
        
        mail.Recipients.ResolveAll()

        # בניית גוף המייל
        html_body = f"""
        <html>
        <body style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; direction: rtl; text-align: right; background-color: #ffffff; color: #333; padding: 20px;">
            <div style="font-size: 10pt; color: #7f8c8d; margin-bottom: 15px;">תאריך הרצה: {run_date}</div>
            
            <div style="font-size: 18pt; font-weight: bold; color: {title_color}; margin-bottom: 10px;">
                {("✅ " if not errors and not ready else ("❌ " if errors else "⚠️ ")) + title_text}
            </div>

            <div style="background-color: #f8f9fa; border-right: 5px solid {title_color}; padding: 15px; margin-bottom: 20px;">
                <div style="font-weight: bold; margin-bottom: 10px;">סטטוס בדיקה:</div>
                <div>סה"כ נבדקו: {len(checked_merchants_local)} בתי עסק.</div>
                {f"<div>לא נמצאו שגיאות.</div>" if not errors else f"<div style='color:#c0392b;'>נמצאו {len(errors)} שגיאות טכניות!</div>"}
                <div style="margin-top: 10px; font-size: 10pt;">
                    {merchants_list_html}
                </div>
            </div>
        """

        if errors:
            html_body += build_html_table(errors, "שגיאות טכניות (Technical Error)", color="#c0392b")
            
        if ready:
            html_body += "<div style='margin-top: 15px; font-weight: bold; color: #e67e22;'>⚠️ יש עסקאות מוכנות לשידור.</div>"
            html_body += build_html_table(ready, "עסקאות מוכנות לשידור (Ready to Settle)", color="#e67e22")

        if not errors and not ready:
            html_body += "<p>אין עסקאות שגויות ואין עסקאות מוכנות לשידור.</p>"

        html_body += """
            <div style="margin-top: 30px; border-top: 1px solid #eee; padding-top: 20px;">
                <p>בברכה,<br><b style="color: #2c3e50;">מערכת הפקת דוחות אוטומטית</b></p>
            </div>
        </body>
        </html>
        """
        
        mail.HTMLBody = html_body

        # הוספת קבצים מצורפים
        if errors or ready:
            zip_path = os.path.join(output_dir, f"STFP_Reports_{datetime.now().strftime('%Y%m%d')}.zip")
            with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                # הוספת כל הקבצים שנוצרו
                unique_files = set()
                for r in records:
                    if "csv_file" in r and os.path.exists(r["csv_file"]):
                        unique_files.add(r["csv_file"])
                for f in unique_files:
                    zf.write(f, os.path.basename(f))
            if os.path.exists(zip_path):
                mail.Attachments.Add(zip_path)

        mail.Save()
        mail.Send()
        print(f"✅ נשלח מייל מאוחד: {subject}")

        # ניסיון "לדחוף" את אאוטלוק לשלוח
        try:
            for sync in namespace.SyncObjects:
                sync.Start()
        except:
            pass
            
        print("✅ תהליך שליחת המיילים הסתיים. ממתין 10 שניות לסיום...")
        time.sleep(10)

    except Exception as e:
        print(f"❌ שגיאה כללית בתהליך המייל: {e}")
    finally:
        pythoncom.CoUninitialize()



def run_gui():
    # --- חלון ראשי ---
    root = Tk()
    root.title("STFP Mail")
    root.geometry("700x600")
    root.configure(bg="#2c3e50")

    Label(root, text="Moshe Systems", font=("Arial", 30, "bold"), bg="#2c3e50", fg="white").pack(pady=15)

    # --- אזור טקסט עם Scrollbar ---
    frame = Frame(root)
    frame.pack(fill=BOTH, expand=True, padx=10, pady=5)
    scrollbar = Scrollbar(frame, orient=VERTICAL)
    scrollbar.pack(side=RIGHT, fill=Y)
    text_widget = Text(frame, yscrollcommand=scrollbar.set, wrap="word", height=20, width=80,
                       bg="#34495e", fg="white", font=("Consolas", 10))
    text_widget.pack(side=LEFT, fill=BOTH, expand=True)
    scrollbar.config(command=text_widget.yview)

    # --- Progressbar ---
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("green.Horizontal.TProgressbar", foreground='#27ae60', background='#27ae60')
    progressbar = ttk.Progressbar(root, style="green.Horizontal.TProgressbar", orient="horizontal",
                                  length=600, mode="determinate", maximum=100)
    progressbar.pack(pady=10)
    percent_label = Label(root, text="0%", font=("Arial", 14), bg="#2c3e50", fg="white")
    percent_label.pack()

    # --- Thread לעבודה ברקע ---
    def start_thread():
        def thread_func():
            try:
                # איסוף נתונים
                records, attachments, checked_merchants = collect_data_all_merchants(
                    text_widget, progressbar, percent_label
                )

                # שליחת מיילים
                send_stfp_emails(
                    records,
                    BRANCH_DB,
                    CSV_DIR,
                    to_recipients="fatema.msarwa@verifone.com; shufersal@verifone.com",
                    cc_recipients="moshe.isakov@verifone.com; nadav.lieber@verifone.com",
                    checked_merchants=checked_merchants
                )


                # הודעה למשתמש
                if records:
                    messagebox.showinfo(
                        "Finished",
                        f"✅ נמצאו {len(records)} סניפים עם שגיאות או עסקאות שלא שודרו.\n"
                        f"המיילים נשלחו בהצלחה.\n\nLog: {log_filename}"
                    )
                else:
                    messagebox.showinfo(
                        "All Clear",
                        f"✅ כל הסניפים תקינים.\nנשלח מייל תקין ללא שגיאות.\n\nLog: {log_filename}"
                    )
            except Exception as e:
                log_write(text_widget, f"❌ Fatal error in thread: {str(e)}")
                messagebox.showerror("Error", str(e))

        threading.Thread(target=thread_func, daemon=True).start()

    root.after(1000, start_thread)
    root.mainloop()


# ================== נקודת כניסה ==================
if __name__ == "__main__":
    run_gui()