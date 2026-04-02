import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

# עיצוב הממשק
st.set_page_config(page_title="בודק תעודות - פסגות", page_icon="📑")
st.markdown("""
    <style>
    .main { text-align: right; direction: rtl; }
    div.stButton > button:first-child { background-color: #007bff; color: white; }
    </style>
    """, unsafe_allow_html=True)

st.title("🍎 מערכת בדיקת תעודות - אורט פסגות")
st.write("העלי קובץ אקסל של ציונים לבדיקת חריגות אוטומטית מול בנק ההערות.")

# --- בנק ההערות המובנה ---
POSITIVE_NOTES = [
    "אתה ראוי לשבח על הישגיך המצויינים", "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה",
    "גילית יכולת טובה בניתוח טקסטים ויישום הידע", "הנך בעל ידע עולם נרחב",
    "הנך מקפיד על נוכחות סדירה, שותף פעיל בשיעורים", "הישגיך בשפה טובים מאוד - יישר כוח!",
    "אתה מבין היטב את הנקרא, מסוגל להבחין בעיקר", "הנך תלמיד מצטיין ויחסך למקצוע רציני",
    "גילית ידע וקישרת בין המושגים הנלמדים לאירועים", "הנך סקרן ובעל ידע עולם נרחב",
    "אתה שולט בניתוח קטעי מקור מילוליים וחזותיים", "אתה מבחין בין עובדות לדעות",
    "הנך שולט במיומנויוות העבודה עם ספר הלימוד", "לקחת חלק פעיל בשיעורים והגשת את כל המטלות",
    "עבודתך לסיכום הספר אותו בחרת לקרוא רצינית ומעמיקה", "עבודת המחקר שלך מקורית",
    "מגלה שליטה מעולה ביישומי המחשב", "הנך נוכח בשיעורים, מבצע את כל הנדרש",
    "אתה ראוי לשבח על התמדה בביצוע משימותיך", "הנך מגלה מוטיבציה ורצון להתקדם",
    "שקדת על עבודתך ועבדת ברצינות", "אתה תלמיד רציני, מגלה עניין והבנה",
    "הנך מגלה אחריות ורצינות בלמידה", "הנך שומר על כללי שיח ודיון",
    "הנך תלמיד נעים הליכות", "הנך מגלה יחס מכבד כלפי התלמידים"
]

IMPROVEMENT_NOTES = [
    "עליך להקפיד על כתיבת תשובה במבנה תקין", "עליך לשפר את מיומנות העבודה עם ספר הלימוד",
    "עליך להקפיד על כתיבת תשובה מפורטת ומנומקת", "עליך להקפיד להגיע לשיעורים ולהגיש את המטלות",
    "אינך עומד בדרישות התלקיט", "עליך לגלות אחריות על למידתך, להגיע בזמן",
    "עליך להיות נוכח בשיעורים", "למרות הידע הרב ויכולותיך, לא התייחסת ברצינות",
    "עליך לגלות יותר מוטיבציה ואחריות ללמידה", "אתה עדיין מתקשה בהבנת החומר",
    "במהלך השיעורים, לא הפגנת רצינות ולא הבעת נכונות ללמידה", "לא שיתפת פעולה ולא גייסת כוחות ללמידה",
    "התנהגותך בשיעורים וחוסר הריכוז פגעו בהישגיך", "לא השקעת מספיק בעבודות הבית",
    "עליך לשתף פעולה בעבודה בקבוצות", "עליך לשפר את מיומנויות הבנת הנקרא",
    "אתה מתקשה בפיתוח נושא ובניסוחו בכתב", "גלה איפוק במהלך השיעורים",
    "עליך להימנע מפטפוטים", "עליך להימנע מאיחורים", "ציונך נפגע עקב היעדרויותיך"
]

FULL_BANK = POSITIVE_NOTES + IMPROVEMENT_NOTES

# העלאת קובץ
uploaded_file = st.file_uploader("בחרי קובץ (Excel בלבד)", type=['xlsx'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    anomalies = []

    for index, row in df.iterrows():
        reasons = []
        note = str(row.get('הערה', '')).strip()
        grade = row.get('ציון', 0)
        
        # בדיקה 1: בנק
        if note not in FULL_BANK:
            reasons.append("הערה חופשית / לא בבנק")
        
        # בדיקה 2: סתירה חיובית
        if grade <= 55 and note in POSITIVE_NOTES:
            reasons.append("סתירה: ציון נכשל עם הערה חיובית")
            
        # בדיקה 3: סתירה שלילית
        if grade >= 90 and note in IMPROVEMENT_NOTES:
            reasons.append("סתירה: ציון מצוין עם הערת לשיפור")

        if reasons:
            anomalies.append({
                "תלמיד/ה": row.get('שם', 'לא ידוע'),
                "מקצוע": row.get('מקצוע', 'לא ידוע'),
                "ציון": grade,
                "סיבת החריגה": " | ".join(reasons)
            })

    if anomalies:
        st.warning(f"נמצאו {len(anomalies)} חריגות:")
        res_df = pd.DataFrame(anomalies)
        st.table(res_df)

        # יצירת WORD
        doc = Document()
        doc.add_heading('דוח חריגות בתעודות', 0)
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        headers = ['שם', 'מקצוע', 'ציון', 'תיאור']
        for i, h in enumerate(headers): table.rows[0].cells[i].text = h
        for a in anomalies:
            cells = table.add_row().cells
            cells[0].text, cells[1].text, cells[2].text, cells[3].text = str(a['תלמיד/ה']), str(a['מקצוע']), str(a['ציון']), a['סיבת החריגה']

        buffer = BytesIO()
        doc.save(buffer)
        st.download_button("📥 הורדי דוח Word סופי", buffer.getvalue(), "report.docx")
    else:
        st.success("מעולה! כל התעודות עומדות בתקן.")