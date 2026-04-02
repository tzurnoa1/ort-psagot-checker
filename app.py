import streamlit as st
import pandas as pd
from docx import Document
import re
from io import BytesIO

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - פסגות", page_icon="📑")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות - אורט פסגות")
st.write("ניתן להעלות קובץ אקסל או קובץ וורד של תעודה.")

# --- בנק ההערות המובנה ---
POSITIVE_NOTES = [
    "אתה ראוי לשבח על הישגיך המצויינים", "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה",
    "גילית יכולת טובה בניתוח טקסטים ויישום הידע שנלמד בכיתה", "הנך בעל ידע עולם נרחב",
    "לקחת חלק פעיל בשיעורים והגשת את כל המטלות בזמן", "הנך תלמיד מצטיין ויחסך למקצוע רציני",
    "את ראויה לשבח על הישגיך המצוינים", "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה",
    "את מבינה היטב את הנקרא... ומסיקה מסקנות" # הוסיפי כאן עוד משפטים מהבנק
]

IMPROVEMENT_NOTES = [
    "עליך לגלות אחריות על למידתך", "עליך לגלות יותר מוטיבציה ואחריות ללמידה",
    "במהלך השיעורים, לא הפגנת רצינות", "אתה מתקשה בפיתוח נושא ובניסוחו בכתב",
    "עליך לגלות אחריות על למידתך... ולבצע משימות באופן עקבי", "את מתקשה בפיתוח נושא ובניסוחו בכתב"
]

FULL_BANK = POSITIVE_NOTES + IMPROVEMENT_NOTES

def analyze_text(name, subject, grade, note):
    reasons = []
    note = str(note).strip()
    
    # 1. בדיקת קיום בבנק (בדיקה גמישה למניעת בעיות רווחים)
    found_in_bank = any(note in b or b in note for b in FULL_BANK)
    if not found_in_bank:
        reasons.append("הערה חופשית (לא בבנק)")
    
    # 2. בדיקת סתירה (לפי הציון)
    try:
        grade_num = int(grade)
        if grade_num <= 55 and any(note in b or b in note for b in POSITIVE_NOTES):
            reasons.append(f"סתירה: ציון {grade_num} עם הערה חיובית")
        if grade_num >= 90 and any(note in b or b in note for b in IMPROVEMENT_NOTES):
            reasons.append(f"סתירה: ציון {grade_num} עם הערת שיפור")
    except:
        pass
        
    return reasons

# --- העלאת קבצים ---
uploaded_file = st.file_uploader("בחרי קובץ (Excel או Word)", type=['xlsx', 'docx'])

if uploaded_file:
    anomalies = []
    
    # טיפול בקובץ WORD
    if uploaded_file.name.endswith('.docx'):
        doc = Document(uploaded_file)
        # חיפוש נתונים בתוך טבלאות הוורד (שם בד"כ נמצאות התעודות)
        for table in doc.tables:
            for row in table.rows[1:]: # מדלג על הכותרת
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 3: # מוודא שיש לפחות מקצוע, ציון והערה
                    # כאן המערכת מנסה לנחש איזה תא הוא הציון (מחפשת מספר)
                    grade = ""
                    subject = cells[0]
                    note = cells[-1]
                    for cell in cells:
                        if cell.isdigit(): grade = cell
                    
                    res = analyze_text("תלמיד", subject, grade, note)
                    if res:
                        anomalies.append({"מקצוע": subject, "ציון": grade, "הערה": note, "חריגה": " | ".join(res)})

    # טיפול בקובץ EXCEL
    else:
        df = pd.read_excel(uploaded_file)
        for _, row in df.iterrows():
            res = analyze_text(row.get('שם',''), row.get('מקצוע',''), row.get('ציון',0), row.get('הערה',''))
            if res:
                anomalies.append({"מקצוע": row.get('מקצוע',''), "ציון": row.get('ציון',0), "הערה": row.get('הערה',''), "חריגה": " | ".join(res)})

    # --- פלט ---
    if anomalies:
        st.error(f"נמצאו {len(anomalies)} חריגות:")
        st.table(pd.DataFrame(anomalies))
    else:
        st.success("הבדיקה הסתיימה - לא נמצאו חריגות!")
