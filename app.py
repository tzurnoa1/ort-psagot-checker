import streamlit as st
import pandas as pd
from docx import Document
import re

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - פסגות", page_icon="📑")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות חכם (תומך מגדר) - פסגות")

# --- בנק ההערות (כאן הכנסתי רק דוגמאות, כדאי להשלים את השאר) ---
BANK_NOTES = [
    "ראוי לשבח על הישגיך המצויינים", "הפגנת ידע והתמודדת עם אתגרים",
    "גילית יכולת טובה בניתוח טקסטים", "בעל ידע עולם נרחב",
    "הישגיך בשפה טובים מאוד", "תלמיד מצטיין ויחסך למקצוע רציני",
    "מגלה אחריות ורצינות בלמידה", "מבינה היטב את הנקרא",
    "עליך לגלות אחריות על למידתך", "מתקשה בהבנת החומר",
    "עליך לשפר את מיומנויות הבנת הנקרא", "היעדרויותיך הרבות פגעו בהישגיך"
]

# פונקציה לניקוי טקסט והתעלמות מהטיות זכר/נקבה
def normalize_hebrew(text):
    if not text: return ""
    text = str(text).replace('\n', ' ').replace('\r', ' ')
    # הסרת סיומות נפוצות של נקבה/רבים וסימני פיסוק כדי להשוות "שורש"
    text = re.sub(r'ות |ים |ה |ת |ך ', ' ', text) 
    text = re.sub(r'[^\w\s]', '', text)
    return " ".join(text.split()).strip()

CLEAN_BANK = [normalize_hebrew(n) for n in BANK_NOTES]

def analyze_row(grade_str, note):
    reasons = []
    clean_note = normalize_hebrew(note)
    
    grade_digits = "".join(filter(str.isdigit, str(grade_str)))
    if not grade_digits or len(clean_note) < 5:
        return None 

    grade_num = int(grade_digits)
    
    # בדיקה אם המשפט (או חלק משמעותי ממנו) קיים בבנק
    is_in_bank = any((clean_note in b or b in clean_note) for b in CLEAN_BANK)
            
    if not is_in_bank:
        reasons.append("הערה חופשית")
    
    # בדיקת סתירה לוגית (דוגמה בסיסית)
    if grade_num <= 55 and "שבח" in clean_note:
        reasons.append(f"סתירה: ציון נמוך ({grade_num}) עם הערת שבח")
    if grade_num >= 95 and "שפר" in clean_note:
        reasons.append(f"סתירה: ציון גבוה ({grade_num}) עם הערת שיפור")
        
    return reasons if reasons else None

uploaded_file = st.file_uploader("בחרי קובץ Word (תעודה)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    anomalies = []
    student_name = ""

    # חיפוש שם התלמיד
    all_text = ""
    for para in doc.paragraphs: all_text += para.text + " "
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells: all_text += cell.text + " "
    
    name_match = re.search(r"(שם|תלמיד|תלמידה|לכבוד)\s*[:/-]\s*([א-ת\s]+)", all_text)
    student_name = name_match.group(2).strip().split('\n')[0] if name_match else "לא נמצא שם"

    # סריקת טבלאות
    for table in doc.tables:
        header_row = [cell.text.strip() for cell in table.rows[0].cells]
        col_grade, col_note = -1, -1
        for i, header in enumerate(header_row):
            if "ציון" in header: col_grade = i
            if "מילולית" in header or "הערכה" in header: col_note = i
        
        if col_grade == -1 or col_note == -1: continue
            
        for row in table.rows[1:]:
            cells = [cell.text.strip() for cell in row.cells]
            if len(cells) > max(col_grade, col_note):
                subject, grade, note = cells[0], cells[col_grade], cells[col_note]
                res = analyze_row(grade, note)
                if res:
                    anomalies.append({"תלמיד": student_name, "מקצוע": subject, "ציון": grade, "חריגה": " | ".join(res)})

    if anomalies:
        st.warning(f"נמצאו חריגות עבור: {student_name}")
        st.table(pd.DataFrame(anomalies))
    else:
        st.success(f"התעודה של {student_name} תקינה!")
