import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import re

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - פסגות", page_icon="📑")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות חכם - אורט פסגות")

# --- בנק ההערות המובנה (מומלץ להעתיק לכאן את כל הרשימה מהקובץ) ---
POSITIVE_NOTES = [
    "אתה ראוי לשבח על הישגיך המצויינים", "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה",
    "גילית יכולת טובה בניתוח טקסטים ויישום הידע", "הנך בעל ידע עולם נרחב",
    "הישגיך בשפה טובים מאוד - יישר כוח!", "הנך תלמיד מצטיין ויחסך למקצוע רציני",
    "את ראויה לשבח על הישגיך המצוינים", "מגלה אחריות ורצינות בלמידה"
]

IMPROVEMENT_NOTES = [
    "עליך לגלות אחריות על למידתך", "עליך לגלות יותר מוטיבציה ואחריות ללמידה",
    "אתה עדיין מתקשה בהבנת החומר", "עליך לשפר את מיומנויות הבנת הנקרא",
    "עליך לגלות אחריות על למידתך... ולבצע משימות באופן עקבי"
]

FULL_BANK = [n.replace('\n', ' ').strip() for n in (POSITIVE_NOTES + IMPROVEMENT_NOTES)]

def clean_text(text):
    # ניקוי רווחים כפולים, ירידות שורה ותווים מיוחדים לבדיקה מדויקת
    return re.sub(r'\s+', ' ', str(text)).strip()

def analyze_row(grade_str, note):
    reasons = []
    cleaned_note = clean_text(note)
    
    grade_digits = "".join(filter(str.isdigit, str(grade_str)))
    if not grade_digits or len(cleaned_note) < 3:
        return None 

    grade_num = int(grade_digits)
    
    # בדיקת קיום בבנק (בדיקה גמישה)
    found = any((cleaned_note in b or b in cleaned_note) for b in FULL_BANK)
    if not found:
        reasons.append("הערה חופשית (לא בבנק)")
    
    # בדיקת סתירה
    if grade_num <= 55 and any((cleaned_note in b or b in cleaned_note) for b in POSITIVE_NOTES):
        reasons.append(f"סתירה: ציון נמוך ({grade_num}) עם הערה חיובית")
    if grade_num >= 90 and any((cleaned_note in b or b in cleaned_note) for b in IMPROVEMENT_NOTES):
        reasons.append(f"סתירה: ציון גבוה ({grade_num}) עם הערת שיפור")
        
    return reasons if reasons else None

uploaded_file = st.file_uploader("בחרי קובץ Word (תעודה)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    anomalies = []
    student_name = "לא נמצא שם"

    # שלב 1: חיפוש שם התלמיד בכל הטקסט של המסמך
    for para in doc.paragraphs:
        if "שם" in para.text and ":" in para.text:
            student_name = para.text.split(":")[-1].strip()
            break

    # שלב 2: סריקת טבלאות
    for table in doc.tables:
        header_cells = [clean_text(cell.text) for cell in table.rows[0].cells]
        
        col_grade = -1
        col_note = -1
        for i, header in enumerate(header_cells):
            if "ציון" in header: col_grade = i
            if "הערכה מילולית" in header: col_note = i
        
        if col_grade == -1 or col_note == -1:
            continue
            
        for row in table.rows[1:]:
            cells = [cell.text.strip() for cell in row.cells]
            if len(cells) > max(col_grade, col_note):
                subject = cells[0]
                grade = cells[col_grade]
                note = cells[col_note]
                
                res = analyze_row(grade, note)
                if res:
                    anomalies.append({
                        "תלמיד": student_name,
                        "מקצוע": subject,
                        "ציון": grade,
                        "חריגה": " | ".join(res)
                    })

    if anomalies:
        st.warning(f"נמצאו חריגות עבור: {student_name}")
        st.table(pd.DataFrame(anomalies))
    else:
        st.success(f"התעודה של {student_name} תקינה לחלוטין!")
