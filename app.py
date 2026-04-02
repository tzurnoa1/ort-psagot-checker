import streamlit as st
import pandas as pd
from docx import Document
import re

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; } </style>""", unsafe_allow_html=True)

st.title("🍎 מערכת בדיקה - אורט פסגות")

# --- בנק שורשים (מעודכן לפי התמונה שלך) ---
BANK_ROOTS = [
    "ראויה לשבח", "הישגיך", "הפגנת ידע", "ניתוח טקסט", "תלמיד", "מצטיין", 
    "רצינות", "אחריות", "התקדמות", "ביצוע משימות", "הבנת הנקרא", 
    "מגלה מוטיבציה", "ורצון להתקדם", "שקדת על עבודתך", "השתתפות פעילה"
]

def clean_name(text):
    """מחלץ את השם בלבד מתוך השורה העלונה"""
    # מסיר את כל מה שמהמילה 'מס' זהות והלאה
    text = text.split("מס'")[0]
    # מסיר את כותרת השדה
    text = re.sub(r'שם התלמיד/ה:', '', text)
    return text.strip()

def analyze_row(grade_str, note):
    reasons = []
    # בדיקת בנק (חיפוש שורש אחד לפחות)
    is_in_bank = any(root in note for root in BANK_ROOTS)
    if not is_in_bank:
        reasons.append("הערה חופשית")
    
    # בדיקת סתירה (לפי התמונה: ציון 45 עם הערת מוטיבציה/אחריות)
    try:
        grade_num = int("".join(filter(str.isdigit, str(grade_str))))
        if grade_num < 50 and "ראויה לשבח" in note:
            reasons.append(f"סתירה: ציון {grade_num} עם 'ראויה לשבח'")
    except:
        pass
    return reasons

uploaded_file = st.file_uploader("העלי את קובץ התעודה", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    all_text = "\n".join([p.text for p in doc.paragraphs])
    
    # 1. חילוץ שם התלמיד לפי התבנית שבתמונה
    student_name = "לא נמצא שם"
    if "שם התלמיד/ה:" in all_text:
        line = [l for l in all_text.split('\n') if "שם התלמיד/ה:" in l][0]
        student_name = clean_name(line)

    data = []
    for table in doc.tables:
        # זיהוי עמודות לפי התמונה
        headers = [c.text.strip() for c in table.rows[0].cells]
        col_grade, col_note, col_sub = -1, -1, -1
        
        for i, h in enumerate(headers):
            if "ציון" in h: col_grade = i
            if "הערכה מילולית" in h: col_note = i
            if "מקצוע" in h: col_sub = i
            
        if col_grade == -1 or col_note == -1: continue

        for row in table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) > max(col_grade, col_note):
                sub = cells[col_sub] if col_sub != -1 else "כללי"
                grade = cells[col_grade]
                note = cells[col_note]
                
                if not grade and not note: continue
                
                errors = analyze_row(grade, note)
                data.append({
                    "תלמיד": student_name,
                    "מקצוע": sub,
                    "ציון": grade,
                    "הערכה מילולית": note[:60] + "...",
                    "חריגות": " | ".join(errors) if errors else "✅ תקין"
                })

    if data:
        df = pd.DataFrame(data)
        st.subheader(f"דו\"ח בדיקה: {student_name}")
        
        def color_errors(val):
            return 'background-color: #ffcccc' if "✅" not in str(val) else ''
        
        st.table(df.style.map(color_errors, subset=['חריגות']))
    else:
        st.error("לא הצלחתי לקרוא את הטבלה. ודאי שהכותרות הן 'ציון' ו-'הערכה מילולית'.")
