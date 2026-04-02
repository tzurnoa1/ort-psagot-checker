import streamlit as st
import pandas as pd
from docx import Document
import re

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; } div.stButton > button { width: 100%; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות פסגות - גרסה מתוקנת")

# --- בנק שורשים (לעקיפת בעיות מגדר) ---
BANK_ROOTS = [
    "ראוי לשבח", "הישגיך המצוי", "הפגנת ידע", "אתגרים", "ניתוח טקסט", 
    "ידע עולם", "הנך תלמיד", "מצטיין", "רצינות", "אחריות", "התקדמות",
    "ביצוע משימות", "השתתפות", "ניסוח", "הבנת הנקרא", "מגלה עניין",
    "השקעה", "מוטיבציה", "עקביות", "שיפור", "למידה", "הישגים"
]

def is_note_in_bank(note):
    clean_n = str(note).replace('.', '').strip()
    if len(clean_n) < 4: return True
    return any(root in clean_n for root in BANK_ROOTS)

def clean_student_name(name_text):
    # ניקוי אגרסיבי של תווים ומילים מיותרות
    name_text = re.sub(r'(מס|זהות|ת\.ז|כתה|כיתה|תלמיד|שם|לכבוד)', '', name_text)
    name_text = re.sub(r'[:\-\d\.]', '', name_text)
    return name_text.strip()

uploaded_file = st.file_uploader("העלי קובץ תעודה (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    all_data = []
    
    # חילוץ שם התלמיד
    full_text = " ".join([p.text for p in doc.paragraphs])
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells: full_text += " " + cell.text
            
    name_search = re.search(r"(שם|לכבוד)\s*[:/-]\s*([א-ת\s]+)", full_text)
    current_student = clean_student_name(name_search.group(2)) if name_search else "תלמיד/ה"

    for table in doc.tables:
        header_cells = [cell.text.strip() for cell in table.rows[0].cells]
        col_grade, col_note = -1, -1
        
        for i, h in enumerate(header_cells):
            if "ציון" in h: col_grade = i
            if "מילולית" in h or "הערכה" in h: col_note = i
        
        if col_grade == -1 or col_note == -1: continue
            
        for row in table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) > max(col_grade, col_note):
                subject = cells[0]
                grade = cells[col_grade]
                note = cells[col_note]
                
                is_free = not is_note_in_bank(note)
                is_contradict = False
                grade_num = "".join(filter(str.isdigit, grade))
                if grade_num and int(grade_num) <= 55:
                    if any(w in note for w in ["מצוין", "שבח", "טוב מאד", "הצלחה"]):
                        is_contradict = True
                
                if subject and (grade or note):
                    all_data.append({
                        "תלמיד": current_student,
                        "מקצוע": subject,
                        "ציון": grade,
                        "הערכה מילולית": note,
                        "הערה חופשית?": "⚠️ כן" if is_free else "✅ בנק",
                        "סתירה?": "❌ סתירה!" if is_contradict else "✅ תקין"
                    })

    if all_data:
        df = pd.DataFrame(all_data)
        st.subheader(f"סיכום בדיקה עבור: {current_student}")
        
        # התיקון לשגיאה: שימוש ב-map במקום applymap
        def style_rows(v):
            if "⚠️" in str(v) or "❌" in str(v):
                return 'background-color: #ffcccc'
            return ''

        st.table(df.style.map(style_rows, subset=['הערה חופשית?', 'סתירה?']))
        
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 הורדי את תוצאות הבדיקה", csv, "check_results.csv", "text/csv")
    else:
        st.info("לא נמצאה טבלת ציונים מתאימה.")
