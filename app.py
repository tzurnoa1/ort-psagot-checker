import streamlit as st
import pandas as pd
from docx import Document
import re

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; } div.stButton > button { width: 100%; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות פסגות - גרסה יציבה")

# --- בנק הערות (שורשי מילים כדי לעקוף בעיות מגדר) ---
# הוספתי כאן "שורשים" - זה יזהה גם תלמיד וגם תלמידה
BANK_ROOTS = [
    "ראוי לשבח", "הישגיך המצוי", "הפגנת ידע", "אתגרים", "ניתוח טקסט", 
    "ידע עולם", "הנך תלמיד", "מצטיין", "רצינות", "אחריות", "התקדמות",
    "ביצוע משימות", "השתתפות", "ניסוח", "הבנת הנקרא", "מגלה עניין",
    "השקעה", "מוטיבציה", "עקביות", "שיפור", "למידה", "הישגים"
]

def is_note_in_bank(note):
    clean_n = str(note).replace('.', '').strip()
    if len(clean_n) < 4: return True # התעלמות משורות ריקות
    # בדיקה אם לפחות אחד מהשורשים נמצא בתוך ההערה
    return any(root in clean_n for root in BANK_ROOTS)

def clean_student_name(name_text):
    # ניקוי מילים מיותרות שנדבקות לשם
    for word in ["מס", "זהות", "ת.ז", "כתה", "כיתה", "תלמיד", "שם"]:
        name_text = name_text.replace(word, "")
    return re.sub(r'[:\-\d]', '', name_text).strip()

uploaded_file = st.file_uploader("העלי קובץ תעודה (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    all_data = []
    
    # חילוץ שם התלמיד בצורה חכמה יותר
    full_text = " ".join([p.text for p in doc.paragraphs])
    name_search = re.search(r"(שם|לכבוד)\s*[:/-]\s*([א-ת\s]+)", full_text)
    current_student = clean_student_name(name_search.group(2)) if name_search else "תלמיד/ה"

    for table in doc.tables:
        # זיהוי עמודות
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
                
                # ניתוח חריגות
                is_free = not is_note_in_bank(note)
                
                # בדיקת סתירה (ציון נמוך עם מילה חיובית מאוד)
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
        
        # הצגת הטבלה המלאה
        st.subheader(f"סיכום בדיקה עבור: {current_student}")
        
        # עיצוב טבלה צבעונית
        def highlight_errors(val):
            color = 'background-color: #ffcccc' if "⚠️" in str(val) or "❌" in str(val) else ''
            return color

        st.table(df.style.applymap(highlight_errors, subset=['הערה חופשית?', 'סתירה?']))
        
        # כפתור להורדת התוצאות
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("הורדי את תוצאות הבדיקה (CSV)", csv, "check_results.csv", "text/csv")
    else:
        st.info("לא נמצאה טבלת ציונים מתאימה בקובץ. ודאי שיש עמודות בשם 'ציון' ו-'הערכה מילולית'.")
