import streamlit as st
import pandas as pd
from docx import Document
import re

# --- עיצוב ---
st.set_page_config(page_title="בודק תעודות - פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות - גרסה קשוחה")

# מילים שנחשבות "חיוביות" ויוצרות סתירה עם ציון נמוך
POSITIVE_WORDS = ["מוטיבציה", "רצון", "שבח", "מצוין", "טוב", "נאה", "רהוט", "משתתף", "פעיל"]

def clean_student_name(text):
    # מנקה את כל מה שמסביב לשם
    text = text.replace("שם התלמיד/ה:", "").replace("שם התלמיד:", "").replace("שם:", "")
    text = text.split("מס'")[0] # חותך כשמגיע למספר זהות
    return text.strip()

def analyze_logic(grade_str, note):
    reasons = []
    try:
        # חילוץ מספר הציון
        grade_num = int("".join(filter(str.isdigit, str(grade_str))))
        
        # חוק סתירה: ציון נכשל עם הערה שנראית חיובית
        if grade_num < 55:
            for word in POSITIVE_WORDS:
                if word in note:
                    reasons.append(f"❓ סתירה: ציון {grade_num} (נכשל) עם הערת שבח/רצון")
                    break
    except:
        pass
    return reasons

uploaded_file = st.file_uploader("העלי קובץ תעודה", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    
    # חיפוש שם התלמיד בכל המסמך (פסקאות וטבלאות)
    student_name = "לא נמצא שם"
    all_content = []
    
    # סריקת כל הטקסט האפשרי
    for p in doc.paragraphs: all_content.append(p.text)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells: all_content.append(c.text)
            
    for line in all_content:
        if "שם התלמיד" in line:
            student_name = clean_student_name(line)
            break

    data = []
    for table in doc.tables:
        # בדיקה אם זו טבלת הציונים
        headers = [c.text.strip() for c in table.rows[0].cells]
        if not any("ציון" in h for h in headers): continue
        
        col_grade = next(i for i, h in enumerate(headers) if "ציון" in h)
        col_note = next(i for i, h in enumerate(headers) if "הערכה מילולית" in h)
        col_sub = next((i for i, h in enumerate(headers) if "מקצוע" in h), 0)

        for row in table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) > max(col_grade, col_note):
                sub = cells[col_sub]
                grade = cells[col_grade]
                note = cells[col_note]
                
                if not grade and not note: continue
                
                logic_errors = analyze_logic(grade, note)
                data.append({
                    "תלמיד": student_name,
                    "מקצוע": sub,
                    "ציון": grade,
                    "הערכה מילולית": note,
                    "סטטוס": "❌ חריגה" if logic_errors else "✅ תקין",
                    "פירוט": " | ".join(logic_errors) if logic_errors else ""
                })

    if data:
        df = pd.DataFrame(data)
        st.subheader(f"בדיקה עבור: {student_name}")
        
        # צביעת שורות עם חריגות
        def highlight(row):
            return ['background-color: #ffcccc' if row['סטטוס'] == "❌ חריגה" else '' for _ in row]
        
        st.table(df.style.apply(highlight, axis=1))
    else:
        st.info("לא נמצאה טבלת ציונים.")
