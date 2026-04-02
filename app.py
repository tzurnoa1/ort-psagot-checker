import streamlit as st
import pandas as pd
from docx import Document
import re

# --- עיצוב ---
st.set_page_config(page_title="בודק תעודות - פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות - גרסה יציבה")

# מילים חיוביות לבדיקת סתירה
POSITIVE_WORDS = ["מוטיבציה", "רצון", "שבח", "מצוין", "טוב", "נאה", "רהוט", "משתתף", "פעיל", "התקדמות"]

def clean_student_name(text):
    # ניקוי השם מסימנים ומספרי זהות
    text = re.sub(r'(שם התלמיד/ה:|שם התלמיד:|שם:|מס\' זהות:.*|ת\.ז:.*)', '', text)
    text = re.sub(r'[\d\-\.]', '', text)
    return text.strip()

def analyze_logic(grade_str, note):
    reasons = []
    try:
        # חילוץ מספר הציון
        nums = re.findall(r'\d+', str(grade_str))
        if nums:
            grade_num = int(nums[0])
            # חוק סתירה: ציון נמוך מ-55 עם מילה חיובית
            if grade_num < 55:
                for word in POSITIVE_WORDS:
                    if word in str(note):
                        reasons.append(f"❓ סתירה: ציון {grade_num} עם הערת '{word}'")
                        break
    except:
        pass
    return reasons

uploaded_file = st.file_uploader("העלי קובץ תעודה", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    all_data = []
    student_name = "לא נמצא שם"
    
    # 1. חיפוש השם בכל המסמך
    for p in doc.paragraphs:
        if "שם התלמיד" in p.text or "לכבוד" in p.text:
            student_name = clean_student_name(p.text)
            break
    if student_name == "לא נמצא שם":
        # ניסיון חיפוש בטבלה הראשונה (לפעמים השם שם)
        for table in doc.tables[:1]:
            for row in table.rows:
                for cell in row.cells:
                    if "שם" in cell.text:
                        student_name = clean_student_name(cell.text)
                        break

    # 2. סריקת טבלאות ציונים
    for table in doc.tables:
        if len(table.rows) < 2: continue
        
        # זיהוי עמודות בצורה גמישה
        header_cells = [c.text.strip() for c in table.rows[0].cells]
        
        col_grade, col_note, col_sub = -1, -1, 0
        
        for i, h in enumerate(header_cells):
            if "ציון" in h: col_grade = i
            if "הערכה" in h or "מילולית" in h: col_note = i
            if "מקצוע" in h: col_sub = i
            
        # אם לא מצאנו עמודות קריטיות, נדלג על הטבלה הזו בלי לקרוס
        if col_grade == -1 or col_note == -1:
            continue

        for row in table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) > max(col_grade, col_note):
                sub = cells[col_sub]
                grade = cells[col_grade]
                note = cells[col_note]
                
                if not grade and not note: continue
                
                logic_errors = analyze_logic(grade, note)
                all_data.append({
                    "תלמיד": student_name,
                    "מקצוע": sub,
                    "ציון": grade,
                    "הערכה מילולית": note,
                    "סטטוס": "❌ חריגה" if logic_errors else "✅ תקין",
                    "פירוט": " | ".join(logic_errors) if logic_errors else ""
                })

    if all_data:
        df = pd.DataFrame(all_data)
        st.subheader(f"בדיקה עבור: {student_name}")
        
        def highlight_errors(row):
            return ['background-color: #ffcccc' if row['סטטוס'] == "❌ חריגה" else '' for _ in row]
        
        st.table(df.style.apply(highlight_errors, axis=1))
    else:
        st.warning("לא נמצאה טבלת ציונים מתאימה. ודאי שיש עמודה שכתוב בה 'ציון' ועמודה שכתוב בה 'הערכה מילולית'.")
