import streamlit as st
import pandas as pd
from docx import Document
import re

# --- עיצוב ---
st.set_page_config(page_title="בודק תעודות - אורט פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות - זיהוי לפי מיקום (שם מעל טבלה)")

# --- בנק משפטים 40-45 ---
LOW_GRADE_SENTENCES = [
    "עליך לגלות יותר מוטיבציה ואחריות ללמידה.",
    "עליך לגלות אחריות על למידתך, להגיע בזמן לשיעורים ולבצע משימות באופן עקבי.",
    "חוסר ההשתתפות ביום השדה פגע בצייונך.",
    "מעורבות בשיעורים, הגשת כל המטלות והשקעת מאמצים לימודיים יקדמו אותך וישפרו את הבנתך בחומר הנלמד.",
    "ציונך נפגע עקב היעדרויותיך הרבות.",
    "היעדרויותיך הרבות פגעו בתפקודך ובהישגיך.",
    "לא שיתפת פעולה ולא גייסת כוחות ללמידה.",
    "עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר.",
    "למידה עקבית, השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע."
]
GRADE_BANK = {40: LOW_GRADE_SENTENCES, 45: LOW_GRADE_SENTENCES}

def normalize_hebrew(text):
    if not text: return ""
    t = str(text).strip()
    t = t.replace("ייך", "ך").replace("יך", "ך").replace("הינך", "הנך")
    t = t.replace("עליך", "עלך").replace("עלייך", "עלך")
    return t

uploaded_file = st.file_uploader("העלי קובץ תעודות (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    all_results = []
    
    # עוברים על כל הטבלאות במסמך
    for i, table in enumerate(doc.tables):
        # זיהוי אם זו טבלת ציונים
        if len(table.rows) < 2: continue
        headers = [c.text.strip() for c in table.rows[0].cells]
        col_grade = next((idx for idx, h in enumerate(headers) if "ציון" in h), -1)
        col_note = next((idx for idx, h in enumerate(headers) if "הערכה" in h or "מילולית" in h), -1)
        col_sub = next((idx for idx, h in enumerate(headers) if "מקצוע" in h), 0)
        
        if col_grade == -1 or col_note == -1: continue

        # --- חיפוש השם מעל הטבלה ---
        # אנחנו מחפשים ב-5 הפסקאות שקדמו לטבלה הזו
        current_student = "לא זוהה שם"
        
        # מחפשים את המיקום של הטבלה בתוך גוף המסמך
        tbl_element = table._element
        parent = tbl_element.getparent()
        elements = list(parent)
        tbl_index = elements.index(tbl_element)
        
        # סורקים אחורה מהטבלה למעלה
        for j in range(tbl_index - 1, max(-1, tbl_index - 6), -1):
            prev_element = elements[j]
            # אם מצאנו פסקה (p)
            if prev_element.tag.endswith('p'):
                para_text = "".join(node.text for node in prev_element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'))
                if para_text.strip():
                    # ניקוי השם
                    name = re.sub(r'(שם התלמיד/ה:|שם התלמיד:|שם התלמידה:|שם:|לכבוד|מס\' זהות:.*|ת\.ז:.*|כיתה:.*)', '', para_text)
                    name = re.sub(r'[^א-ת\s]', '', name).strip()
                    if len(name.split()) >= 2:
                        current_student = name
                        break

        # --- עיבוד נתוני הטבלה ---
        for row in table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) > max(col_grade, col_note):
                sub, grade, note = cells[col_sub], cells[col_grade], cells[col_note]
                if not (grade or note): continue
                
                # בדיקת התאמה לבנק
                nums = re.findall(r'\d+', str(grade))
                grade_num = int(nums[0]) if nums else 0
                error = None
                
                if grade_num in GRADE_BANK:
                    note_norm = normalize_hebrew(note)
                    is_match = any(normalize_hebrew(allowed) in note_norm for allowed in GRADE_BANK[grade_num])
                    if not is_match:
                        error = "📝 הערה חופשית"
                
                all_results.append({
                    "תלמיד/ה": current_student,
                    "מקצוע": sub,
                    "ציון": grade,
                    "הערכה מילולית": note,
                    "סטטוס": "✅ תקין" if not error else "❌ בדיקה",
                    "פירוט": error if error else "תקין"
                })

    if all_results:
        df = pd.DataFrame(all_results)
        st.table(df)
        st.download_button("📥 הורדת הדו\"ח", df.to_csv(index=False, encoding='utf-16'), "report.csv", "text/csv")
    else:
        st.warning("לא נמצאו טבלאות ציונים. ודאי שיש עמודה בשם 'ציון'.")
