import streamlit as st
import pandas as pd
from docx import Document
import re

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - אורט פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות - פתרון סופי לזיהוי שם")

# --- בנק המשפטים (40-45) ---
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

def find_student_name_final(doc):
    """סורק כל פיסת טקסט אפשרית במסמך"""
    full_text_list = []
    
    # 1. פסקאות רגילות
    for p in doc.paragraphs: full_text_list.append(p.text)
    
    # 2. טבלאות
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells: full_text_list.append(c.text)
            
    # 3. כותרות עליונות ותחתונות
    for section in doc.sections:
        for p in section.header.paragraphs: full_text_list.append(p.text)
        for p in section.footer.paragraphs: full_text_list.append(p.text)

    # 4. חיפוש בתוך "אלמנטים צפים" (תיבות טקסט) דרך ה-XML של המסמך
    import xml.etree.ElementTree as ET
    namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    root = ET.fromstring(doc._element.xml)
    for t in root.findall('.//w:t', namespace):
        full_text_list.append(t.text)

    # שלב הניתוח: חיפוש תבנית של שם
    for line in full_text_list:
        if not line: continue
        # מחפש "שם התלמיד" או "שם התלמידה" או רק "שם:"
        if any(keyword in line for keyword in ["שם התלמיד", "שם התלמידה", "שם:", "לכבוד"]):
            # ניקוי השורה: משאיר רק אותיות בעברית ורווחים
            clean = re.sub(r'(שם התלמיד/ה:|שם התלמיד:|שם התלמידה:|שם:|לכבוד|מס\' זהות:.*|ת\.ז:.*|כיתה:.*)', '', line)
            clean = re.sub(r'[^א-ת\s]', '', clean).strip()
            # אם נשאר טקסט שהוא לפחות שתי מילים (שם פרטי ומשפחה)
            if len(clean.split()) >= 2:
                return clean
                
    return "לא נמצא שם (בדקי אם השם הוא תמונה)"

uploaded_file = st.file_uploader("העלי קובץ תעודה (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    student_name = find_student_name_final(doc)
    
    all_results = []
    for table in doc.tables:
        if len(table.rows) < 2: continue
        headers = [c.text.strip() for c in table.rows[0].cells]
        col_grade = next((i for i, h in enumerate(headers) if "ציון" in h), -1)
        col_note = next((i for i, h in enumerate(headers) if "הערכה" in h or "מילולית" in h), -1)
        col_sub = next((i for i, h in enumerate(headers) if "מקצוע" in h), 0)
        
        if col_grade == -1 or col_note == -1: continue

        for row in table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) > max(col_grade, col_note):
                sub, grade, note = cells[col_sub], cells[col_grade], cells[col_note]
                if not (grade or note): continue
                
                # בדיקת לוגיקה
                nums = re.findall(r'\d+', str(grade))
                grade_num = int(nums[0]) if nums else 0
                error = None
                
                note_norm = normalize_hebrew(note)
                is_match = False
                if grade_num in GRADE_BANK:
                    for allowed in GRADE_BANK[grade_num]:
                        if normalize_hebrew(allowed) in note_norm:
                            is_match = True
                            break
                    if not is_match:
                        error = "📝 הערה חופשית"
                
                all_results.append({
                    "תלמיד": student_name,
                    "מקצוע": sub,
                    "ציון": grade,
                    "הערכה מילולית": note,
                    "סטטוס": "✅ תקין" if not error else "❌ בדיקה",
                    "פירוט": error if error else "תקין לפי הבנק"
                })

    if all_results:
        df = pd.DataFrame(all_results)
        st.subheader(f"דו\"ח בדיקה עבור: {student_name}")
        st.table(df.style.apply(lambda r: ['background-color: #ffcccc' if "❌" in str(r['סטטוס']) else '' for _ in r], axis=1))
