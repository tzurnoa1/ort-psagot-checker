import streamlit as st
import pandas as pd
from docx import Document
import re

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - אורט פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות - זיהוי שם לפי עוגן טקסט")

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
    
    # 1. חילוץ כל האלמנטים (פסקאות וטבלאות) לפי הסדר
    elements = []
    for block in doc.element.body:
        if block.tag.endswith('p'):
            text = "".join(node.text for node in block.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'))
            elements.append({'type': 'p', 'text': text})
        elif block.tag.endswith('tbl'):
            from docx.table import Table
            elements.append({'type': 'tbl', 'obj': Table(block, doc)})

    current_student = "לא זוהה שם"
    
    # 2. סריקה רציפה
    for i, el in enumerate(elements):
        if el['type'] == 'p':
            txt = el['text']
            # אם מצאנו את משפט העוגן שציינת
            if "להלן חוות דעתם" in txt:
                # השם נמצא בדרך כלל בפסקה אחת או שתיים מעל המשפט הזה
                for backtrack in range(i-1, max(-1, i-4), -1):
                    potential_name = elements[backtrack].get('text', "")
                    if potential_name.strip() and len(potential_name.split()) >= 2:
                        # ניקוי תארים וסימנים
                        clean = re.sub(r'(שם התלמיד/ה:|שם התלמיד:|שם:|לכבוד|תלמיד:|תלמידה:)', '', potential_name)
                        clean = re.sub(r'[^א-ת\s]', '', clean).strip()
                        if len(clean.split()) >= 2:
                            current_student = clean
                            break
        
        elif el['type'] == 'tbl':
            table = el['obj']
            if len(table.rows) < 2: continue
            
            headers = [c.text.strip() for c in table.rows[0].cells]
            col_grade = next((idx for idx, h in enumerate(headers) if "ציון" in h), -1)
            col_note = next((idx for idx, h in enumerate(headers) if "הערכה" in h or "מילולית" in h), -1)
            col_sub = next((idx for idx, h in enumerate(headers) if "מקצוע" in h), 0)
            
            if col_grade != -1 and col_note != -1:
                for row in table.rows[1:]:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) > max(col_grade, col_note):
                        sub, grade, note = cells[col_sub], cells[col_grade], cells[col_note]
                        if not (grade or note): continue
                        
                        # בדיקה מול הבנק
                        nums = re.findall(r'\d+', str(grade))
                        grade_num = int(nums[0]) if nums else 0
                        error = None
                        
                        if grade_num in GRADE_BANK:
                            note_norm = normalize_hebrew(note)
                            if not any(normalize_hebrew(a) in note_norm for a in GRADE_BANK[grade_num]):
                                error = "📝 הערה חופשית"
                        
                        all_results.append({
                            "שם התלמיד/ה": current_student,
                            "מקצוע": sub,
                            "ציון": grade,
                            "הערכה מילולית": note,
                            "סטטוס": "✅ תקין" if not error else "❌ בדיקה",
                            "פירוט": error if error else "תקין"
                        })

    if all_results:
        df = pd.DataFrame(all_results)
        st.table(df)
    else:
        st.error("לא נמצאו טבלאות ציונים תקינות.")
