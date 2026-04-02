import streamlit as st
import pandas as pd
from docx import Document
import re

# --- עיצוב ---
st.set_page_config(page_title="בודק תעודות - אורט פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; border-radius: 10px; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות - דוח ריבוי תלמידים")

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

def extract_name(text):
    """מחלץ שם מתוך שורה נתונה"""
    if any(k in text for k in ["שם התלמיד", "שם התלמידה", "שם:", "לכבוד"]):
        clean = re.sub(r'(שם התלמיד/ה:|שם התלמיד:|שם התלמידה:|שם:|לכבוד|מס\' זהות:.*|ת\.ז:.*|כיתה:.*)', '', text)
        clean = re.sub(r'[^א-ת\s]', '', clean).strip()
        if len(clean.split()) >= 2:
            return clean
    return None

uploaded_file = st.file_uploader("העלי קובץ תעודות (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    all_results = []
    current_student = "לא זוהה שם"

    # אנחנו עוברים על כל האלמנטים במסמך לפי הסדר שלהם
    for element in doc.element.body:
        # אם זה פסקה - נבדוק אם יש בה שם חדש
        if element.tag.endswith('p'):
            para_text = "".join(t.text for t in element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'))
            found_name = extract_name(para_text)
            if found_name:
                current_student = found_name

        # אם זו טבלה - נסרוק אותה ונשייך לסטודנט הנוכחי
        elif element.tag.endswith('tbl'):
            # מוצאים את הטבלה האמיתית בתוך האובייקט של doc
            # (דרך קצת טכנית אבל הכרחית כדי לשמור על סדר התלמידים)
            from docx.table import Table
            table = Table(element, doc)
            
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
                    
                    if grade_num in GRADE_BANK:
                        note_norm = normalize_hebrew(note)
                        is_match = any(normalize_hebrew(allowed) in note_norm for allowed in GRADE_BANK[grade_num])
                        if not is_match:
                            error = "📝 הערה חופשית"

                    all_results.append({
                        "שם התלמיד/ה": current_student,
                        "מקצוע": sub,
                        "ציון": grade,
                        "הערכה מילולית": note,
                        "סטטוס": "✅ תקין" if not error else "❌ בדיקה",
                        "פירוט": error if error else "תקין לפי הבנק"
                    })

    if all_results:
        df = pd.DataFrame(all_results)
        # תצוגה
        st.dataframe(df.style.apply(lambda r: ['background-color: #ffcccc' if "❌" in str(r['סטטוס']) else '' for _ in r], axis=1), use_container_width=True)
        
        # אפשרות להורדה
        csv = df.to_csv(index=False, encoding='utf-16')
        st.download_button("📥 הורדת הדו\"ח המלא (Excel/CSV)", csv, "report.csv", "text/csv")
    else:
        st.error("לא נמצאו טבלאות ציונים במסמך.")
