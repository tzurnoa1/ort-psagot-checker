import streamlit as st
import pandas as pd
from docx import Document
import re

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - אורט פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות - בנק משפטים מעודכן (40-45)")

# --- רשימת המשפטים המשותפת לציון 40 ו-45 ---
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

# --- בנק ההערות המלא ---
GRADE_BANK = {
    40: LOW_GRADE_SENTENCES,
    45: LOW_GRADE_SENTENCES,
    # כאן תוכלי להמשיך להוסיף לשאר הציונים
    95: [
        "את ראויה לשבח על הישגייך המצוינים",
        "יכולתך להתבטא בכתב ראויה לשבח"
    ]
}

def normalize_hebrew(text):
    """מנקה הטיות מגדריות בסיסיות להשוואה חכמה"""
    if not text: return ""
    text = str(text).strip()
    # החלפת סיומות נפוצות בנקבה לצורה בסיסית
    text = text.replace("ייך", "ך").replace("יך", "ך")
    text = text.replace("הינך", "הנך")
    text = text.replace("עליך", "עלך").replace("עלייך", "עלך")
    text = text.replace("שייך", "שך")
    return text

def check_grade_match(grade_str, note):
    try:
        nums = re.findall(r'\d+', str(grade_str))
        if not nums: return None
        grade_num = int(nums[0])
        rounded_grade = 5 * round(grade_num / 5)
        
        note_norm = normalize_hebrew(note)

        if rounded_grade in GRADE_BANK:
            # בדיקה אם המשפט קיים בבנק של הציון הנוכחי
            for allowed in GRADE_BANK[rounded_grade]:
                if normalize_hebrew(allowed) in note_norm or note_norm in normalize_hebrew(allowed):
                    return None # תקין
            
            # בדיקה אם המשפט שייך לציון גבוה משמעותית (סתירה)
            for other_grade, notes in GRADE_BANK.items():
                for n in notes:
                    if normalize_hebrew(n) in note_norm:
                        if other_grade > rounded_grade + 10:
                            return f"❌ סתירה: ציון {grade_num} עם הערה שמתאימה לציון {other_grade}"
            
            return "📝 הערה חופשית (לא מהבנק המוגדר לציון זה)"
    except:
        pass
    return None

def find_student_name(doc):
    # חיפוש "רדאר" מקיף לשם התלמיד
    search_areas = []
    for p in doc.paragraphs: search_areas.append(p.text)
    for s in doc.sections:
        for p in s.header.paragraphs: search_areas.append(p.text)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells: search_areas.append(c.text)
            
    for text in search_areas:
        if "שם התלמיד" in text or "לכבוד" in text:
            # מחפש שם בעברית אחרי מילת מפתח
            match = re.search(r"(שם התלמיד/ה?|לכבוד)\s*[:/-]?\s*([א-ת\s]+)", text)
            if match:
                name = match.group(2).split("מס'")[0].split("ת.ז")[0].strip()
                if len(name.split()) >= 2: # מוודא שיש לפחות שם פרטי ומשפחה
                    return name
    return "לא נמצא שם"

uploaded_file = st.file_uploader("העלי קובץ תעודה (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    student_name = find_student_name(doc)
    
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
                
                error = check_grade_match(grade, note)
                all_results.append({
                    "תלמיד": student_name,
                    "מקצוע": sub,
                    "ציון": grade,
                    "הערכה מילולית": note,
                    "סטטוס": "✅ תקין" if not error else "❌ חריגה",
                    "פירוט": error if error else "התאמה לבנק המאושר"
                })

    if all_results:
        df = pd.DataFrame(all_results)
        st.subheader(f"דו\"ח בדיקה עבור: {student_name}")
        st.table(df.style.apply(lambda r: ['background-color: #ffcccc' if "❌" in str(r['סטטוס']) else '' for _ in r], axis=1))
    else:
        st.error("לא נמצאו נתונים תקינים לסריקה.")
