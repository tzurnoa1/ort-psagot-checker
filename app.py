import streamlit as st
import pandas as pd
from docx import Document
import re

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - אורט פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בדיקה חכמה לפי ציון מדויק (40-100)")

# --- בנק ההערות המלא לפי ציונים (עדכני כאן את המשפטים מהאקסל שלך) ---
# המפתח הוא הציון, והערך הוא רשימת המשפטים המותרים לאותו ציון
GRADE_BANK = {
    40: ["עליך לגלות יותר מוטיבציה ואחריות ללמידה.","עליך לגלות אחריות על למידתך, להגיע בזמן לשיעורים ולבצע משימות באופן עקבי.","חוסר ההשתתפות ביום השדה פגע בצייונך.","מעורבות בשיעורים, הגשת כל המטלות והשקעת מאמצים לימודיים יקדמו אותך וישפרו את הבנתך בחומר הנלמד.","ציונך נפגע עקב היעדרויותיך הרבות.","היעדרויותיך הרבות פגעו בתפקודך ובהישגיך.","לא שיתפת פעולה ולא גייסת כוחות ללמידה.","עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר.","למידה עקבית, השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע."],
    45: ["עליך לגלות יותר מוטיבציה ואחריות ללמידה.","עליך לגלות אחריות על למידתך, להגיע בזמן לשיעורים ולבצע משימות באופן עקבי.","חוסר ההשתתפות ביום השדה פגע בצייונך.","מעורבות בשיעורים, הגשת כל המטלות והשקעת מאמצים לימודיים יקדמו אותך וישפרו את הבנתך בחומר הנלמד.","ציונך נפגע עקב היעדרויותיך הרבות.","היעדרויותיך הרבות פגעו בתפקודך ובהישגיך.","לא שיתפת פעולה ולא גייסת כוחות ללמידה.","עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר.","למידה עקבית, השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע
    50: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים"],
    55: ["ניכר כי הינך משקיעה מאמצים בלימודייך"],
    60: ["הישגייך בתחום זה דורשים השקעה נוספת"],
    65: ["הנך בנתיב הנכון, המשך התמדה תביא לשיפור"],
    70: ["הישגייך טובים, מומלץ להעמיק בחומר"],
    75: ["שקדת על עבודתך ברצינות, מתוך אחריות ובגרות"],
    80: ["הנך משתתפת באופן פעיל ומגלה עניין"],
    85: ["הישגייך טובים מאוד, יישר כוח"],
    90: ["ראויה לשבח על הישגייך המצוינים", "כתיבתך רהוטה ועניינית"],
    95: ["הישגייך מעולים, את מפגינה ידע רב"],
    100: ["הצטיינות יתרה, כל הכבוד על ההשקעה והידע"]
}

def check_grade_note_match(grade_str, note):
    try:
        # חילוץ הציון המספרי
        nums = re.findall(r'\d+', str(grade_str))
        if not nums: return None
        grade_num = int(nums[0])
        
        # עיגול לקפיצות של 5 (ליתר ביטחון)
        rounded_grade = 5 * round(grade_num / 5)
        note_text = str(note).strip()

        # 1. בדיקה אם הציון קיים בבנק
        if rounded_grade in GRADE_BANK:
            allowed_for_this_grade = GRADE_BANK[rounded_grade]
            # בדיקה אם ההערה מופיעה ברשימה המותרת לציון הזה
            is_match = any(allowed in note_text or note_text in allowed for allowed in allowed_for_this_grade)
            
            if is_match:
                return None # הכל תקין
            
            # 2. אם לא נמצאה התאמה, נבדוק אם זו הערה של ציון אחר (סתירה)
            for other_grade, other_notes in GRADE_BANK.items():
                if any(n in note_text for n in other_notes):
                    if other_grade > rounded_grade + 10:
                        return f"❌ סתירה חמורה: ציון {grade_num} עם הערה ששייכת לציון {other_grade}"
                    if other_grade < rounded_grade - 10:
                        return f"⚠️ חוסר התאמה: ציון {grade_num} עם הערה שמתאימה לציון נמוך ({other_grade})"
            
            return "📝 הערה חופשית (לא מהבנק המוגדר)"
    except:
        pass
    return None

uploaded_file = st.file_uploader("העלי קובץ תעודה (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    all_data = []
    student_name = "לא נמצא שם"
    
    # חיפוש שם יסודי בכל פינה במסמך
    search_areas = []
    for p in doc.paragraphs: search_areas.append(p.text)
    for s in doc.sections:
        for p in s.header.paragraphs: search_areas.append(p.text)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells: search_areas.append(c.text)

    for text in search_areas:
        if "שם התלמיד" in text or "לכבוד" in text:
            match = re.search(r"(שם התלמיד/ה?|לכבוד)\s*[:/-]?\s*([א-ת\s]+)", text)
            if match:
                student_name = match.group(2).split("מס'")[0].split("ת.ז")[0].strip()
                break

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
                if not grade and not note: continue
                
                error_msg = check_grade_note_match(grade, note)
                all_data.append({
                    "תלמיד": student_name,
                    "מקצוע": sub,
                    "ציון": grade,
                    "הערכה מילולית": note,
                    "סטטוס": "✅ תקין" if not error_msg else "❌ חריגה",
                    "פירוט": error_msg if error_msg else "התאמה מושלמת לבנק"
                })

    if all_data:
        df = pd.DataFrame(all_data)
        st.subheader(f"דו\"ח בדיקה: {student_name}")
        st.table(df.style.apply(lambda r: ['background-color: #ffcccc' if "❌" in str(r['סטטוס']) else '' for _ in r], axis=1))
    else:
        st.error("לא נמצאו נתונים. ודאי שהטבלה מכילה עמודות 'ציון' ו-'הערכה מילולית'.")
