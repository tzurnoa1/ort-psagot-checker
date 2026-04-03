import streamlit as st
import pandas as pd
from docx import Document
import re

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - אורט פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות - גרסה פשוטה ויציבה")

def normalize_simple(text):
    if not text: return ""
    t = str(text).strip()
    # האחדת גרשיים וסימני פיסוק
    t = re.sub(r'[.!?,:;\"״”’\'\-]', '', t)
    t = " ".join(t.split())
    # תיקוני מגדר נפוצים להשוואה בלבד
    t = t.replace("ייך", "ך").replace("יך", "ך").replace("הינך", "הנך")
    return t

# --- בנק המשפטים המלא (בזכר/נקבה מעורב - המערכת תסתדר) ---
GRADE_BANK = {
    40: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע", "היעדרויותיך הרבות פגעו בתפקודך ובהישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "מעורבות בשיעורים הגשת כל המטלות והשקעת מאמצים לימודיים יקדמו אותך וישפרו את הבנתך בחומר הנלמד", "את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עלייך לגלות אחריות על למידתך", "עלייך להתאמץ יותר ולתפקד בשיעורים", "עלייך להקפיד להגיש את המטלות הנדרשות"],
    50: ["עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "עליך להיות נוכח בשיעורים נוכחות סדירה תקדם למידתך ותשפר הישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "התנהגותך בשיעורים וחוסר הריכוז פגעו בהישגיך הלימודיים", "I believe you can do better"],
    60: ["את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים"],
    65: ["את ילדה שקטה ונעימת הליכות", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים עלייך להימנע מהעדרויות", "הקפידי להשתתף באופן פעיל בשיח הקבוצתי על מנת לשפר את הבנתך ואת הישגייך ביכולתך לתרום מרעיונותייך לקבוצה"],
    70: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "עלייך להקפיד על תלבושת ספורט כנדרש"],
    75: ["ניכר שהשקעת זמן ומאמצים בהכנת שיעורי בית ועבודות והגעת להישגים נאים", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות", "הנך משתתפת באופן פעיל בשיעורי חנ'ג", "הנך מקפידה על נוכחות סדירה שותפה פעילה בשיעורים"],
    80: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות", "הנך נוכח בשיעורים מבצע את כל הנדרש"],
    85: ["את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו", "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות"],
    90: ["הנך נוכחת בשיעורים מבצעת את כל הנדרש", "את תלמידה רצינית מגלה עניין והבנה", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות", "הנך מגלה אחריות ורצינות בלמידה"],
    95: ["את ראויה לשבח על הישגייך המצוינים", "יכולתך להתבטא בכתב ראויה לשבח"],
    100: ["את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו", "הנך תלמידה מצטיינת ויחסך למקצוע רציני", "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה"]
}

# הכנת הבנק להשוואה
normalized_bank = {k: [normalize_simple(n) for n in v] for k, v in GRADE_BANK.items()}

uploaded_file = st.file_uploader("העלי קובץ תעודות (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    results = []
    current_student = "לא זוהה"

    # סריקה לאיתור שם התלמיד והטבלאות
    for element in doc.element.body.iter():
        if element.tag.endswith('p'):
            text = "".join(t.text for t in element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') if t.text)
            if any(k in text for k in ["שם", "לכבוד", "תלמיד"]):
                clean_name = re.sub(r'[^א-ת\s]', '', text).replace("שם התלמיד", "").replace("התלמידה", "").replace("לכבוד", "").strip()
                if len(clean_name) > 2: current_student = clean_name

        elif element.tag.endswith('tbl'):
            from docx.table import Table
            table = Table(element, doc)
            if len(table.rows) < 2: continue
            
            headers = [c.text.strip() for c in table.rows[0].cells]
            col_grade = next((i for i, h in enumerate(headers) if "ציון" in h), -1)
            col_sub = next((i for i, h in enumerate(headers) if "מקצוע" in h), 0)
            
            if col_grade != -1:
                for row in table.rows[1:]:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) <= col_grade: continue
                    
                    grade_match = re.findall(r'\d+', cells[col_grade])
                    if not grade_match: continue
                    g_num = int(grade_match[0])
                    
                    # מציאת התא עם ההערה (הכי ארוך)
                    notes = [c for i, c in enumerate(cells) if i != col_grade and i != col_sub and len(c) > 5]
                    note_text = max(notes, key=len) if notes else ""
                    
                    if note_text:
                        norm_note = normalize_simple(note_text)
                        is_valid = False
                        
                        # בדיקה אם לפחות משפט אחד מהבנק נמצא בתוך הטקסט של המורה
                        if g_num in normalized_bank:
                            for bank_sentence in normalized_bank[g_num]:
                                if bank_sentence in norm_note or norm_note in bank_sentence:
                                    is_valid = True
                                    break
                        
                        results.append({
                            "תלמיד/ה": current_student,
                            "מקצוע": cells[col_sub],
                            "ציון": g_num,
                            "הערה": note_text,
                            "סטטוס": "✅ תקין" if is_valid else "❌ שגיאה"
                        })

    if results:
        st.table(pd.DataFrame(results))
    else:
        st.warning("לא נמצאו ציונים בקובץ.")
