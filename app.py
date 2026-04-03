import streamlit as st
import pandas as pd
from docx import Document
import re
import pdfplumber

# --- הגדרות תצוגה ---
st.set_page_config(page_title="בודק תעודות פסגות", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)
st.title("🍎 בודק תעודות - בדיקה כפולה (בנק + התאמה לציון)")

def clean_txt(text):
    if not text: return ""
    t = str(text).strip()
    t = re.sub(r'[.!?,:;\"״”’\'\-]', '', t)
    t = " ".join(t.split())
    t = t.replace("ייך", "ך").replace("יך", "ך").replace("הינך", "הנך")
    return t

# --- 1. בנק המשפטים המלא (כולל את כל מה ששלחת) ---
GRADE_BANK = {
    40: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע", "היעדרויותיך הרבות פגעו בתפקודך ובהישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "מעורבות בשיעורים הגשת כל המטלות והשקעת מאמצים לימודיים יקדמו אותך וישפרו את הבנתך בחומר הנלמד", "את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עלייך לגלות אחריות על למידתך", "עלייך להתאמץ יותר ולתפקד בשיעורים", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה", "עלייך להקפיד להגיש את המטלות הנדרשות"],
    50: ["עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "עליך להיות נוכח בשיעורים נוכחות סדירה תקדם למידתך ותשפר הישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "היה עליך להקפיד על הגשת העבודות ולגלות מעורבות ואחריות ללמידה", "התנהגותך בשיעורים וחוסר הריכוז פגעו בהישגיך הלימודיים", "I believe you can do better"],
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

# --- 2. עיבוד בנקים ---
clean_bank = {k: [clean_txt(n) for n in v] for k, v in GRADE_BANK.items()}

# העלאת קבצים
st.sidebar.header("העלאת נתונים")
pdf_file = st.sidebar.file_uploader("העלי בנק הערות (PDF)", type=['pdf'])
docx_file = st.file_uploader("העלי קובץ תעודה (Word)", type=['docx'])

# הוספת משפטים מה-PDF אם הועלה
if pdf_file:
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    row_txt = [str(c) for c in row if c]
                    nums = re.findall(r'\d+', " ".join(row_txt))
                    if nums and len(row_txt) > 1:
                        g = int(nums[0])
                        txt = max(row_txt, key=len)
                        clean_bank.setdefault(g, []).append(clean_txt(txt))

# רשימה של כל המשפטים המותרים (מכל הדרגות)
all_allowed_sentences = [s for sublist in clean_bank.values() for s in sublist]

if docx_file:
    doc = Document(docx_file)
    results = []
    current_student = "לא זוהה"

    for element in doc.element.body.iter():
        # חילוץ שם התלמיד
        if element.tag.endswith('p'):
            t = "".join(t.text for t in element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') if t.text)
            if any(k in t for k in ["שם התלמיד", "שם התלמידה"]):
                if ":" in t: current_student = t.split(":", 1)[1].strip()

        # עיבוד טבלה
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
                    
                    g_match = re.findall(r'\d+', cells[col_grade])
                    if not g_match: continue
                    grade_val = int(g_match[0])
                    
                    notes = [c for i, c in enumerate(cells) if i != col_grade and i != col_sub and len(c) > 5]
                    note_text = max(notes, key=len) if notes else ""
                    
                    if note_text:
                        # פירוק למשפטים
                        sentences = re.split(r'[.!?\n]+', note_text)
                        sentences = [s.strip() for s in sentences if len(s.strip()) > 5]
                        
                        summary_exists = []
                        summary_matches_grade = []
                        
                        for s in sentences:
                            c_s = clean_txt(s)
                            # 1. האם המשפט בכלל קיים בבנק?
                            exists = any(c_s in b or b in c_s for b in all_allowed_sentences)
                            summary_exists.append("✅" if exists else f"❌ ({s})")
                            
                            # 2. האם הוא מתאים לציון הספציפי?
                            matches_grade = False
                            if grade_val in clean_bank:
                                matches_grade = any(c_s in b or b in c_s for b in clean_bank[grade_val])
                            summary_matches_grade.append("✅" if matches_grade else "❌")

                        results.append({
                            "תלמיד/ה": current_student,
                            "מקצוע": cells[col_sub],
                            "ציון": grade_val,
                            "הערה מקורית": note_text,
                            "נמצא בבנק?": " | ".join(summary_exists),
                            "תואם לציון?": " | ".join(summary_matches_grade)
                        })

    if results:
        st.table(pd.DataFrame(results))
    else:
        st.warning("לא נמצאו נתונים לעיבוד.")
