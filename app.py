import streamlit as st
import pandas as pd
from docx import Document
import re
import pdfplumber
import difflib

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - אורט פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות - גרסה יציבה")

def normalize_hebrew(text):
    if not text: return ""
    t = str(text).strip()
    # האחדת גרשיים
    for char in ['"', '״', '”', '“', '’', '‘', '´']:
        t = t.replace(char, "'")
    # הסרת פיסוק ורווחים
    t = re.sub(r'[.!?,:;\-]', '', t)
    t = " ".join(t.split())
    # השוואת מגדר בסיסית
    t = t.replace("ייך", "ך").replace("יך", "ך").replace("הינך", "הנך")
    return t

# --- בנק המשפטים המלא ---
GRADE_BANK = {
    40: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע", "היעדרויותיך הרבות פגעו בתפקודך ובהישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "מעורבות בשיעורים הגשת כל המטלות והשקעת מאמצים לימודיים יקדמו אותך וישפרו את הבנתך בחומר הנלמד", "את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עלייך לגלות אחריות על למידתך להגיע בזמן לשיעורים ולבצע משימות באופן עקבי", "עלייך להתאמץ יותר ולתפקד בשיעורים", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה", "עלייך להקפיד להגיש את המטלות הנדרשות"],
    45: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע"],
    50: ["עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "עליך להיות נוכח בשיעורים נוכחות סדירה תקדם למידתך ותשפר הישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "היה עליך להקפיד על הגשת העבודות ולגלות מעורבות ואחריות ללמידה", "התנהגותך בשיעורים וחוסר הריכוז פגעו בהישגיך הלימודיים", "I believe you can do better"],
    55: ["לא השקעת מספיק בעבודת הבית והדבר יצר פערים בהבנה ובידיעת החומר", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים"],
    60: ["את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים"],
    65: ["את ילדה שקטה ונעימת הליכות היי קשובה בשיעורים ואל תחששי לדבר בעברית", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים עלייך להימנע מהעדרויות", "הקפידי להשתתף באופן פעיל בשיח הקבוצתי על מנת לשפר את הבנתך ואת הישגייך ביכולתך לתרום מרעיונותייך לקבוצה"],
    70: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "עלייך להקפיד על תלבושת ספורט כנדרש", "את מתקשה בהבנת החומר"],
    75: ["ניכר שהשקעת זמן ומאמצים בהכנת שיעורי בית ועבודות והגעת להישגים נאים", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות", "הנך משתתפת באופן פעיל בשיעורי חנ'ג", "הנך מקפידה על נוכחות סדירה שותפה פעילה בשיעורים תורמת לשיח ומקדמת את הלמידה בקבוצה"],
    80: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות", "הנך נוכח בשיעורים מבצע את כל הנדרש בשיעור בקביעות וביסודיות"],
    85: ["את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו", "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות"],
    90: ["הנך נוכחת בשיעורים מבצעת את כל הנדרש בשיעור בקביעות וביסודיות", "את תלמידה רצינית מגלה עניין והבנה ובעלת מוטיבציה להצלחה", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות"],
    95: ["את ראויה לשבח על הישגייך המצוינים", "יכולתך להתבטא בכתב ראויה לשבח כתיבתך רהוטה מעניינת ועניינית"],
    100: ["את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו", "הנך תלמידה מצטיינת ויחסך למקצוע רציני", "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה"]
}

# עיבוד בנק (כולל PDF אם יש)
active_bank = {k: [normalize_hebrew(n) for n in v] for k, v in GRADE_BANK.items()}
pdf_file = st.sidebar.file_uploader("העלי PDF של הבנק (אופציונלי)", type=['pdf'])
if pdf_file:
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table:
                    clean = [str(c).strip() for c in row if c]
                    digits = re.findall(r'\d+', " ".join(clean))
                    if digits and len(clean) > 1:
                        g = int(digits[0])
                        txt = normalize_hebrew(max(clean, key=len))
                        if g not in active_bank: active_bank[g] = []
                        active_bank[g].append(txt)

# עיבוד התעודה
uploaded_file = st.file_uploader("העלי קובץ תעודות (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    results = []
    current_student = "לא זוהה שם"

    for block in doc.element.body.iter():
        # זיהוי שם תלמיד מתוך פסקאות
        if block.tag.endswith('p'):
            para_text = "".join(t.text for t in block.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') if t.text)
            if any(k in para_text for k in ["שם התלמיד", "שם התלמידה", "שם:", "לכבוד"]):
                name = re.sub(r'(שם התלמיד|התלמידה|:|לכבוד)', '', para_text).strip()
                if name: current_student = name

        # עיבוד טבלאות ציונים
        elif block.tag.endswith('tbl'):
            from docx.table import Table
            table = Table(block, doc)
            if len(table.rows) < 2: continue
            
            headers = [c.text.strip() for c in table.rows[0].cells]
            col_grade = next((i for i, h in enumerate(headers) if "ציון" in h), -1)
            col_sub = next((i for i, h in enumerate(headers) if "מקצוע" in h), 0)
            
            if col_grade != -1:
                for row in table.rows[1:]:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) <= col_grade: continue
                    
                    # חילוץ ציון והערה
                    grade_match = re.findall(r'\d+', cells[col_grade])
                    if not grade_match: continue
                    g_num = int(grade_match[0])
                    
                    # בחירת ההערה (התא הכי ארוך שאינו הציון או המקצוע)
                    note_candidates = [c for i, c in enumerate(cells) if i != col_grade and i != col_sub and len(c) > 5]
                    note = max(note_candidates, key=len) if note_candidates else ""
                    
                    if not note: continue
                    
                    # בדיקת התאמה (מפרק למשפטים ובודק דמיון)
                    sentences = re.split(r'[.!?\n]+', note)
                    sentences = [s.strip() for s in sentences if len(s.strip()) > 5]
                    
                    invalid_parts = []
                    if g_num in active_bank:
                        for s in sentences:
                            norm_s = normalize_hebrew(s)
                            # בדיקת דמיון מעל 75%
                            is_found = any(difflib.SequenceMatcher(None, norm_s, opt).ratio() > 0.75 for opt in active_bank[g_num])
                            if not is_found:
                                invalid_parts.append(s)
                    
                    status = "✅ תקין" if not invalid_parts else "❌ שגיאה"
                    results.append({
                        "תלמיד/ה": current_student,
                        "מקצוע": cells[col_sub],
                        "ציון": g_num,
                        "הערכה": note,
                        "סטטוס": status,
                        "פירוט": "תקין" if not invalid_parts else f"לא מזוהה: {', '.join(invalid_parts)}"
                    })

    if results:
        st.table(pd.DataFrame(results))
    else:
        st.warning("לא נמצאו נתונים בקובץ. ודאי שיש טבלה עם עמודת 'ציון'.")
