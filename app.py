import streamlit as st
import pandas as pd
from docx import Document
import re
import pdfplumber

# --- הגדרות תצוגה ---
st.set_page_config(page_title="בודק תעודות פסגות", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)
st.title("🍎 בודק תעודות - דוח חריגות והתאמות")

def clean_txt(text):
    if not text: return ""
    t = str(text).strip()
    t = re.sub(r'[.!?,:;\"״”’\'\-]', '', t)
    t = " ".join(t.split())
    t = t.replace("ייך", "ך").replace("יך", "ך").replace("הינך", "הנך")
    return t

# --- 1. בנק המשפטים המלא ---
GRADE_BANK = {
    40: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע", "היעדרויותיך הרבות פגעו בתפקודך ובהישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "מעורבות בשיעורים הגשת כל המטלות והשקעת מאמצים לימודיים יקדמו אותך וישפרו את הבנתך בחומר הנלמד", "את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עלייך לגלות אחריות על למידתך להגיע בזמן לשיעורים ולבצע משימות באופן עקבי", "עלייך להתאמץ יותר ולתפקד בשיעורים", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה", "עלייך להקפיד להגיש את המטלות הנדרשות"],
    45: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך"],
    50: ["עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "עליך להיות נוכח בשיעורים נוכחות סדירה תקדם למידתך ותשפר הישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "היה עליך להקפיד על הגשת העבודות ולגלות מעורבות ואחריות ללמידה", "התנהגותך בשיעורים וחוסר הריכוז פגעו בהישגיך הלימודיים", "I believe you can do better"],
    55: ["לא השקעת מספיק בעבודת הבית והדבר יצר פערים בהבנה ובידיעת החומר", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים"],
    60: ["את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "עליך להקפיד על כתיבת תשובה מפורטת ומנומקת"],
    65: ["את ילדה שקטה ונעימת הליכות היי קשובה בשיעורים ואל תחששי לדבר בעברית", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים עלייך להימנע מהעדרויות", "הקפידי להשתתף באופן פעיל בשיח הקבוצתי על מנת לשפר את הבנתך ואת הישגייך ביכולתך לתרום מרעיונותייך לקבוצה"],
    70: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "עלייך להקפיד על תלבושת ספורט כנדרש", "את מתקשה בהבנת החומר"],
    75: ["ניכר שהשקעת זמן ומאמצים בהכנת שיעורי בית ועבודות והגעת להישגים נאים", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות", "הנך משתתפת באופן פעיל בשיעורי חנ'ג", "הנך מקפידה על נוכחות סדירה שותפה פעילה בשיעורים תורמת לשיח ומקדמת את הלמידה בקבוצה"],
    80: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות", "הנך נוכח בשיעורים מבצע את כל הנדרש בשיעור בקביעות וביסודיות"],
    85: ["את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו", "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות", "תרומתך לשיעורים מבורכת והישגייך טובים מאוד"],
    90: ["הנך נוכחת בשיעורים מבצעת את כל הנדרש בקביעות וביסודיות", "את תלמידה רצינית מגלה עניין והבנה ובעלת מוטיבציה להצלחה", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות", "הנך מגלה אחריות ורצינות בלמידה והדבר בא לידי ביטוי בנוכחות בהשתתפות פעילה בשיעורים"],
    95: ["את ראויה לשבח על הישגייך המצוינים", "את בעלת מוטיבציה פנימית וחשוב לך להצליח ולהתקדם בלמידה", "יכולתך להתבטא בכתב ראויה לשבח כתיבתך רהוטה מעניינת ועניינית"],
    100: ["את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו", "הנך תלמידה מצטיינת ויחסך למקצוע רציני", "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה"]
}

# --- 2. לוגיקה לעיבוד ---
clean_bank = {k: [clean_txt(n) for n in v] for k, v in GRADE_BANK.items()}

st.sidebar.header("העלאת קבצים")
pdf_file = st.sidebar.file_uploader("העלי בנק הערות PDF", type=['pdf'])
docx_file = st.file_uploader("העלי קובץ תעודה (Word)", type=['docx'])

if pdf_file:
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    clean_row = [str(c) for c in row if c]
                    nums = re.findall(r'\d+', " ".join(clean_row))
                    if nums and len(clean_row) > 1:
                        g = int(nums[0])
                        txt = max(clean_row, key=len)
                        clean_bank.setdefault(g, []).append(clean_txt(txt))

all_valid = [s for sublist in clean_bank.values() for s in sublist]

if docx_file:
    doc = Document(docx_file)
    results = []
    current_student = "לא זוהה"

    for element in doc.element.body.iter():
        # חילוץ שם
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
                        sentences = re.split(r'[.!?\n]+', note_text)
                        sentences = [s.strip() for s in sentences if len(s.strip()) > 5]
                        
                        not_in_bank = []
                        grade_matches = []
                        
                        for s in sentences:
                            c_s = clean_txt(s)
                            # בדיקת קיום בבנק
                            if not any(c_s in b or b in c_s for b in all_valid):
                                not_in_bank.append(s)
                            
                            # בדיקת התאמה לציון
                            matches = False
                            if grade_val in clean_bank:
                                matches = any(c_s in b or b in c_s for b in clean_bank[grade_val])
                            grade_matches.append("✅" if matches else "❌")

                        results.append({
                            "תלמיד/ה": current_student,
                            "מקצוע": cells[col_sub],
                            "ציון": grade_val,
                            "הערה מקורית": note_text,
                            "הערות שלא נמצאות בבנק": ", ".join(not_in_bank) if not_in_bank else "",
                            "תואם לציון?": " | ".join(grade_matches)
                        })

    if results:
        st.table(pd.DataFrame(results))
