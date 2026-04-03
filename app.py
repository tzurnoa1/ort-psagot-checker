import streamlit as st
import pandas as pd
from docx import Document
import re
import pdfplumber

# --- הגדרות תצוגה ---
st.set_config(page_title="בודק תעודות פסגות", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

def clean_txt(text):
    if not text: return ""
    t = str(text).strip()
    t = re.sub(r'[.!?,:;\"״”’\'\-]', '', t)
    t = " ".join(t.split())
    t = t.replace("ייך", "ך").replace("יך", "ך").replace("הינך", "הנך")
    return t

# --- 1. בנק המשפטים המובנה ---
GRADE_BANK = {
    40: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע", "היעדרויותיך הרבות פגעו בתפקודך ובהישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "מעורבות בשיעורים הגשת כל המטלות והשקעת מאמצים לימודיים יקדמו אותך וישפרו את הבנתך בחומר הנלמד", "את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עלייך לגלות אחריות על למידתך", "עלייך להתאמץ יותר ולתפקד בשיעורים"],
    50: ["עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "עליך להיות נוכח בשיעורים נוכחות סדירה תקדם למידתך ותשפר הישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "התנהגותך בשיעורים וחוסר הריכוז פגעו בהישגיך הלימודיים", "I believe you can do better"],
    60: ["את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים"],
    75: ["ניכר שהשקעת זמן ומאמצים בהכנת שיעורי בית ועבודות והגעת להישגים נאים", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות", "הנך משתתפת באופן פעיל בשיעורי חנ'ג"],
    90: ["הנך נוכחת בשיעורים מבצעת את כל הנדרש בקביעות וביסודיות", "את תלמידה רצינית מגלה עניין והבנה ובעלת מוטיבציה להצלחה", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות"],
    100: ["את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו", "הנך תלמידה מצטיינת ויחסך למקצוע רציני", "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה"]
}

# --- 2. העלאת קבצים ---
st.sidebar.header("קבצים")
pdf_file = st.sidebar.file_uploader("העלי בנק הערות (PDF)", type=['pdf'])
docx_file = st.file_uploader("העלי קובץ תעודה (Word)", type=['docx'])

# עיבוד בנק PDF (אם קיים)
pdf_bank = {}
if pdf_file:
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    clean_row = [str(c).strip() for c in row if c]
                    nums = re.findall(r'\d+', " ".join(clean_row))
                    if nums and len(clean_row) > 1:
                        grade = int(nums[0])
                        txt = max(clean_row, key=len)
                        if grade not in pdf_bank: pdf_bank[grade] = []
                        pdf_bank[grade].append(clean_txt(txt))

# איחוד בנקים
combined_bank = {k: [clean_txt(s) for s in v] for k, v in GRADE_BANK.items()}
for k, v in pdf_bank.items():
    combined_bank.setdefault(k, []).extend(v)

# כל המשפטים הקיימים בעולם (לבדיקה אם המשפט בכלל קיים בבנק)
all_sentences_in_bank = [s for sublist in combined_bank.values() for s in sublist]

if docx_file:
    doc = Document(docx_file)
    results = []
    current_student = "לא זוהה"

    for element in doc.element.body.iter():
        # חילוץ שם (אחרי נקודתיים)
        if element.tag.endswith('p'):
            text = "".join(t.text for t in element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') if t.text)
            if any(k in text for k in ["שם התלמיד", "שם התלמידה"]):
                if ":" in text: current_student = text.split(":", 1)[1].strip()

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
                    
                    # זיהוי ההערה
                    notes_cands = [c for i, c in enumerate(cells) if i != col_grade and i != col_sub and len(c) > 5]
                    original_note = max(notes_cands, key=len) if notes_cands else ""
                    
                    if original_note:
                        # פירוק למשפטים
                        sentences = re.split(r'[.!?\n]+', original_note)
                        sentences = [s.strip() for s in sentences if len(s.strip()) > 4]
                        
                        summary_exists = []
                        summary_match = []
                        
                        for s in sentences:
                            c_s = clean_txt(s)
                            # 1. האם המשפט קיים בבנק בכלל?
                            exists = any(c_s in b or b in c_s for b in all_sentences_in_bank)
                            summary_exists.append("✅" if exists else f"❌ ({s})")
                            
                            # 2. האם המשפט מתאים לציון הספציפי?
                            matches_grade = False
                            if grade_val in combined_bank:
                                matches_grade = any(c_s in b or b in c_s for b in combined_bank[grade_val])
                            summary_match.append("✅" if matches_grade else "❌")

                        results.append({
                            "תלמיד/ה": current_student,
                            "מקצוע": cells[col_sub],
                            "ציון": grade_val,
                            "הערה": original_note,
                            "קיים בבנק?": " | ".join(summary_exists),
                            "מתאים לציון?": " | ".join(summary_match)
                        })

    if results:
        st.table(pd.DataFrame(results))
    else:
        st.info("לא נמצאו נתוני ציונים לעיבוד.")
