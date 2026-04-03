import streamlit as st
import pandas as pd
from docx import Document
import re

# --- הגדרות תצוגה ---
st.set_page_config(page_title="בודק תעודות - אורט פסגות", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)
st.title("🍎 מערכת בדיקת תעודות - גרסה מלאה")

def clean_txt(text):
    if not text: return ""
    t = str(text).strip()
    # הסרת פיסוק וגרשיים לצורך השוואה חלקה
    t = re.sub(r'[.!?,:;\"״”’\'\-]', '', t)
    t = " ".join(t.split())
    # נרמול מגדר בסיסי להשוואה
    t = t.replace("ייך", "ך").replace("יך", "ך").replace("הינך", "הנך")
    return t

# --- 1. בנק המשפטים המלא (כל ההתאמות ששלחת) ---
GRADE_BANK = {
    40: [
        "הנך מגלה מוטיבציה ורצון להתקדם בלימודים",
        "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע",
        "היעדרויותיך הרבות פגעו בתפקודך ובהישגיך",
        "ציונך נפגע עקב היעדרויותיך הרבות",
        "מעורבות בשיעורים הגשת כל המטלות והשקעת מאמצים לימודיים יקדמו אותך וישפרו את הבנתך בחומר הנלמד",
        "את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה",
        "עלייך לגלות אחריות על למידתך להגיע בזמן לשיעורים ולבצע משימות באופן עקבי ביכולתך לתרום מידיעותייך ומרעיונותייך לקבוצה ולהעשיר את השיח הקבוצתי",
        "עלייך להתאמץ יותר ולתפקד בשיעורים",
        "עלייך לגלות יותר מוטיבציה ואחריות ללמידה",
        "עלייך להקפיד להגיש את המטלות הנדרשות"
    ],
    45: [
        "הנך מגלה מוטיבציה ורצון להתקדם בלימודים",
        "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע"
    ],
    50: [
        "עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר",
        "עליך להיות נוכח בשיעורים נוכחות סדירה תקדם למידתך ותשפר הישגיך",
        "ציונך נפגע עקב היעדרויותיך הרבות",
        "היה עליך להקפיד על הגשת העבודות ולגלות מעורבות ואחריות ללמידה",
        "התנהגותך בשיעורים וחוסר הריכוז פגעו בהישגיך הלימודיים ובהבנת החומר הנלמד בכיתה",
        "I believe you can do better"
    ],
    55: [
        "לא השקעת מספיק בעבודת הבית והדבר יצר פערים בהבנה ובידיעת החומר",
        "עלייך לגלות יותר מוטיבציה ואחריות ללמידה",
        "הנך מגלה מוטיבציה ורצון להתקדם בלימודים"
    ],
    60: [
        "את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה",
        "עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר",
        "הנך מגלה מוטיבציה ורצון להתקדם בלימודים",
        "עליך להקפיד על כתיבת תשובה מפורטת ומנומקת"
    ],
    65: [
        "את ילדה שקטה ונעימת הליכות היי קשובה בשיעורים ואל תחששי לדבר בעברית תרגול קריאה וכתיבה יקדמו אותך ברכישת השפה העברית",
        "עלייך לגלות יותר מוטיבציה ואחריות ללמידה עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר",
        "הנך מגלה מוטיבציה ורצון להתקדם בלימודים עלייך להימנע מהעדרויות",
        "הקפידי להשתתף באופן פעיל בשיח הקבוצתי על מנת לשפר את הבנתך ואת הישגייך ביכולתך לתרום מרעיונותייך לקבוצה"
    ],
    70: [
        "הנך מגלה מוטיבציה ורצון להתקדם בלימודים",
        "עלייך להקפיד על תלבושת ספורט כנדרש",
        "את מתקשה בהבנת החומר"
    ],
    75: [
        "ניכר שהשקעת זמן ומאמצים בהכנת שיעורי בית ועבודות והגעת להישגים נאים",
        "שקדת על עבודתך ברצינות מתוך אחריות ובגרות",
        "הנך משתתפת באופן פעיל בשיעורי חנ'ג",
        "הנך מקפידה על נוכחות סדירה שותפה פעילה בשיעורים תורמת לשיח ומקדמת את הלמידה בקבוצה"
    ],
    80: [
        "הנך מגלה מוטיבציה ורצון להתקדם בלימודים",
        "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות",
        "הנך נוכח בשיעורים מבצע את כל הנדרש בשיעור בקביעות וביסודיות"
    ],
    85: [
        "את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו",
        "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות",
        "תרומתך לשיעורים מבורכת והישגייך טובים מאוד"
    ],
    90: [
        "הנך נוכחת בשיעורים מבצעת את כל הנדרש בקביעות וביסודיות",
        "את תלמידה רצינית מגלה עניין והבנה ובעלת מוטיבציה להצלחה",
        "שקדת על עבודתך ברצינות מתוך אחריות ובגרות",
        "הנך מגלה אחריות ורצינות בלמידה והדבר בא לידי ביטוי בנוכחות בהבנת החומר הנלמד בהשתתפות פעילה בשיעורים ובביצוע כל המטלות"
    ],
    95: [
        "את ראויה לשבח על הישגייך המצוינים",
        "את בעלת מוטיבציה פנימית וחשוב לך להצליח ולהתקדם בלמידה",
        "יכולתך להתבטא בכתב ראויה לשבח כתיבתך רהוטה מעניינת ועניינית"
    ],
    100: [
        "את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו",
        "הנך תלמידה מצטיינת ויחסך למקצוע רציני",
        "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה"
    ]
}

# --- 2. לוגיקה לעיבוד הקובץ ---
clean_bank = {k: [clean_txt(n) for n in v] for k, v in GRADE_BANK.items()}

uploaded_file = st.file_uploader("העלי קובץ תעודות (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    results = []
    current_student = "לא זוהה שם"

    def get_name_after_colon(text):
        if "שם התלמיד" in text or "שם התלמידה" in text:
            if ":" in text:
                return text.split(":", 1)[1].strip()
        return None

    # סריקה
    for element in doc.element.body.iter():
        # חילוץ שם
        if element.tag.endswith('p'):
            p_text = "".join(t.text for t in element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') if t.text)
            potential_name = get_name_after_colon(p_text)
            if potential_name:
                current_student = potential_name

        # עיבוד טבלאות
        elif element.tag.endswith('tbl'):
            from docx.table import Table
            table = Table(element, doc)
            
            # בדיקה אם השם מופיע בתוך טבלה
            for row in table.rows:
                row_text = " ".join(c.text for c in row.cells)
                potential_name = get_name_after_colon(row_text)
                if potential_name:
                    current_student = potential_name
                    break

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
                    g_num = int(g_match[0])
                    
                    # חילוץ ההערה (התא הארוך ביותר)
                    notes = [c for i, c in enumerate(cells) if i != col_grade and i != col_sub and len(c) > 5]
                    note_text = max(notes, key=len) if notes else ""
                    
                    if note_text:
                        norm_note = clean_txt(note_text)
                        is_valid = False
                        if g_num in clean_bank:
                            # בדיקה אם לפחות משפט אחד מהבנק נמצא בתוך ההערה
                            for bank_txt in clean_bank[g_num]:
                                if bank_txt in norm_note or norm_note in bank_txt:
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
        st.warning("לא נמצאו נתונים.")
