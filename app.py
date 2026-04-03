import streamlit as st
import pandas as pd
from docx import Document
import re

# --- עיצוב ---
st.set_page_config(page_title="בודק תעודות - אורט פסגות", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)
st.title("🍎 בודק תעודות - הגרסה היציבה והמלאה")

# פונקציית ניקוי פשוטה - רק להשוואה
def clean_txt(text):
    if not text: return ""
    t = str(text).strip()
    # מסיר פיסוק וגרשיים כדי שההשוואה תהיה קלה
    t = re.sub(r'[.!?,:;\"״”’\'\-]', '', t)
    t = " ".join(t.split())
    # השוואת מגדר בסיסית (הופך הכל למכנה משותף)
    t = t.replace("ייך", "ך").replace("יך", "ך").replace("הינך", "הנך")
    return t

# --- בנק המשפטים המלא (כל מה ששלחת) ---
GRADE_BANK = {
    40: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע", "היעדרויותיך הרבות פגעו בתפקודך ובהישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "מעורבות בשיעורים הגשת כל המטלות והשקעת מאמצים לימודיים יקדמו אותך וישפרו את הבנתך בחומר הנלמד", "את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עלייך לגלות אחריות על למידתך להגיע בזמן לשיעורים ולבצע משימות באופן עקבי", "עלייך להתאמץ יותר ולתפקד בשיעורים", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה", "עלייך להקפיד להגיש את המטלות הנדרשות"],
    50: ["עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "עליך להיות נוכח בשיעורים נוכחות סדירה תקדם למידתך ותשפר הישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "היה עליך להקפיד על הגשת העבודות ולגלות מעורבות ואחריות ללמידה", "התנהגותך בשיעורים וחוסר הריכוז פגעו בהישגיך הלימודיים", "I believe you can do better"],
    60: ["את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים"],
    65: ["את ילדה שקטה ונעימת הליכות", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים עלייך להימנע מהעדרויות", "הקפידי להשתתף באופן פעיל בשיח הקבוצתי על מנת לשפר את הבנתך ואת הישגייך ביכולתך לתרום מרעיונותייך לקבוצה"],
    75: ["ניכר שהשקעת זמן ומאמצים בהכנת שיעורי בית ועבודות והגעת להישגים נאים", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות", "הנך משתתפת באופן פעיל בשיעורי חנ'ג", "הנך מקפידה על נוכחות סדירה שותפה פעילה בשיעורים"],
    80: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות", "הנך נוכח בשיעורים מבצע את כל הנדרש"],
    85: ["את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו", "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות"],
    90: ["הנך נוכחת בשיעורים מבצעת את כל הנדרש בקביעות וביסודיות", "את תלמידה רצינית מגלה עניין והבנה ובעלת מוטיבציה להצלחה", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות", "הנך מגלה אחריות ורצינות בלמידה"],
    95: ["את ראויה לשבח על הישגייך המצוינים", "יכולתך להתבטא בכתב ראויה לשבח"],
    100: ["את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו", "הנך תלמידה מצטיינת ויחסך למקצוע רציני", "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה"]
}

# בנק נקי להשוואה
clean_bank = {k: [clean_txt(n) for n in v] for k, v in GRADE_BANK.items()}

uploaded_file = st.file_uploader("העלי קובץ תעודות (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    results = []
    student_name = "לא זוהה"

    # סריקה לאיתור נתונים
    for element in doc.element.body.iter():
        # זיהוי שם התלמיד/ה
        if element.tag.endswith('p'):
            p_text = "".join(t.text for t in element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') if t.text)
            if any(k in p_text for k in ["שם", "לכבוד", "תלמיד"]):
                name = re.sub(r'(שם התלמיד|התלמידה|:|לכבוד|שם)', '', p_text).strip()
                if name: student_name = name

        # זיהוי טבלה
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
                    
                    # חילוץ ציון
                    g_match = re.findall(r'\d+', cells[col_grade])
                    if not g_match: continue
                    g_num = int(g_match[0])
                    
                    # חילוץ הערה (התא הארוך ביותר)
                    note_candidates = [c for i, c in enumerate(cells) if i != col_grade and i != col_sub and len(c) > 5]
                    note_text = max(note_candidates, key=len) if note_candidates else ""
                    
                    if note_text:
                        norm_note = clean_txt(note_text)
                        found = False
                        
                        # בדיקה אם לפחות משפט אחד מהבנק נמצא בטקסט (תומך במגדר ושילובים)
                        if g_num in clean_bank:
                            for b_txt in clean_bank[g_num]:
                                if b_txt in norm_note or norm_note in b_txt:
                                    found = True
                                    break
                        
                        results.append({
                            "תלמיד/ה": student_name,
                            "מקצוע": cells[col_sub],
                            "ציון": g_num,
                            "הערה בתעודה": note_text,
                            "סטטוס": "✅ תקין" if found else "❌ שגיאה"
                        })

    if results:
        st.table(pd.DataFrame(results))
    else:
        st.warning("לא נמצאו ציונים בקובץ.")
