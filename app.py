import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - פסגות", page_icon="📑")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 בודק תעודות חכם - אורט פסגות")
st.write("המערכת סורקת את טבלת הציונים לפי העמודות 'ציון' ו-'הערכה מילולית'.")

# --- בנק ההערות המובנה ---
# (כאן יש להכניס את כל הרשימה המלאה מה-PDF ששלחת)
POSITIVE_NOTES = [
    "אתה ראוי לשבח על הישגיך המצויינים", "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה",
    "גילית יכולת טובה בניתוח טקסטים ויישום הידע", "הנך בעל ידע עולם נרחב",
    "הישגיך בשפה טובים מאוד - יישר כוח!", "הנך תלמיד מצטיין ויחסך למקצוע רציני",
    "את ראויה לשבח על הישגיך המצוינים", "את מבינה היטב את הנקרא... ומסיקה מסקנות"
]

IMPROVEMENT_NOTES = [
    "עליך לגלות אחריות על למידתך", "עליך לגלות יותר מוטיבציה ואחריות ללמידה",
    "אתה עדיין מתקשה בהבנת החומר", "עליך לשפר את מיומנויות הבנת הנקרא",
    "את מתקשה בפיתוח נושא ובניסוחו בכתב", "עליך לגלות אחריות על למידתך... ולבצע משימות באופן עקבי"
]

FULL_BANK = POSITIVE_NOTES + IMPROVEMENT_NOTES

def analyze_row(subject, grade_str, note):
    reasons = []
    note = str(note).strip()
    
    # ניקוי הציון - השארת מספרים בלבד
    grade_digits = "".join(filter(str.isdigit, str(grade_str)))
    
    if not grade_digits or not note or len(note) < 3:
        return None 

    grade_num = int(grade_digits)
    
    # 1. בדיקת קיום בבנק
    if not any(b in note for b in FULL_BANK):
        reasons.append("הערה חופשית (לא מהבנק)")
    
    # 2. בדיקת סתירה
    if grade_num <= 55 and any(b in note for b in POSITIVE_NOTES):
        reasons.append(f"סתירה: ציון נמוך ({grade_num}) עם הערה חיובית")
    if grade_num >= 90 and any(b in note for b in IMPROVEMENT_NOTES):
        reasons.append(f"סתירה: ציון גבוה ({grade_num}) עם הערת שיפור")
        
    return reasons if reasons else None

# --- העלאת קבצים ---
uploaded_file = st.file_uploader("בחרי קובץ Word (תעודה)", type=['docx'])

if uploaded_file:
    anomalies = []
    doc = Document(uploaded_file)
    
    for table in doc.tables:
        # זיהוי עמודות לפי כותרת
        header_cells = [cell.text.strip() for cell in table.rows[0].cells]
        
        col_grade = -1
        col_note = -1
        col_subject = 0 # ברירת מחדל עמודה ראשונה
        
        for i, header in enumerate(header_cells):
            if "ציון" in header: col_grade = i
            if "הערכה מילולית" in header: col_note = i
        
        # אם לא מצאנו את העמודות המתאימות, דלג על הטבלה הזו
        if col_grade == -1 or col_note == -1:
            continue
            
        for row in table.rows[1:]:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) > max(col_grade, col_note):
                subject = cells[col_subject]
                grade = cells[col_grade]
                note = cells[col_note]
                
                res = analyze_row(subject, grade, note)
                if res:
                    anomalies.append({
                        "מקצוע": subject,
                        "ציון": grade,
                        "הערכה": (note[:50] + '...') if len(note) > 50 else note,
                        "חריגה": " | ".join(res)
                    })

    if anomalies:
        st.warning(f"נמצאו {len(anomalies)} חריגות:")
        st.table(pd.DataFrame(anomalies))
        
        # יצירת WORD להורדה
        doc_out = Document()
        doc_out.add_heading('דוח חריגות בתעודה', 0)
        t = doc_out.add_table(rows=1, cols=4)
        t.style = 'Table Grid'
        for i, h in enumerate(['מקצוע', 'ציון', 'הערה', 'חריגה']): t.rows[0].cells[i].text = h
        for a in anomalies:
            row_cells = t.add_row().cells
            row_cells[0].text, row_cells[1].text, row_cells[2].text, row_cells[3].text = str(a['מקצוע']), str(a['ציון']), a['הערה'], a['חריגה']
        
        buffer = BytesIO()
        doc_out.save(buffer)
        st.download_button("📥 הורדי דוח Word", buffer.getvalue(), "report.docx")
    else:
        st.success("לא נמצאו חריגות בטבלת הציונים.")
