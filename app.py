import streamlit as st
import pandas as pd
from docx import Document
import re
import pdfplumber
import difflib

# --- הגדרות עיצוב ---
st.set_page_config(page_title="בודק תעודות - אורט פסגות", page_icon="🍎", layout="wide")
st.markdown("""<style>.main { text-align: right; direction: rtl; }</style>""", unsafe_allow_html=True)

st.title("🍎 מערכת בדיקת תעודות חכמה")
st.write("העלי את קובץ התעודות (Word) ואופציונלית את ה-PDF של הבנק. המערכת תזהה התאמות גם בין זכר לנקבה.")

# --- פונקציית נירמול חזקה ---
def normalize_hebrew(text):
    if not text: return ""
    t = str(text).strip()
    t = " ".join(t.split()) # ניקוי רווחים כפולים
    
    # האחדת כל סוגי הגרשיים לגרש בודד
    for char in ['"', '״', '”', '“', '’', '‘', '´']:
        t = t.replace(char, "'")
    
    # ניקוי סימני פיסוק בסיסיים לצורך ההשוואה
    for char in ['.', ',', '!', '?', ':', ';', '-']:
        t = t.replace(char, "")
        
    # טיפול ביו"ד כפולה/בודדת ובהטיות נפוצות
    t = t.replace("ייך", "ך").replace("יך", "ך")
    t = t.replace("הינך", "הנך").replace("עלייך", "עלך").replace("עליך", "עלך")
    
    return t.strip()

# --- 1. בנק המשפטים הידני שלך ---
GRADE_BANK = {
    40: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע", "היעדרויותיך הרבות פגעו בתפקודך ובהישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "מעורבות בשיעורים הגשת כל המטלות והשקעת מאמצים לימודיים יקדמו אותך וישפרו את הבנתך בחומר הנלמד", "את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עלייך לגלות אחריות על למידתך להגיע בזמן לשיעורים ולבצע משימות באופן עקבי", "עלייך להתאמץ יותר ולתפקד בשיעורים", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה", "עלייך להקפיד להגיש את המטלות הנדרשות"],
    45: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "למידה עקבית השקעת מאמצים והתמקדות בחומר הלימוד ישפרו את ידיעותייך ויקדמו אותך להישגים טובים במקצוע"],
    50: ["עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "עליך להיות נוכח בשיעורים נוכחות סדירה תקדם למידתך ותשפר הישגיך", "ציונך נפגע עקב היעדרויותיך הרבות", "היה עליך להקפיד על הגשת העבודות ולגלות מעורבות ואחריות ללמידה", "התנהגותך בשיעורים וחוסר הריכוז פגעו בהישגיך הלימודיים", "I believe you can do better"],
    55: ["לא השקעת מספיק בעבודת הבית והדבר יצר פערים בהבנה ובידיעת החומר", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים"],
    60: ["את מתקשה בהבנת החומר חשוב מאוד שתשאלי כשאינך מבינה", "עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "עליך להקפיד על כתיבת תשובה מפורטת ומנומקת"],
    65: ["את ילדה שקטה ונעימת הליכות היי קשובה בשיעורים ואל תחששי לדבר בעברית", "עלייך לגלות יותר מוטיבציה ואחריות ללמידה עקביות בלמידה והכנת כל המשימות היו מובילות להישגים גבוהים יותר", "הנך מגלה מוטיבציה ורצון להתקדם בלימודים עלייך להימנע מהעדרויות", "הקפידי להשתתף באופן פעיל בשיח הקבוצתי על מנת לשפר את הבנתך ואת הישגייך ביכולתך לתרום מרעיונותייך לקבוצה"],
    70: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "עלייך להקפיד על תלבושת ספורט כנדרש", "את מתקשה בהבנת החומר"],
    75: ["ניכר שהשקעת זמן ומאמצים בהכנת שיעורי בית ועבודות והגעת להישגים נאים", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות", "הנך משתתפת באופן פעיל בשיעורי חנ'ג", "הנך מקפידה על נוכחות סדירה שותפה פעילה בשיעורים תורמת לשיח ומקדמת את הלמידה בקבוצה"],
    80: ["הנך מגלה מוטיבציה ורצון להתקדם בלימודים", "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות", "הנך נוכח בשיעורים מבצע את כל הנדרש בשיעור בקביעות וביסודיות"],
    85: ["את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו", "לקחת חלק פעיל בשיעורים וגילית רצינות ואחריות"],
    90: ["הנך נוכחת בשיעורים מבצעת את כל הנדרש בשיעור בקביעות וביסודיות", "את תלמידה רצינית מגלה עניין והבנה ובעלת מוטיבציה להצלחה", "שקדת על עבודתך ברצינות מתוך אחריות ובגרות"],
    95: ["את ראויה לשבח על הישגייך המצוינים", "את בעלת מוטיבציה פנימית וחשוב לך להצליח ולהתקדם בלמידה"],
    100: ["את ראויה להערכה רבה על מאמצייך הלימודיים במחצית זו", "הנך תלמידה מצטיינת ויחסך למקצוע רציני", "הפגנת ידע והתמודדת עם אתגרים ברמת חשיבה גבוהה"]
}

# --- 2. טעינת תוספות מה-PDF ---
active_bank = {k: [normalize_hebrew(n) for n in v] for k, v in GRADE_BANK.items()}

st.sidebar.header("בנק חיצוני")
pdf_file = st.sidebar.file_uploader("העלי PDF של הבנק (אופציונלי)", type=['pdf'])

if pdf_file:
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    clean_row = [str(c).strip() for c in row if c]
                    if len(clean_row) < 2: continue
                    g_val = None
                    for cell in clean_row:
                        digits = re.findall(r'\d+', cell)
                        if digits: g_val = int(digits[0]); break
                    if g_val:
                        text_cells = [c for c in clean_row if not re.fullmatch(r'\d+', c)]
                        if text_cells:
                            n_val = normalize_hebrew(max(text_cells, key=len))
                            if g_val not in active_bank: active_bank[g_val] = []
                            if n_val not in active_bank[g_val]: active_bank[g_val].append(n_val)

# --- 3. עיבוד קובץ התעודות (Word) ---
uploaded_file = st.file_uploader("העלי קובץ תעודות (Word)", type=['docx'])

if uploaded_file:
    doc = Document(uploaded_file)
    all_results = []
    current_student = "לא זוהה שם"
    
    for child in doc._element.body.iter():
        if child.tag.endswith('p'):
            para_text = "".join(t.text for t in child.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t') if t.text)
            if any(k in para_text for k in ["שם התלמיד", "שם התלמידה", "שם:", "לכבוד"]):
                name = re.sub(r'[^א-ת\s]', '', para_text).replace("שם התלמיד", "").replace("התלמידה", "").strip()
                if len(name.split()) >= 2: current_student = name

        elif child.tag.endswith('tbl'):
            from docx.table import Table
            table = Table(child, doc)
            if len(table.rows) < 2: continue
            
            headers = [c.text.strip() for c in table.rows[0].cells]
            col_grade = next((idx for idx, h in enumerate(headers) if "ציון" in h), -1)
            col_sub = next((idx for idx, h in enumerate(headers) if "מקצוע" in h), 0)

            if col_grade != -1:
                for row in table.rows[1:]:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) > col_grade:
                        grade_str = cells[col_grade]
                        # שליפת ההערה (התא הכי ארוך שאינו המקצוע או הציון)
                        text_candidates = [c for i, c in enumerate(cells) if i != col_grade and i != col_sub and len(c) > 5]
                        note = max(text_candidates, key=len) if text_candidates else ""
                        
                        if not (grade_str and note): continue
                        
                        grade_match = re.findall(r'\d+', str(grade_str))
                        grade_num = int(grade_match[0]) if grade_match else 0
                        
                        status = "❌ לא נמצא בבנק"
                        detail = "הערה לא מזוהה"
                        
                        if grade_num in active_bank:
                            note_norm = normalize_hebrew(note)
                            
                            # בדיקה 1: דמיון גבוה (מטפל בזכר/נקבה)
                            best_match_ratio = 0
                            for opt in active_bank[grade_num]:
                                ratio = difflib.SequenceMatcher(None, note_norm, opt).ratio()
                                if ratio > best_match_ratio: best_match_ratio = ratio
                            
                            if best_match_ratio > 0.82: # סף רגישות לשינויי מגדר
                                status = "✅ תקין"
                                detail = f"התאמה גבוהה ({int(best_match_ratio*100)}%)"
                            else:
                                # בדיקה 2: האם ההערה מורכבת מ-2 משפטים מהבנק?
                                found_parts = 0
                                temp_note = note_norm
                                for opt in sorted(active_bank[grade_num], key=len, reverse=True):
                                    if opt in temp_note:
                                        temp_note = temp_note.replace(opt, "")
                                        found_parts += 1
                                
                                if found_parts >= 1:
                                    status = "✅ תקין (שילוב)"
                                    detail = f"זוהו {found_parts} משפטים מהבנק"
                                else:
                                    status = "❌ בדיקה"
                                    detail = "לא תואם למגדר או לניסוח בבנק"

                        all_results.append({
                            "תלמיד/ה": current_student,
                            "מקצוע": cells[col_sub] if col_sub < len(cells) else "---",
                            "ציון": grade_str,
                            "הערה בתעודה": note,
                            "סטטוס": status,
                            "פירוט": detail
                        })

    if all_results:
        st.table(pd.DataFrame(all_results))
    else:
        st.info("העלי קובץ כדי להתחיל בסריקה.")
