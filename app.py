from flask import Flask, request, send_file, render_template
import pdfplumber
import pandas as pd
import re
from io import BytesIO
from openpyxl.styles import Alignment

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

CSV_COLUMNS = [
    'גוש', 'חלקה', 'תת חלקה', 'קומה', 'שטח במ"ר', 'כתובת מלאה',
    'שם בעל הדירה (מלא)', 'תז', 'הזכות בנכס',
    'הערות אזהרה/אחרות',
    'טלפון נייד', 'כתובת אי מייל', 'מגורים בדירה כן/לא',
    'קשיש / בעל מוגבלות', 'מספר נפשות', 'האם חתם על מסמך', 'האם חבר בנציגות'
]

def check_if_reversed(text):
    """בודק אם ה-PDF הפוך"""
    if not text: return False
    if 'שוג' in text or 'הקלח' in text or 'ת.ז' in text[::-1]:
        return True
    return False

def clean_and_reverse_name(name, needs_reversal):
    """מנקה שם והופך אותו אם צריך"""
    if not name: return ""

    if 'רשות הפיתוח' in name or 'חותיפה תוש ר' in name:
        return 'עמיגור'

    bad_words = ['מכר', 'ירושה', 'בשלמות', 'העברה', 'ת.ז', 'ז.ת', 'תז', 'זת', 'ללא', 'תמורה']
    for w in bad_words:
        name = name.replace(w, '')
        name = name.replace(w[::-1], '')

    name = re.sub(r'[0-9\.,:\-\(\)]', '', name).strip()

    if needs_reversal:
        name = name[::-1]

    return " ".join(name.split())

def extract_metadata_geometry(page):
    """חילוץ גוש, חלקה וכתובת לפי מיקום גיאומטרי"""
    gush = "Unknown"
    halaka = "Unknown"
    address = ""

    try:
        words = page.extract_words()
        # מיון לפי Y ואז X
        words.sort(key=lambda w: (w['top'], w['x0']))

        # בניית טקסט מלא לחיפוש כתובת
        full_text = page.extract_text() or ""

        # 1. חילוץ גוש וחלקה לפי קרבה
        for i, word in enumerate(words):
            txt = word['text'].replace(':', '')
            # גוש
            if txt in ['גוש', 'שוג']:
                if i+1 < len(words) and words[i+1]['text'].isdigit():
                    gush = words[i+1]['text']
                elif i-1 >= 0 and words[i-1]['text'].isdigit():
                    gush = words[i-1]['text']
            # חלקה
            if txt in ['חלקה', 'הקלח']:
                if i+1 < len(words) and words[i+1]['text'].isdigit():
                    halaka = words[i+1]['text']
                elif i-1 >= 0 and words[i-1]['text'].isdigit():
                    halaka = words[i-1]['text']

        # 2. חילוץ כתובת (לרוב בטקסט הרציף בשורות הראשונות)
        # מחפשים שורה שמתחילה ב"כתובת:"
        lines = full_text.split('\n')
        for line in lines:
            if 'כתובת' in line or 'תבותכ' in line:
                # ניקוי המילה "כתובת:"
                clean_addr = line.replace('כתובת:', '').replace(':תבותכ', '').replace('כתובת', '').strip()
                # אם השורה מכילה עברית הפוכה (אינדיקציה: 'ןולקשא' במקום 'אשקלון')
                if 'ןולקשא' in clean_addr or 'ביבא' in clean_addr or check_if_reversed(clean_addr):
                    clean_addr = clean_addr[::-1]

                if len(clean_addr) > 5:
                    address = clean_addr
                    break

    except Exception:
        pass

    return gush, halaka, address

def analyze_single_pdf(pdf_file_stream):
    try:
        with pdfplumber.open(pdf_file_stream) as pdf:
            # --- שלב 1: נתונים גלובליים ---
            page0 = pdf.pages[0]
            gush, halaka, address = extract_metadata_geometry(page0)

            parcel_number = halaka if halaka != "Unknown" else "Unknown"

            # בדיקת כיווניות כללית
            page0_text = page0.extract_text() or ""
            is_reversed = check_if_reversed(page0_text)

            # --- שלב 2: סריקת נתונים ---
            sub_parcels_data = {}
            current_sub = "0"
            sub_parcels_data["0"] = {'owners': [], 'notes': [], 'area': '', 'floor': ''}

            for page in pdf.pages:
                text = page.extract_text()
                if not text: continue
                lines = text.split('\n')

                for line in lines:
                    line = line.strip()
                    if not line: continue

                    # זיהוי תת חלקה
                    sub_m = re.search(r'תת חלקה\s*(\d+)', line) or re.search(r'(\d+)\s*הקלח\s*תת', line)
                    if sub_m:
                        current_sub = sub_m.group(1)
                        if current_sub not in sub_parcels_data:
                            sub_parcels_data[current_sub] = {'owners': [], 'notes': [], 'area': '', 'floor': ''}
                        continue

                    # זיהוי שטח וקומה
                    # מחפשים שורה עם מספר עשרוני ומילות מפתח
                    if any(x in line for x in ['שטח', 'חטש', 'קומה', 'המוק', 'מ"ר', 'ר"מ']):
                        # חילוץ שטח (מספר כמו 109, 44.70, 1,355)
                        # ננקה פסיקים קודם (עבור 1,355)
                        clean_line_nums = line.replace(',', '')
                        area_m = re.search(r'(\d+\.?\d*)', clean_line_nums)

                        # אנחנו מחפשים מספר שהוא כנראה השטח (ולא 1/12)
                        # לרוב השטח מופיע ליד המילה מ"ר.
                        # אסטרטגיה: אם יש מספר > 10 בשורה שיש בה "שטח" או "קומה", ניקח אותו.
                        all_nums = re.findall(r'(\d+\.?\d*)', clean_line_nums)
                        for num in all_nums:
                            if float(num) > 20 and float(num) < 10000: # טווח הגיוני לדירה
                                sub_parcels_data[current_sub]['area'] = num
                                break

                        # חילוץ קומה
                        floor = ""
                        if 'ראשונה' in line or 'הנושאר' in line: floor = 'ראשונה'
                        elif 'שניה' in line or 'הינש' in line: floor = 'שניה'
                        elif 'שלישית' in line or 'תישילש' in line: floor = 'שלישית'
                        elif 'רביעית' in line or 'תיעיבר' in line: floor = 'רביעית'
                        elif 'קרקע' in line or 'עקרק' in line: floor = 'קרקע'

                        if floor: sub_parcels_data[current_sub]['floor'] = floor
                        continue

                    # זיהוי הערות 126
                    if '126' in line or 'אזהרה' in line or 'הרהזא' in line:
                        clean_note = line
                        if 'הרהזא' in line or is_reversed: clean_note = clean_note[::-1]

                        beneficiary = ""
                        if 'עמיגור' in clean_note or 'רוגימע' in clean_note: beneficiary = "עמיגור"
                        elif 'רשות הפיתוח' in clean_note: beneficiary = "עמיגור (רשות הפיתוח)"
                        elif 'בנק' in clean_note: beneficiary = "בנק (משכנתה)"
                        else:
                            beneficiary = clean_note.replace('הערת אזהרה', '').replace('סעיף 126', '').strip()
                            beneficiary = re.sub(r'[0-9]+', '', beneficiary).strip()

                        if len(beneficiary) > 2:
                            note_text = f"הערת אזהרה (ס.126): {beneficiary}"
                            if note_text not in sub_parcels_data[current_sub]['notes']:
                                sub_parcels_data[current_sub]['notes'].append(note_text)
                        continue

                    # זיהוי בעלים
                    words = line.split()
                    found_id = None
                    for w in words:
                        clean_w = re.sub(r'[^\d]', '', w)
                        if 7 <= len(clean_w) <= 9:
                            if '/' not in w and '\\' not in w and '.' not in w:
                                found_id = clean_w
                                break

                    if found_id:
                        if any(k in line for k in ['בנק', 'קנב', 'משכנתה', 'התנכשמ', 'לפקודת', 'תדוקפל']): continue

                        name_part = line.replace(found_id, '')
                        name_part = re.sub(r'\d+/\d+', '', name_part)
                        final_name = clean_and_reverse_name(name_part, is_reversed)

                        if len(final_name) < 2: continue

                        sub_parcels_data[current_sub]['owners'].append({
                            'name': final_name,
                            'id': found_id
                        })

            # Flattening
            final_output_rows = []
            sorted_keys = sorted(sub_parcels_data.keys(), key=lambda x: int(x) if x.isdigit() else 0)

            for sp_num in sorted_keys:
                data = sub_parcels_data[sp_num]

                # אם לא נמצאו בעלים (למשל תת חלקה 0 שהיא רק כותרת), מדלגים
                if not data['owners']: continue

                notes_str = " | ".join(data['notes'])

                for owner in data['owners']:
                    row = {
                        'גוש': gush,
                        'חלקה': halaka,
                        'תת חלקה': sp_num,
                        'קומה': data['floor'],
                        'שטח במ"ר': data['area'],
                        'כתובת מלאה': address,
                        'שם בעל הדירה (מלא)': owner['name'],
                        'תז': owner['id'],
                        'הזכות בנכס': 'בעלות',
                        'הערות אזהרה/אחרות': notes_str,
                        'טלפון נייד': '', 'כתובת אי מייל': '', 'מגורים בדירה כן/לא': '',
                        'קשיש / בעל מוגבלות': '', 'מספר נפשות': '', 'האם חתם על מסמך': '', 'האם חבר בנציגות': ''
                    }

                    # בדיקת כפילויות
                    is_duplicate = False
                    for existing in final_output_rows:
                        if existing['תת חלקה'] == sp_num and existing['תז'] == owner['id']:
                            is_duplicate = True
                            break

                    if not is_duplicate:
                        final_output_rows.append(row)

            return final_output_rows, parcel_number

    except Exception as e:
        print(f"Error parsing file: {e}")
        return [], "Error"

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    uploaded_files = request.files.getlist('tabu_files')
    if not uploaded_files or uploaded_files[0].filename == '': return 'No files', 400

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        files_processed = 0
        existing_sheet_names = []

        for file in uploaded_files:
            if file and file.filename.endswith('.pdf'):
                file_stream = BytesIO(file.read())
                data, parcel_num = analyze_single_pdf(file_stream)

                if data:
                    df = pd.DataFrame(data, columns=CSV_COLUMNS)
                    sheet_name = f"חלקה {parcel_num}"

                    # טיפול בשמות גיליונות
                    sheet_name = re.sub(r'[\\/*?:\[\]]', '', sheet_name)[:30]
                    counter = 1
                    original_name = sheet_name
                    while sheet_name in existing_sheet_names:
                        sheet_name = f"{original_name} ({counter})"
                        counter += 1

                    existing_sheet_names.append(sheet_name)

                    # כתיבה לאקסל
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    # --- עיצוב מימין לשמאל (RTL) ויישור ---
                    worksheet = writer.sheets[sheet_name]
                    worksheet.sheet_view.rightToLeft = True # הגדרת הגיליון כ-RTL

                    # יישור טקסט לימין בכל התאים
                    for row in worksheet.iter_rows():
                        for cell in row:
                            cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=False)

                    # הרחבת עמודות אוטומטית (בסיסית)
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except: pass
                        adjusted_width = (max_length + 2) * 1.2
                        worksheet.column_dimensions[column_letter].width = min(adjusted_width, 50) # מקסימום 50 רוחב

                    files_processed += 1

        if files_processed == 0:
            return "לא נמצאו נתונים תקינים", 400

    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='Tabu_Final_RTL.xlsx')

if __name__ == '__main__':
    app.run(debug=True)