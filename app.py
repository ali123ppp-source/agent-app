import streamlit as st
import pandas as pd
from io import BytesIO
import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
from docx.shared import RGBColor, Pt, Inches
import re
import html
from datetime import datetime
from weasyprint import HTML as WeasyHTML

# أرقام هندية-عربية وفارسية شائعة بالملفات الرسمية — لازم تتوحّد لأرقام
# لاتينية قبل استخدامها كمفتاح مطابقة، وإلا "١٢٣٤٥٦٧" و"1234567" يُعتبران
# بطاقتين مختلفتين لنفس العائلة (يظهر محذوف بملف ومضاف بالثاني بالغلط).
_DIGIT_TRANSLATION = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")

def normalize_digits(value):
    return str(value).translate(_DIGIT_TRANSLATION)

def esc(value):
    """يهرّب أي نص مصدره الملف (اسم، رقم بطاقة، نص إحالة) قبل حقنه داخل
    HTML — يمنع كسر تخطيط الجدول أو حقن ماركب غير مقصود لو احتوى اسم على
    أحرف زي &lt; أو &amp;."""
    return html.escape(str(value), quote=True)


# -----------------------------------------------------------------------------
# 1. محرك الاستشعار الزمني المحدث (يدعم الملفات ككائنات)
# -----------------------------------------------------------------------------
def extract_document_date(file_obj):
    file_ext = file_obj.name.split('.')[-1].lower()
    if file_ext != 'docx':
        return None  # الإكسل لا يحتوي على فقرات نصية قياسية للتاريخ بنفس صيغة الوورد

    doc = Document(file_obj)
    file_obj.seek(0) # إعادة مؤشر القراءة للبداية بعد استخدام الملف
    
    patterns = [
        r"([A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})", 
        r"(\d{1,2}[/-]\d{1,2}[/-]\d{4})",               
        r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})",               
        r"(\d{1,2}\s+[\u0600-\u06FF]+\s+\d{4})"         
    ]
    for section in doc.sections:
        if section.footer:
            for para in reversed(section.footer.paragraphs):
                for pattern in patterns:
                    match = re.search(pattern, para.text)
                    if match:
                        try:
                            d_str = match.group(1)
                            if "-" in d_str or "/" in d_str: return pd.to_datetime(d_str, dayfirst=True).to_pydatetime()
                            if re.search(r"[\u0600-\u06FF]", d_str): return datetime.now()
                            return datetime.strptime(d_str, "%A, %B %d, %Y")
                        except: continue
                        
    paragraphs = doc.paragraphs[-50:] if len(doc.paragraphs) > 50 else doc.paragraphs
    for para in reversed(paragraphs):
        for pattern in patterns:
            match = re.search(pattern, para.text)
            if match:
                try:
                    d_str = match.group(1)
                    if "-" in d_str or "/" in d_str: return pd.to_datetime(d_str, dayfirst=True).to_pydatetime()
                    if re.search(r"[\u0600-\u06FF]", d_str): return datetime.now()
                    return datetime.strptime(d_str, "%A, %B %d, %Y")
                except: continue
    
    try:
        if doc.core_properties.modified:
            return doc.core_properties.modified.replace(tzinfo=None)
    except:
        pass
    return None

# -----------------------------------------------------------------------------
# 2. محرك الاستخراج الدقيق المحدث (يدعم Word و Excel - للنماذج 1، 2، و3)
# -----------------------------------------------------------------------------
def extract_clean_records(file_obj, card_type="old"):
    records = {}
    duplicates = []
    file_ext = file_obj.name.split('.')[-1].lower()
    rows_data = []

    if file_ext == 'docx':
        doc = Document(file_obj)
        file_obj.seek(0)

        # استخراج الفقرات (للأنماط التي لا تعتمد على الجداول)
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text: continue
            cells = [c.strip() for c in text.split(',')]
            if len(cells) >= 6 and any(char.isdigit() for char in cells[0]) and any('\u0600' <= char <= '\u06FF' for char in cells[3]):
                try:
                    withheld, eligible, total, name = int(cells[0]), int(cells[1]), int(cells[2]), cells[3]
                    old_card = normalize_digits(cells[4])
                    new_card = normalize_digits(cells[5]) if len(cells) > 5 else old_card
                    selected_card = old_card if card_type == "old" else new_card
                    alt_card = new_card if card_type == "old" else old_card
                    if alt_card == selected_card:
                        alt_card = ""
                    seq = cells[6] if len(cells) > 6 else "-"
                    if selected_card:
                        if selected_card in records:
                            duplicates.append((selected_card, name))
                        else:
                            records[selected_card] = {"seq": seq, "name": name, "total": total, "eligible": eligible, "withheld": withheld, "alt_card": alt_card}
                except ValueError: continue

        # استخراج جداول الوورد
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                rows_data.append(cells)

    elif file_ext == 'xlsx':
        xls = pd.ExcelFile(file_obj)
        file_obj.seek(0)
        for sheet_name in xls.sheet_names:
            df_excel = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            for row in df_excel.values:
                cells = []
                for cell in row:
                    if pd.isna(cell):
                        continue
                    if isinstance(cell, float) and cell.is_integer():
                        cells.append(str(int(cell)))
                    else:
                        cells.append(str(cell).strip().replace('\n', ' '))
                rows_data.append(cells)

    # تطبيق نفس منطق التنظيف والاستخراج المعتاد على الأسطر إذا لم يتم سحب البيانات من الفقرات
    if not records:
        for cells in rows_data:
            if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "اسم رب" in "".join(cells): continue
            parsed = _heuristic_parse_row(cells, card_type)
            if parsed is None:
                continue
            selected_card = parsed["selected_card"]
            if selected_card in records:
                duplicates.append((selected_card, parsed["name"]))
            else:
                records[selected_card] = {k: v for k, v in parsed.items() if k != "selected_card"}

    return records, duplicates

def _heuristic_parse_row(cells, card_type="old"):
    """يفسّر صف بيانات خام (بلا أي اعتماد على نص عناوين) بترتيب ثابت شوهد
    فعلياً بملفات حقيقية: [محجوب, مستحق, كلي, اسم, ...بطاقات بخانات ≥5،
    تسلسل بآخر عمود رقمي]. يُستخدم كخط دفاع أخير من extract_clean_records
    وأيضاً لبناء عيّنة معاينة لملفات عناوينها مدمجة بخلية واحدة (preview_columns_for_file)
    حيث يفشل التعرف على عناوين منفصلة بالكامل. يرجع None لو الصف ما يطابق الشكل."""
    cells = [normalize_digits(c) if c.strip() and c.strip().isdigit() else c for c in cells]
    name_idx = -1
    max_len = 0
    for i, c in enumerate(cells):
        if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
            if len(c) > max_len: max_len, name_idx = len(c), i
    if name_idx == -1:
        return None
    card_indices = [i for i, c in enumerate(cells) if c.isdigit() and len(c) >= 5]
    if not card_indices:
        return None

    old_card = cells[card_indices[0]]
    new_card = cells[card_indices[-1]] if len(card_indices) > 1 else old_card
    selected_card = old_card if card_type == "old" else new_card
    alt_card = new_card if card_type == "old" else old_card
    if alt_card == selected_card:
        alt_card = ""
    seq = "-"
    for i in range(len(cells)-1, card_indices[-1], -1):
        if cells[i].isdigit():
            seq = cells[i]
            break
    digit_cells = [int(cells[i]) for i in range(name_idx) if cells[i].isdigit()]
    if len(digit_cells) >= 3: withheld, eligible, total = digit_cells[0], digit_cells[1], digit_cells[2]
    elif len(digit_cells) == 2: withheld, eligible, total = 0, digit_cells[0], digit_cells[1]
    else: return None
    return {"seq": seq, "name": cells[name_idx], "total": total, "eligible": eligible, "withheld": withheld, "alt_card": alt_card, "selected_card": selected_card}

# -----------------------------------------------------------------------------
# 2.5. محرك الاستخراج المخصص للنموذج الرابع المحدث (المستحق فقط)
# -----------------------------------------------------------------------------
def extract_eligible_only_records(file_obj):
    records = {}
    file_ext = file_obj.name.split('.')[-1].lower()
    rows_data = []

    if file_ext == 'docx':
        doc = Document(file_obj)
        file_obj.seek(0)
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                rows_data.append(cells)
                
    elif file_ext == 'xlsx':
        xls = pd.ExcelFile(file_obj)
        file_obj.seek(0)
        for sheet_name in xls.sheet_names:
            df_excel = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            for row in df_excel.values:
                cells = []
                for cell in row:
                    if pd.isna(cell):
                        continue
                    if isinstance(cell, float) and cell.is_integer():
                        cells.append(str(int(cell)))
                    else:
                        cells.append(str(cell).strip().replace('\n', ' '))
                rows_data.append(cells)

    duplicates = []
    for cells in rows_data:
        # التأكد من وجود 6 أعمدة على الأقل حسب الترتيب المطلوب
        if len(cells) >= 6:
            if "اسم" in cells[3] or "المركز" in cells[0]: continue

            seq = cells[0]
            old_card = normalize_digits(cells[2])
            name = cells[3]
            eligible_str = cells[5]

            if old_card.isdigit() and len(old_card) >= 4:
                try:
                    el_val = int(''.join(filter(str.isdigit, eligible_str)))
                    if old_card in records:
                        duplicates.append((old_card, name))
                        continue
                    records[old_card] = {
                        "seq": seq,
                        "name": name,
                        "total": 0,       # تصفير الكلي لتجاهله
                        "eligible": el_val,
                        "withheld": 0     # تصفير المحجوب لتجاهله
                    }
                except ValueError:
                    continue
    return records, duplicates

# -----------------------------------------------------------------------------
# 2.6. محرك الاستخراج الذكي بالتعرف التلقائي على العناوين (النموذج الخامس)
#      يفهم أعمدة أي جدول (وورد أو إكسل) عبر قراءة صف العناوين مهما كان ترتيبها
# -----------------------------------------------------------------------------
def _smart_match_header_role(header_text):
    h = str(header_text).strip()
    if not h:
        return None
    if "تسلسل" in h or h == "ت" or h.startswith("ت ") or h.startswith("ت(") or h.startswith("ت-"):
        return "seq"
    if "قديم" in h or "سابق" in h:
        return "old_card"
    if "جديد" in h or "حديث" in h:
        return "new_card"
    if "محجوب" in h:
        return "withheld"
    if "مستحق" in h:
        return "eligible"
    if "كلي" in h:
        return "total"
    if "بطاقة" in h:
        return "card_generic"
    if "اسم" in h:
        return "name"
    return None

def _smart_detect_header_map(rows_data, max_scan=3):
    for r_idx in range(min(max_scan, len(rows_data))):
        role_map = {}
        for c_idx, cell in enumerate(rows_data[r_idx]):
            role = _smart_match_header_role(cell)
            if role and role not in role_map:
                role_map[role] = c_idx
        has_card = any(k in role_map for k in ("old_card", "new_card", "card_generic"))
        has_amount = any(k in role_map for k in ("total", "eligible", "withheld"))
        if "name" in role_map and has_card and has_amount:
            return role_map, r_idx
    return None, -1

def _smart_clean_card(value):
    v = normalize_digits(str(value).strip())
    if not v.isdigit() or len(v) < 4:
        return ""
    return v

def _smart_to_int(value):
    digits = "".join(filter(str.isdigit, str(value)))
    return int(digits) if digits else 0

def _read_tables_rows(file_obj):
    """يقرأ كل الجداول/الأوراق بملف (docx أو xlsx) كقوائم صفوف نصية خام،
    بدون أي تفسير لمعنى الأعمدة — الأساس المشترك بين الاستخراج الفعلي
    ومعاينة الأعمدة قبل المقارنة."""
    file_ext = file_obj.name.split('.')[-1].lower()
    tables_rows = []

    if file_ext == 'docx':
        doc = Document(file_obj)
        file_obj.seek(0)
        for table in doc.tables:
            rows = [[cell.text.strip().replace('\n', ' ') for cell in row.cells] for row in table.rows]
            tables_rows.append(rows)

    elif file_ext == 'xlsx':
        xls = pd.ExcelFile(file_obj)
        file_obj.seek(0)
        for sheet_name in xls.sheet_names:
            df_excel = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            rows = []
            for row in df_excel.values:
                cells = []
                for cell in row:
                    if pd.isna(cell):
                        cells.append("")
                    elif isinstance(cell, float) and cell.is_integer():
                        cells.append(str(int(cell)))
                    else:
                        cells.append(str(cell).strip().replace('\n', ' '))
                rows.append(cells)
            tables_rows.append(rows)

    return tables_rows

_ROLE_LABELS_AR = {
    "seq": "التسلسل", "old_card": "البطاقة القديمة", "new_card": "البطاقة الجديدة",
    "card_generic": "رقم بطاقة", "name": "اسم رب الأسرة", "total": "الأفراد الكلية",
    "eligible": "الأفراد المستحقة", "withheld": "الأفراد المحجوبين",
}

def preview_columns_for_file(file_obj):
    """يكتشف تخطيط أعمدة الملف (نفس منطق النموذج الخامس) بدون استخراج
    كامل السجلات، ويرجع قاموس {دور: نص العنوان المكتشف} + أول صفين بيانات
    مقروءة حسب الدور (مو أعمدة خام) كعيّنة، لعرضها على المستخدم قبل ما
    يبدأ المقارنة فعلياً. يرجع None لو ما لقى جدول بعناوين واضحة."""
    tables_rows = _read_tables_rows(file_obj)
    for rows in tables_rows:
        role_map, header_row_idx = _smart_detect_header_map(rows)
        if role_map is None:
            continue
        header_row = rows[header_row_idx]
        detected = {}
        for role, col_idx in role_map.items():
            if role == "card_generic" and ("old_card" in role_map or "new_card" in role_map):
                continue  # ما نعرضه لو عندنا دور أدق (قديم/جديد) لنفس أو عمود مختلف
            label = _ROLE_LABELS_AR.get(role, role)
            header_text = header_row[col_idx] if col_idx < len(header_row) else ""
            detected[label] = header_text

        sample_records = []
        for row in rows[header_row_idx + 1: header_row_idx + 3]:
            sample = {}
            for role, col_idx in role_map.items():
                if role == "card_generic" and ("old_card" in role_map or "new_card" in role_map):
                    continue
                label = _ROLE_LABELS_AR.get(role, role)
                sample[label] = row[col_idx] if col_idx < len(row) else ""
            sample_records.append(sample)

        return {"detected": detected, "sample_records": sample_records}

    # لو ما نفع التعرف على عناوين منفصلة بأي جدول، نفحص هل عناوينه مدمجة
    # كلها بخلية وحدة (صيغة ملفات قديمة شائعة — رأس الجدول نص طويل واحد
    # بدل عمود لكل حقل) قبل ما نستسلم كلياً ونطلع تحذير "ما فيه جدول واضح"
    # المضلل، بينما فعلياً فيه عناوين وبيانات سليمة بس بشكل مختلف.
    for rows in tables_rows:
        header_row_idx, labels = _detect_merged_header_row(rows)
        if header_row_idx is None:
            continue
        detected = {_ROLE_LABELS_AR.get(role, role): text for role, text in labels}
        sample_records = []
        for row in rows[header_row_idx + 1:]:
            if len(sample_records) >= 2:
                break
            parsed = _heuristic_parse_row(row, card_type="old")
            if parsed is None:
                continue
            sample_records.append({
                "اسم رب الأسرة": parsed["name"],
                "رقم البطاقة": parsed["selected_card"],
                "الأفراد الكلية": parsed["total"],
                "الأفراد المستحقة": parsed["eligible"],
                "الأفراد المحجوبين": parsed["withheld"],
            })
        return {"detected": detected, "sample_records": sample_records, "merged_header": True}

    return None

def _split_merged_header_labels(cell_text):
    """يفكّك خلية عناوين مدمجة (كل أسماء الأعمدة بنص واحد مفصول بفراغات
    متعددة) لقائمة (دور، نص العنوان) — لغرض العرض فقط، مو لتحديد ترتيب
    الأعمدة الفعلي بالبيانات (غير موثوق لأن ترتيب النص لا يطابق بالضرورة
    ترتيب الأعمدة الحقيقي بجداول كهذي)."""
    tokens = [t.strip() for t in re.split(r"\s{2,}", cell_text) if t.strip()]
    labels = []
    for t in tokens:
        role = _smart_match_header_role(t)
        if role:
            labels.append((role, t))
    return labels

def _detect_merged_header_row(rows, max_scan=5):
    """يفحص أول عدة صفوف بحثاً عن صف فيه خلية واحدة فقط غير فارغة تحمل كل
    أسماء الأعمدة مدمجة سوية (بدل خلية منفصلة لكل عمود) — يرجع (رقم الصف،
    قائمة الأدوار المكتشفة) لو لقى الشكل المطلوب (اسم + بطاقة + عدد على
    الأقل)، وإلا (None, None)."""
    for r_idx in range(min(max_scan, len(rows))):
        non_empty = [c for c in rows[r_idx] if c and c.strip()]
        if len(non_empty) != 1:
            continue
        labels = _split_merged_header_labels(non_empty[0])
        roles_found = {role for role, _ in labels}
        has_name = "name" in roles_found
        has_card = any(k in roles_found for k in ("old_card", "new_card", "card_generic"))
        has_amount = any(k in roles_found for k in ("total", "eligible", "withheld"))
        if has_name and has_card and has_amount:
            return r_idx, labels
    return None, None

def extract_records_smart(file_obj, card_type="old"):
    records = {}
    duplicates = []
    tables_rows = _read_tables_rows(file_obj)

    last_role_map = None
    for rows in tables_rows:
        role_map, header_row_idx = _smart_detect_header_map(rows)
        if role_map is not None:
            last_role_map = role_map
        elif last_role_map is not None:
            # جدول/ورقة تكمل بيانات سابقة دون تكرار صف العناوين (مثل الورقة الثانية في نفس الملف)
            role_map = last_role_map
            header_row_idx = -1
        else:
            continue

        old_idx = role_map.get("old_card", role_map.get("card_generic"))
        new_idx = role_map.get("new_card", role_map.get("card_generic"))
        name_idx = role_map.get("name")
        total_idx = role_map.get("total")
        eligible_idx = role_map.get("eligible")
        withheld_idx = role_map.get("withheld")
        seq_idx = role_map.get("seq")

        for r_idx in range(header_row_idx + 1, len(rows)):
            cells = rows[r_idx]
            max_idx = len(cells) - 1
            if name_idx is None or name_idx > max_idx:
                continue
            name = cells[name_idx].strip()
            if not name or not any('؀' <= ch <= 'ۿ' for ch in name):
                continue
            if any(kw in name for kw in ("الإجمالي", "الاجمالي", "المجموع", "اجمالي", "إجمالي")):
                continue

            old_card = _smart_clean_card(cells[old_idx]) if old_idx is not None and old_idx <= max_idx else ""
            new_card = _smart_clean_card(cells[new_idx]) if new_idx is not None and new_idx <= max_idx else ""

            selected_card = (old_card or new_card) if card_type == "old" else (new_card or old_card)
            if not selected_card:
                continue
            alt_card = (new_card or old_card) if card_type == "old" else (old_card or new_card)
            if alt_card == selected_card:
                alt_card = ""

            total = _smart_to_int(cells[total_idx]) if total_idx is not None and total_idx <= max_idx else 0
            eligible = _smart_to_int(cells[eligible_idx]) if eligible_idx is not None and eligible_idx <= max_idx else 0
            withheld = _smart_to_int(cells[withheld_idx]) if withheld_idx is not None and withheld_idx <= max_idx else 0
            seq_val = cells[seq_idx].strip() if seq_idx is not None and seq_idx <= max_idx and cells[seq_idx].strip() else "-"

            if selected_card in records:
                duplicates.append((selected_card, name))
            else:
                records[selected_card] = {"seq": seq_val, "name": name, "total": total, "eligible": eligible, "withheld": withheld, "alt_card": alt_card}

    return records, duplicates

def merge_records_by_either_card(old_data, new_data):
    """يعتبر عائلتين متطابقتين إذا تطابق رقم البطاقة القديم بينهما أو
    الحديث (أيهما نجح) بدل الاعتماد على رقم واحد فقط لكل المقارنة — يستخدم
    الاثنين معاً كنقطة قوة للمطابقة. old_data/new_data مستخرجة بمفتاح
    أساسي واحد (البطاقة القديمة) وتحمل حقل alt_card للبطاقة الأخرى؛ تُرجع
    نسختين موحّدتين بنفس مفاتيح old_data لكل زوج تم العثور على تطابق له
    (بأي من الرقمين)، ليقدر محرك المقارنة الحالي يشتغل عليهما بدون تعديل."""
    new_by_alt = {}
    for k, rec in new_data.items():
        alt = rec.get("alt_card")
        if alt and alt not in new_data:
            new_by_alt.setdefault(alt, k)

    unified_old, unified_new = {}, {}
    used_new_keys = set()

    for old_key, old_rec in old_data.items():
        alt = old_rec.get("alt_card")
        target_key = None
        if old_key in new_data:
            target_key = old_key
        elif alt and alt in new_data:
            target_key = alt
        elif old_key in new_by_alt:
            target_key = new_by_alt[old_key]
        elif alt and alt in new_by_alt:
            target_key = new_by_alt[alt]

        unified_old[old_key] = old_rec
        if target_key and target_key not in used_new_keys:
            unified_new[old_key] = new_data[target_key]
            used_new_keys.add(target_key)

    for new_key, new_rec in new_data.items():
        if new_key not in used_new_keys:
            unified_new[new_key] = new_rec

    return unified_old, unified_new

def extract_matched_by_either_card(extract_fn, file_old, file_new):
    """يستخرج الملفين برقم البطاقة القديم كمفتاح أساسي (مع حفظ الحديث
    كبديل)، ثم يدمجهما بالاعتماد على أي الرقمين ينجح بالمطابقة. يرجع
    أيضاً قوائم البطاقات المكررة (لو وجدت) بكل ملف، للتنبيه بدل الإسقاط
    الصامت لعائلة كاملة."""
    file_old.seek(0); file_new.seek(0)
    old_data, old_duplicates = extract_fn(file_old, card_type="old")
    file_old.seek(0)
    new_data, new_duplicates = extract_fn(file_new, card_type="old")
    file_new.seek(0)
    unified_old, unified_new = merge_records_by_either_card(old_data, new_data)
    return unified_old, unified_new, "رقم البطاقة", old_duplicates, new_duplicates

def _extract_with_smart_fallback(file_obj, card_type="old"):
    """يجرب المحرك الذكي أولاً لهذا الملف تحديداً، ولا يلجأ للمحرك القديم
    إلا لنفس الملف لو فشل هوة تحديداً — بدل ما يهبط الملفين مع بعض
    للمحرك القديم لمجرد فشل واحد منهم بس. هذا يمنع سيناريو فعلي شوهد: ملف
    عناوينه مدمجة بخلية وحدة يفشل بالمحرك الذكي فيسحب معه الملف الثاني
    (اللي كان ناجح تماماً بالمحرك الذكي) للمحرك القديم، والمحرك القديم قد
    يفسّر ترتيب أعمدة الملف الثاني غلط ويسرّب رقم بطاقة لعمود عدد."""
    file_obj.seek(0)
    data, duplicates = extract_records_smart(file_obj, card_type=card_type)
    used_fallback = False
    if not data:
        file_obj.seek(0)
        data, duplicates = extract_clean_records(file_obj, card_type=card_type)
        used_fallback = True
    file_obj.seek(0)
    return data, duplicates, used_fallback

# -----------------------------------------------------------------------------
# 2.7. حارس سلامة البيانات: يرفض الاعتماد على سجلات فاسدة بدل تمريرها
# بصمت للمقارنة والتقارير (هذا بالضبط ما كان غايباً وسبب ظهور أرقام
# مستحيلة مثل ملايين "الأفراد المستحقين" لعشرات العوائل فقط، نتيجة تسرّب
# رقم بطاقة لعمود عدد).
# -----------------------------------------------------------------------------
MAX_REASONABLE_FAMILY_SIZE = 60  # سخي جداً مقارنة بأكبر عائلة شوهدت فعلياً (~13 فرد) لتفادي رفض حالات نادرة صحيحة

def validate_record(card, rec, skip_consistency_check=False):
    """يفحص سجل واحد مقابل قيود منطقية معروفة للنطاق، ويرجع نص الخطأ إذا
    فيه مشكلة أو None إذا سليم. لا يرمي استثناء أبداً — أي قيمة غير متوقعة
    الشكل تُعتبر بحد ذاتها خطأ يوصف ويُرجع، مو تُكسر تنفيذ البرنامج.
    skip_consistency_check=True لنموذج "المستحق فقط" اللي يُصفّر الكلي
    والمحجوب عمداً (extract_eligible_only_records) — فحص الاتساق بينهم
    وبين المستحق غير منطقي أصلاً بهذا النموذج ولازم يُستثنى، وإلا أي
    عائلة بأكثر من مستحقين اثنين ترفض غلط."""
    name = rec.get("name", "؟")
    total = rec.get("total")
    eligible = rec.get("eligible")
    withheld = rec.get("withheld")

    # 1) القيم لازم تكون أعداد صحيحة غير سالبة (مو نص أو None متسرب)
    if not all(isinstance(v, int) and not isinstance(v, bool) and v >= 0 for v in (total, eligible, withheld)):
        return f"'{name}' (بطاقة {card}): قيم أعداد غير صالحة (كلي={total!r}, مستحق={eligible!r}, محجوب={withheld!r})"

    # 2) حد أقصى منطقي لحجم العائلة — أي تجاوز يعني شبه مؤكد تسرّب رقم
    # بطاقة أو رقم تسلسل لعمود عدد، مو عائلة فعلية بهذا الحجم
    if total > MAX_REASONABLE_FAMILY_SIZE or eligible > MAX_REASONABLE_FAMILY_SIZE or withheld > MAX_REASONABLE_FAMILY_SIZE:
        return f"'{name}' (بطاقة {card}): عدد غير منطقي — كلي={total}, مستحق={eligible}, محجوب={withheld}"

    # 3) اتساق حسابي: المستحق + المحجوب ما يتجاوز الكلي (هامش تساهل بسيط
    # لفروقات تقريب/طباعة نادرة بالملفات المصدرية نفسها)
    if not skip_consistency_check and eligible + withheld > total + 2:
        return f"'{name}' (بطاقة {card}): المستحق({eligible}) + المحجوب({withheld}) أكبر من الكلي({total})"

    # 4) شكل رقم البطاقة نفسه — أرقام فقط وطول معقول
    if not (str(card).isdigit() and 3 <= len(str(card)) <= 10):
        return f"'{name}': رقم بطاقة مشبوه الشكل '{card}'"

    return None

def filter_valid_records(records, file_label, skip_consistency_check=False):
    """يفصل السجلات السليمة عن الفاسدة سجلاً بسجل — السجلات الفاسدة
    تُستبعد تماماً من أي حساب لاحق (لا تدخل المقارنة ولا التقارير أبداً)
    بدل ما تُترك تلوّث المجاميع، وتُرجع أوصافها لعرضها للمستخدم."""
    clean, errors = {}, []
    for card, rec in records.items():
        err = validate_record(card, rec, skip_consistency_check=skip_consistency_check)
        if err:
            errors.append(f"[{file_label}] {err}")
        else:
            clean[card] = rec
    return clean, errors

def validate_and_clean_pair(old_data, new_data, old_label, new_label, error_threshold_ratio=0.10, error_threshold_count=3, skip_consistency_check=False):
    """ينظّف الملفين من أي سجل فاسد، ويقرر: نسبة الفساد ضئيلة (يكمل
    بالسجلات النظيفة بعد تحذير) أو عالية (خلل منهجي بقراءة الأعمدة —
    يوقف كاملاً ولا يعتمد حتى على السجلات "السليمة" ظاهرياً، لأن الثقة
    بكل الملف تصير مهزوزة). يرجع (is_safe, clean_old, clean_new, errors)."""
    clean_old, errors_old = filter_valid_records(old_data, old_label, skip_consistency_check=skip_consistency_check)
    clean_new, errors_new = filter_valid_records(new_data, new_label, skip_consistency_check=skip_consistency_check)
    all_errors = errors_old + errors_new
    total_records = len(old_data) + len(new_data)

    if not all_errors:
        return True, clean_old, clean_new, []

    error_ratio = len(all_errors) / total_records if total_records else 1.0
    is_critical = len(all_errors) >= error_threshold_count and error_ratio >= error_threshold_ratio
    return (not is_critical), clean_old, clean_new, all_errors

# -----------------------------------------------------------------------------
# 3. محرك المقارنة الذكي الثابت (محدث لدعم النموذج الرابع)
# -----------------------------------------------------------------------------
def process_comparison(old_data, new_data, mode, card_col_name, matching_engine):
    results = []
    results_type_1_reference = []
    counters = {
        "total_fam": 0, "eligible_fam": 0, "withheld_fam": 0, "added_fam": 0, "deleted_fam": 0,
        "inc_total": 0, "dec_total": 0, "net_total": 0, "inc_eligible": 0, "dec_eligible": 0, "net_eligible": 0,
        "inc_withheld": 0, "dec_withheld": 0, "net_withheld": 0
    }
    
    skip_seq_matching = (matching_engine == "محرك تخطي التسلسل (بطاقة فقط)")
    all_cards = set(old_data.keys()).union(set(new_data.keys()))
    
    for card in all_cards:
        if card in old_data and card in new_data:
            old_v, new_v = old_data[card], new_data[card]
            d_tot = new_v["total"] - old_v["total"]
            d_elig = new_v["eligible"] - old_v["eligible"]
            d_with = new_v["withheld"] - old_v["withheld"]
            
            target_seq = new_v["seq"] if skip_seq_matching else old_v["seq"]
            notes = []
            
            # شروط النموذج الرابع (المستحق فقط)
            if mode == "النموذج الرابع (المستحق فقط)":
                is_changed = (d_elig != 0)
                if d_elig > 0: notes.append(f"زيادة مستحق ({d_elig})")
                elif d_elig < 0: notes.append(f"نقصان مستحق ({abs(d_elig)})")
                referral_text = " | ".join(notes) if notes else ""
            else:
                is_changed = d_tot != 0 or d_elig != 0 or d_with != 0
                if old_v["name"] != new_v["name"]:
                    notes.append(f"تم تغيير الاسم / السابق / {old_v['name']}")
                    is_changed = True
                if new_v["withheld"] == new_v["total"] and new_v["total"] > 0 and d_with > 0:
                    notes.append("حجب كلي ❌")
                else:
                    if d_with > 0: notes.append(f"تم حجب {d_with} نفر")
                    elif d_with < 0: notes.append(f"تم رفع الحجب عن {abs(d_with)} نفر")
                if d_tot > 0: notes.append("إضافة طفل 👶")
                elif d_tot < 0: notes.append(f"نقصان {abs(d_tot)} نفر")
                referral_text = " | ".join(notes) if notes else ("تحديث بيانات" if is_changed else "")
            
            if is_changed:
                if d_tot != 0:
                    counters["total_fam"] += 1
                    counters["net_total"] += d_tot
                    if d_tot > 0: counters["inc_total"] += d_tot
                    else: counters["dec_total"] += abs(d_tot)
                if d_elig != 0:
                    counters["eligible_fam"] += 1
                    counters["net_eligible"] += d_elig
                    if d_elig > 0: counters["inc_eligible"] += d_elig
                    else: counters["dec_eligible"] += abs(d_elig)
                if d_with != 0:
                    counters["withheld_fam"] += 1
                    counters["net_withheld"] += d_with
                    if d_with > 0: counters["inc_withheld"] += d_with
                    else: counters["dec_withheld"] += abs(d_with)
                
                total_val = "-" if mode == "النموذج الرابع (المستحق فقط)" else new_v["total"]
                withheld_val = "-" if mode == "النموذج الرابع (المستحق فقط)" else new_v["withheld"]

                base_dict = {
                    "التسلسل": target_seq, "اسم رب الأسرة": new_v["name"], card_col_name: card,
                    "الأفراد الكلية": total_val, "الأفراد المستحقة": new_v["eligible"], 
                    "الأفراد المحجوبين": withheld_val, "الإحالة": referral_text, "meta_card": card
                }
                
                results_type_1_reference.append({**base_dict, "meta_status": "modified"})
                
                if mode in ["النوع الأول", "النوع الثالث", "النموذج الرابع (المستحق فقط)", "النموذج الخامس (كشف تلقائي بالعناوين)"]:
                    results.append({**base_dict, "meta_status": "modified", "meta_sort": 1})
                elif mode == "النوع الثاني":
                    results.append({
                        "التسلسل": target_seq, "اسم رب الأسرة": old_v["name"], card_col_name: card, "الحالة": "السابق",
                        "الأفراد الكلية": "-" if mode == "النموذج الرابع (المستحق فقط)" else old_v["total"], 
                        "الأفراد المستحقة": old_v["eligible"], 
                        "الأفراد المحجوبين": "-" if mode == "النموذج الرابع (المستحق فقط)" else old_v["withheld"],
                        "الإحالة": "", "meta_status": "type2_old", "meta_card": card, "meta_sort": 1
                    })
                    results.append({
                        "التسلسل": target_seq, "اسم رب الأسرة": new_v["name"], card_col_name: card, "الحالة": "الحديث",
                        "الأفراد الكلية": total_val, "الأفراد المستحقة": new_v["eligible"], 
                        "الأفراد المحجوبين": withheld_val,
                        "الإحالة": referral_text, "meta_status": "type2_new", "meta_card": card, "meta_sort": 2
                    })
            elif mode == "النوع الثالث":
                results.append({
                    "التسلسل": target_seq, "اسم رب الأسرة": old_v["name"], card_col_name: card,
                    "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], 
                    "الأفراد المحجوبين": new_v["withheld"], "الإحالة": "", "meta_status": "normal", "meta_card": card, "meta_sort": 1
                })
                
        elif card in old_data and card not in new_data:
            old_v = old_data[card]
            counters["deleted_fam"] += 1
            counters["dec_total"] += old_v["total"]
            counters["net_total"] -= old_v["total"]
            counters["dec_eligible"] += old_v["eligible"]
            counters["net_eligible"] -= old_v["eligible"]
            counters["dec_withheld"] += old_v["withheld"]
            counters["net_withheld"] -= old_v["withheld"]
            
            total_val = "-" if mode == "النموذج الرابع (المستحق فقط)" else old_v["total"]
            withheld_val = "-" if mode == "النموذج الرابع (المستحق فقط)" else old_v["withheld"]

            base_row = {"التسلسل": old_v["seq"], "اسم رب الأسرة": old_v["name"], card_col_name: card,
                        "الأفراد الكلية": total_val, "الأفراد المستحقة": old_v["eligible"], 
                        "الأفراد المحجوبين": withheld_val, "الإحالة": "عائلة منقولة ❌", "meta_card": card}
            results_type_1_reference.append({**base_row, "meta_status": "deleted"})
            if mode == "النوع الثاني": results.append({**base_row, "الحالة": "منقول", "meta_status": "deleted", "meta_card": card, "meta_sort": 1})
            else: results.append({**base_row, "meta_status": "deleted", "meta_sort": 1})
                
        elif card not in old_data and card in new_data:
            new_v = new_data[card]
            counters["added_fam"] += 1
            counters["inc_total"] += new_v["total"]
            counters["net_total"] += new_v["total"]
            counters["inc_eligible"] += new_v["eligible"]
            counters["net_eligible"] += new_v["eligible"]
            counters["inc_withheld"] += new_v["withheld"]
            counters["net_withheld"] += new_v["withheld"]
            
            total_val = "-" if mode == "النموذج الرابع (المستحق فقط)" else new_v["total"]
            withheld_val = "-" if mode == "النموذج الرابع (المستحق فقط)" else new_v["withheld"]

            base_row = {"التسلسل": new_v["seq"], "اسم رب الأسرة": new_v["name"], card_col_name: card,
                        "الأفراد الكلية": total_val, "الأفراد المستحقة": new_v["eligible"], 
                        "الأفراد المحجوبين": withheld_val, "الإحالة": "عائلة مضافة ✨", "meta_card": card}
            results_type_1_reference.append({**base_row, "meta_status": "added"})
            if mode == "النوع الثاني": results.append({**base_row, "الحالة": "مضاف", "meta_status": "added", "meta_card": card, "meta_sort": 1})
            else: results.append({**base_row, "meta_status": "added", "meta_sort": 1})
                
    return results, results_type_1_reference, counters

# -----------------------------------------------------------------------------
# 4. دوال التصدير المحدثة والمصقولة لملف الوورد الناتج
# -----------------------------------------------------------------------------
def set_cell_background(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_width(cell, width_inches):
    cell.width = Inches(width_inches)
    tcPr = cell._tc.get_or_add_tcPr()
    dxa_val = int(width_inches * 1440)
    tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{dxa_val}" w:type="dxa"/>')
    tcPr.append(tcW)

def clean_to_triple_name(name_str):
    if not name_str or pd.isna(name_str): return ""
    words = str(name_str).strip().split()
    return " ".join(words[:3])

def format_run(run, font_name="Microsoft Sans Serif", size_pt=14, color_rgb=None, bold=False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color_rgb: run.font.color.rgb = color_rgb
    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>')
    rPr.append(rFonts)

def add_page_number(paragraph):
    p = paragraph._p
    run_text = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = "الصفحة "
    run_text.append(t)
    p.append(run_text)

    run_fld = OxmlElement('w:r')
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    run_fld.append(fldChar1)
    run_fld.append(instrText)
    run_fld.append(fldChar2)
    run_fld.append(fldChar3)
    p.append(run_fld)

def create_word_table_report(doc_df, title, mode, card_col_name, old_data, new_data, new_file_name):
    doc = Document()
    
    for section in doc.sections:
        section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
        w, h = section.page_height, section.page_width
        section.page_width = w
        section.page_height = h
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        
        footer = section.footer
        p_footer = footer.paragraphs[0]
        p_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_page_number(p_footer)
        
    def append_table_to_doc(target_doc, df_to_write, table_title):
        agent_name = new_file_name.replace(".docx", "").replace(".xlsx", "")
        agent_name = re.sub(r'(FOOD|FLOUR)', '', agent_name, flags=re.IGNORECASE)
        agent_name = agent_name.strip("- ").strip()
        
        banner_table = target_doc.add_table(rows=1, cols=1)
        banner_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        banner_cell = banner_table.rows[0].cells[0]
        
        set_cell_background(banner_cell, "111E38")
        
        tcPr = banner_cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="180" w:type="dxa"/><w:bottom w:w="180" w:type="dxa"/><w:left w:w="250" w:type="dxa"/><w:right w:w="250" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)
        p_banner = banner_cell.paragraphs[0]
        p_banner.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        run_b1 = p_banner.add_run(f"تقرير متغيرات الوكيل: {agent_name}")
        format_run(run_b1, font_name="Microsoft Sans Serif", size_pt=16, color_rgb=RGBColor(255, 255, 255), bold=True)
        
        if "FOOD" in new_file_name.upper():
            run_b2 = p_banner.add_run(" (غذائية)")
            format_run(run_b2, font_name="Microsoft Sans Serif", size_pt=16, color_rgb=RGBColor(255, 243, 79), bold=True)
        elif "FLOUR" in new_file_name.upper():
            run_b2 = p_banner.add_run(" (طحين)")
            format_run(run_b2, font_name="Microsoft Sans Serif", size_pt=16, color_rgb=RGBColor(240, 254, 240), bold=True)
            
        p_sub = target_doc.add_paragraph()
        p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_sub = p_sub.add_run(f"\n{table_title}")
        format_run(run_sub, font_name="Microsoft Sans Serif", size_pt=13, color_rgb=RGBColor(44, 62, 80), bold=True)
        
        display_df = df_to_write.drop(columns=["meta_status", "meta_card", "meta_sort"], errors="ignore")
        cols = list(display_df.columns) 
        
        table = target_doc.add_table(rows=1, cols=len(cols))
        table.style = 'Table Grid'
        table.autofit = False 
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        tblPr = table._element.tblPr
        tblLayout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
        tblPr.append(tblLayout)
        
        bidiVisual = parse_xml(f'<w:bidiVisual {nsdecls("w")}/>')
        tblPr.append(bidiVisual)
        
        width_map = {
            "التسلسل": 0.55,
            "اسم رب الأسرة": 3.0,
            card_col_name: 0.90,
            "الحالة": 0.6,
            "الأفراد الكلية": 0.45,
            "الأفراد المستحقة": 0.45,
            "الأفراد المحجوبين": 0.45,
            "الإحالة": 4.0
        }
        
        hdr_cells = table.rows[0].cells
        for i, col in enumerate(cols):
            display_name = col
            if col == "اسم رب الأسرة": display_name = "اسم المواطن"
            elif col == "الإحالة": display_name = "الحالة"
            elif col == card_col_name: display_name = "رقم البطاقة"
            elif col == "التسلسل": display_name = "ت"
            elif col == "الأفراد الكلية": display_name = "الكلي"
            elif col == "الأفراد المستحقة": display_name = "المستحق"
            elif col == "الأفراد المحجوبين": display_name = "المحجوب"
            
            hdr_cells[i].text = display_name
            set_cell_width(hdr_cells[i], width_map.get(col, 1.0))
            set_cell_background(hdr_cells[i], "E8ECEF")
            
            if col in ["الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين"]:
                tcPr = hdr_cells[i]._tc.get_or_add_tcPr()
                textDirection = parse_xml(f'<w:textDirection {nsdecls("w")} w:val="btLr"/>')
                tcPr.append(textDirection)
            
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if p.runs:
                format_run(p.runs[0], font_name="Microsoft Sans Serif", size_pt=14, color_rgb=RGBColor(44, 62, 80), bold=True)
        
        prev_cells = None
        for row_idx, (_, row) in enumerate(df_to_write.iterrows()):
            row_cells = table.add_row().cells
            status = row.get("meta_status", "normal")
            
            if mode == "النوع الثاني":
                bg_color = "FFFFFF" if (row_idx // 2) % 2 == 0 else "F2F4F4"
            else:
                bg_color = "FFFFFF" if row_idx % 2 == 0 else "F2F4F4"
            
            for i, col in enumerate(cols):
                set_cell_width(row_cells[i], width_map.get(col, 1.0))
                set_cell_background(row_cells[i], bg_color)
            
            for i, col in enumerate(cols):
                cell = row_cells[i]
                val_text = str(row[col]) if pd.notna(row[col]) and str(row[col]) != "" else ""
                if col == "اسم رب الأسرة": val_text = clean_to_triple_name(val_text)
                
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                if col == "الإحالة" and val_text:
                    parts = val_text.split(" | ")
                    for p_idx, part in enumerate(parts):
                        run = p.add_run(part)
                        part_color = RGBColor(0, 0, 0)
                        if "طفل" in part: part_color = RGBColor(0, 0, 255)
                        elif "رفع" in part or "زيادة مستحق" in part: part_color = RGBColor(0, 128, 0)
                        elif "حجب كلي" in part: part_color = RGBColor(128, 0, 0)
                        elif "حجب" in part or "نقصان مستحق" in part: part_color = RGBColor(255, 0, 0)
                        elif "مضافة" in part: part_color = RGBColor(0, 128, 0)
                        elif "منقولة" in part: part_color = RGBColor(255, 0, 0)
                        
                        format_run(run, font_name="Calibri", size_pt=14, color_rgb=part_color, bold=True)
                        if p_idx < len(parts) - 1:
                            sep_run = p.add_run(" | ")
                            format_run(sep_run, font_name="Calibri", size_pt=14, color_rgb=RGBColor(0,0,0), bold=True)
                else:
                    run = p.add_run(val_text)
                    c_font, c_size, c_color, c_bold = "Microsoft Sans Serif", 14, RGBColor(0, 0, 0), False
                    
                    if col == "اسم رب الأسرة":
                        c_size = 16
                        c_bold = True
                    elif col == "الحالة":
                        c_font = "Calibri"
                        c_color, c_bold = RGBColor(102, 0, 153), True
                    elif col == "الأفراد الكلية": c_color, c_bold = RGBColor(0, 51, 204), True
                    elif col == "الأفراد المستحقة": c_color, c_bold = RGBColor(0, 128, 0), True
                    elif col == "الأفراد المحجوبين": c_color, c_bold = RGBColor(204, 0, 0), True
                        
                    format_run(run, font_name=c_font, size_pt=c_size, color_rgb=c_color, bold=c_bold)
            
            if mode == "النوع الثاني":
                if status == "type2_old": prev_cells = row_cells
                elif status == "type2_new" and prev_cells:
                    for merge_col in ["التسلسل", "اسم رب الأسرة", card_col_name, "الإحالة"]:
                        if merge_col in cols:
                            m_idx = cols.index(merge_col)
                            if merge_col == "الإحالة": text_to_keep = str(row["الإحالة"])
                            elif merge_col == "اسم رب الأسرة": text_to_keep = clean_to_triple_name(row["اسم رب الأسرة"])
                            elif merge_col == card_col_name: text_to_keep = str(row[card_col_name])
                            else: text_to_keep = prev_cells[m_idx].text
                                
                            prev_cells[m_idx].merge(row_cells[m_idx])
                            set_cell_width(prev_cells[m_idx], width_map.get(merge_col, 1.0))
                            set_cell_background(prev_cells[m_idx], bg_color) 
                            
                            prev_cells[m_idx].text = ""
                            p_merge = prev_cells[m_idx].paragraphs[0]
                            p_merge.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            
                            if merge_col == "الإحالة" and text_to_keep:
                                parts = text_to_keep.split(" | ")
                                for p_idx, part in enumerate(parts):
                                    run = p_merge.add_run(part)
                                    part_color = RGBColor(0, 0, 0)
                                    if "طفل" in part: part_color = RGBColor(0, 0, 255)
                                    elif "رفع" in part: part_color = RGBColor(0, 128, 0)
                                    elif "حجب كلي" in part: part_color = RGBColor(128, 0, 0)
                                    elif "حجب" in part: part_color = RGBColor(255, 0, 0)
                                    
                                    format_run(run, font_name="Calibri", size_pt=14, color_rgb=part_color, bold=True)
                                    if p_idx < len(parts) - 1:
                                        sep_run = p_merge.add_run(" | ")
                                        format_run(sep_run, font_name="Calibri", size_pt=14, color_rgb=RGBColor(0,0,0), bold=True)
                            else:
                                run = p_merge.add_run(text_to_keep)
                                c_font, c_size, c_bold = "Microsoft Sans Serif", 14, False
                                if merge_col == "اسم رب الأسرة": c_size, c_bold = 16, True
                                format_run(run, font_name=c_font, size_pt=c_size, color_rgb=RGBColor(0,0,0), bold=c_bold)

    append_table_to_doc(doc, doc_df, title)
    
    cases_to_extract = [
        ("حالات تغيير اسم رب الأسرة", "تم تغيير الاسم"), ("حالات إضافة طفل", "إضافة طفل"),
        ("حالات حجب كلي", "حجب كلي"), ("حالات حجب نفر", "تم حجب"), ("حالات رفع الحجب", "تم رفع الحجب"),
        ("حالات زيادة مستحق", "زيادة مستحق"), ("حالات نقصان مستحق", "نقصان مستحق"),
        ("العوائل المضافة", "عائلة مضافة"), ("العوائل المنقولة", "عائلة منقولة")
    ]
    for case_title, keyword in cases_to_extract:
        if keyword == "تم حجب": matched_mask = doc_df['الإحالة'].str.contains("تم حجب", na=False) & ~doc_df['الإحالة'].str.contains("حجب كلي", na=False)
        else: matched_mask = doc_df['الإحالة'].str.contains(keyword, na=False)
            
        if mode == "النوع الثاني":
            matched_cards = doc_df[matched_mask]['meta_card'].dropna().unique()
            case_df = doc_df[doc_df['meta_card'].isin(matched_cards)]
        else: case_df = doc_df[matched_mask]
            
        if not case_df.empty:
            doc.add_page_break()
            append_table_to_doc(doc, case_df, f"كشف مستقل: {case_title}")
                
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# 5. التنسيق البصري للويب 
# -----------------------------------------------------------------------------
def style_all_types(doc_df, old_data, new_data, card_col_name, mode):
    styles = pd.DataFrame('', index=doc_df.index, columns=doc_df.columns)
    for idx, row in doc_df.iterrows():
        status = row.get("meta_status", "")
        card = row.get("meta_card")
        notes = str(row.get("الإحالة", ""))
        
        if "تم تغيير الاسم" in notes: styles.loc[idx, "الإحالة"] = 'color: #2980B9; font-weight: bold;'
        elif "إضافة طفل" in notes or "زيادة مستحق" in notes: styles.loc[idx, "الإحالة"] = 'color: #1ABC9C; font-weight: bold;'
        elif "حجب كلي" in notes or "تم حجب" in notes or "نقصان مستحق" in notes: styles.loc[idx, "الإحالة"] = 'color: #C0392B; font-weight: bold;'
        elif "تم رفع الحجب" in notes or "مضافة" in notes: styles.loc[idx, "الإحالة"] = 'color: #27AE60; font-weight: bold;'
        
        if status == "type2_old": styles.loc[idx, "الحالة"] = 'background-color: #F5F5F5; font-weight: bold; color: #7F8C8D;'
        elif status == "type2_new": styles.loc[idx, "الحالة"] = 'background-color: #E8F8F5; font-weight: bold; color: #16A085;'
        elif status == "added": styles.loc[idx] = 'background-color: #E8F5E9; color: #2E7D32;'
        elif status == "deleted": styles.loc[idx] = 'background-color: #ECEFF1; color: #455A64; text-decoration: line-through;'

        if status in ["modified", "type2_old", "type2_new"] and card in old_data and card in new_data:
            o_val, n_val = old_data[card], new_data[card]
            if mode != "النموذج الرابع (المستحق فقط)" and o_val["total"] != n_val["total"]: styles.loc[idx, "الأفراد الكلية"] = 'background-color: #FDE0DC; font-weight: bold; color: #C0392B;'
            if o_val["eligible"] != n_val["eligible"]: styles.loc[idx, "الأفراد المستحقة"] = 'background-color: #FDE0DC; font-weight: bold; color: #C0392B;'
            if mode != "النموذج الرابع (المستحق فقط)" and o_val["withheld"] != n_val["withheld"]: styles.loc[idx, "الأفراد المحجوبين"] = 'background-color: #FDE0DC; font-weight: bold; color: #C0392B;'
    return styles

def create_word_stats_report(counters, filename_base):
    doc = Document()
    doc.add_heading(f"تقرير الإحصاء - للملف: {filename_base}", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph().add_run("أولاً: إحصاء حركة الأفراد").bold = True
    stats_individuals = [
        ("زيادة الكلية:", f"+{counters['inc_total']}"), ("نقصان الكلية:", f"-{counters['dec_total']}"), ("صافي الكلية:", f"{counters['net_total']:+d}"),
        ("زيادة المستحقة:", f"+{counters['inc_eligible']}"), ("نقصان المستحقة:", f"-{counters['dec_eligible']}"), ("صافي المستحقة:", f"{counters['net_eligible']:+d}"),
        ("زيادة المحجوبين:", f"+{counters['inc_withheld']}"), ("نقصان المحجوبين:", f"-{counters['dec_withheld']}"), ("صافي المحجوبين:", f"{counters['net_withheld']:+d}")
    ]
    for text, val in stats_individuals: doc.add_paragraph().add_run(f"{val} : {text}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph().add_run("ثانياً: العوائل").bold = True
    stats_families = [
        ("تغيرت الكلية:", counters['total_fam']), ("تغيرت المستحقة:", counters['eligible_fam']),
        ("تغيرت المحجوبين:", counters['withheld_fam']), ("عوائل مضافة:", counters['added_fam']), ("عوائل منقولة:", counters['deleted_fam'])
    ]
    for text, val in stats_families: doc.add_paragraph().add_run(f"{val} : {text}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# 5.5. تقارير PDF أنيقة منفصلة لكل حالة من حالات المتغيرات (تصميم A4 قابل للطباعة)
# -----------------------------------------------------------------------------
# كل عنصر: مفتاح الحالة، عنوان التقرير، وصف فرعي، كلمة الشارة، الأيقونة، ألوان
# التمييز، وهل يُعرض عمود "الإحالة" (سبب التغيير) بالجدول، ودالة تحدد هل الصف
# ينتمي لهذه الحالة (صف واحد قد ينتمي لأكثر من حالة، مثل زيادة أفراد + حجب معاً)
CATEGORY_DEFS = [
    {"key": "added", "title": "تقرير العوائل المضافة",
     "subtitle": "العوائل الجديدة التي ظهرت في كشف الوكيل {agent} الحالي",
     "badge_label": "مضافة", "icon": "🆕", "accent": "#1E8449", "accent_soft": "#EAFAF1", "accent_dark": "#145A32",
     "show_referral": False, "match": lambda r: r.get("meta_status") == "added"},
    {"key": "deleted", "title": "تقرير العوائل المنقولة",
     "subtitle": "العوائل الموجودة سابقاً والمنقولة من كشف الوكيل {agent} الحالي",
     "badge_label": "منقولة", "icon": "📤", "accent": "#C0392B", "accent_soft": "#FDEDEC", "accent_dark": "#922B21",
     "show_referral": False, "match": lambda r: r.get("meta_status") == "deleted"},
    {"key": "full_block", "title": "تقرير الحجب الكلي",
     "subtitle": "عوائل تم حجب كامل أفرادها في كشف الوكيل {agent} الحالي",
     "badge_label": "حجب كلي", "icon": "⛔", "accent": "#7B241C", "accent_soft": "#FDEDEC", "accent_dark": "#5B1A12",
     "show_referral": True, "match": lambda r: "حجب كلي" in str(r.get("الإحالة") or "")},
    {"key": "block_up", "title": "تقرير زيادة الحجب",
     "subtitle": "عوائل ارتفع فيها عدد الأفراد المحجوبين في كشف الوكيل {agent} الحالي",
     "badge_label": "زيادة حجب", "icon": "🔒", "accent": "#B9770E", "accent_soft": "#FEF9E7", "accent_dark": "#7D6608",
     "show_referral": True, "match": lambda r: "تم حجب" in str(r.get("الإحالة") or "") and "حجب كلي" not in str(r.get("الإحالة") or "")},
    {"key": "block_down", "title": "تقرير رفع الحجب",
     "subtitle": "عوائل رُفع عنها الحجب جزئياً أو كلياً في كشف الوكيل {agent} الحالي",
     "badge_label": "رفع حجب", "icon": "🔓", "accent": "#117864", "accent_soft": "#E8F8F5", "accent_dark": "#0B5345",
     "show_referral": True, "match": lambda r: "رفع الحجب" in str(r.get("الإحالة") or "")},
    {"key": "members_up", "title": "تقرير زيادة عدد الأفراد",
     "subtitle": "عوائل زاد فيها عدد الأفراد الكلي (إضافة مولود) في كشف الوكيل {agent} الحالي",
     "badge_label": "زيادة أفراد", "icon": "👶", "accent": "#1F618D", "accent_soft": "#EBF5FB", "accent_dark": "#154360",
     "show_referral": True, "match": lambda r: "إضافة طفل" in str(r.get("الإحالة") or "")},
    {"key": "members_down", "title": "تقرير نقصان عدد الأفراد",
     "subtitle": "عوائل نقص فيها عدد الأفراد الكلي في كشف الوكيل {agent} الحالي",
     "badge_label": "نقصان أفراد", "icon": "📉", "accent": "#A04000", "accent_soft": "#FDF2E9", "accent_dark": "#6E2C00",
     "show_referral": True, "match": lambda r: bool(re.search(r'نقصان\s+\d+\s+نفر', str(r.get("الإحالة") or "")))},
    {"key": "eligible_up", "title": "تقرير زيادة المستحقين",
     "subtitle": "عوائل زاد فيها عدد الأفراد المستحقين في كشف الوكيل {agent} الحالي",
     "badge_label": "زيادة مستحق", "icon": "📈", "accent": "#1F618D", "accent_soft": "#EBF5FB", "accent_dark": "#154360",
     "show_referral": True, "match": lambda r: "زيادة مستحق" in str(r.get("الإحالة") or "")},
    {"key": "eligible_down", "title": "تقرير نقصان المستحقين",
     "subtitle": "عوائل نقص فيها عدد الأفراد المستحقين في كشف الوكيل {agent} الحالي",
     "badge_label": "نقصان مستحق", "icon": "📉", "accent": "#B9770E", "accent_soft": "#FEF9E7", "accent_dark": "#7D6608",
     "show_referral": True, "match": lambda r: "نقصان مستحق" in str(r.get("الإحالة") or "")},
    {"key": "name_change", "title": "تقرير تغيير الأسماء",
     "subtitle": "عوائل تغيّر اسم رب الأسرة فيها بين الملفين في كشف الوكيل {agent} الحالي",
     "badge_label": "تغيير اسم", "icon": "✎", "accent": "#6C3483", "accent_soft": "#F4ECF7", "accent_dark": "#4A235A",
     "show_referral": True, "match": lambda r: "تغيير الاسم" in str(r.get("الإحالة") or "")},
    {"key": "generic_update", "title": "تقرير تحديثات أخرى",
     "subtitle": "عوائل طرأ عليها تحديث غير مصنف ضمن الحالات أعلاه في كشف الوكيل {agent} الحالي",
     "badge_label": "تحديث عام", "icon": "↻", "accent": "#566573", "accent_soft": "#F4F6F6", "accent_dark": "#2C3E50",
     "show_referral": True, "match": lambda r: str(r.get("الإحالة") or "").strip() == "تحديث بيانات"},
]

def _hex_to_rgb(hex_color):
    h = hex_color.lstrip('#')
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))

def _rgba(hex_color, alpha):
    r, g, b = _hex_to_rgb(hex_color)
    return f"rgba({r},{g},{b},{alpha})"

_FONT_LINK = """<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">"""

# تصميم زجاجي (Glassmorphism) موحّد لكل تقارير الـPDF: خط Cairo (من أشهر
# وأجمل الخطوط العربية الرسمية المستخدمة بالمواقع الكبيرة)، عناوين وأرقام
# أكبر وأوضح، ألوان نص دائماً واضحة مبنية على لون كل حالة (بدل الرصاصي
# الخافت)، وجدول بأعمدة عرضها ثابت بالنسبة المئوية + قفل كامل لالتفاف
# النص (nowrap + ellipsis) بحيث لا يتجاوز الجدول عرض الورقة مهما كبر الخط.
_PDF_CSS = """
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; font-family: 'Cairo', 'Tajawal', 'Arial', sans-serif; color: #1B2631; background: #fff; }
  @page { size: A4; margin: 12mm 8mm 14mm 8mm; }
  .report-section.with-break { page-break-before: always; }
  .header { text-align: center; padding-bottom: 16px; margin-bottom: 22px; border-bottom: 3px solid var(--accent); }
  .icon-badge { display: inline-block; width: 64px; height: 64px; line-height: 64px; text-align: center; margin-bottom: 10px; border-radius: 24px; background: var(--accent); color: #fff; font-size: 30px; font-weight: 900; box-shadow: 0 4px 14px var(--glass-shadow); }
  .header h1 { margin: 4px 0 12px; font-size: 28px; font-weight: 800; color: var(--accent-dark); }
  .pills { display: flex; justify-content: center; align-items: center; gap: 10px; flex-wrap: wrap; }
  .desc-pill { display: inline-block; background: var(--glass-bg); border: 1px solid var(--glass-border); border-radius: 999px; padding: 10px 28px; font-size: 13px; color: var(--accent-dark); font-weight: 600; box-shadow: 0 2px 10px var(--glass-shadow); }
  .agent-pill { display: inline-block; background: var(--accent-dark); color: #fff; border-radius: 999px; padding: 8px 24px; font-size: 12.5px; font-weight: 700; box-shadow: 0 2px 10px var(--glass-shadow); white-space: nowrap; }
  .stats { display: flex; gap: 12px; margin: 22px 0; }
  .stat-card { flex: 1; text-align: center; padding: 16px 8px; border-radius: 20px; background: var(--glass-bg); border: 1px solid var(--glass-border); box-shadow: 0 3px 12px var(--glass-shadow); }
  .stat-card .num { font-size: 32px; font-weight: 900; color: var(--accent-dark); display: block; line-height: 1.25; }
  .stat-card .lbl { font-size: 12px; color: var(--accent-dark); font-weight: 600; opacity: 0.85; }
  .table-wrap { border-radius: 22px; overflow: hidden; border: 1px solid var(--glass-border); box-shadow: 0 3px 14px var(--glass-shadow); }
  table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 13.5px; }
  thead th { background: var(--accent); color: #fff; font-weight: 700; padding: 11px 4px; text-align: center; font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  tbody td { padding: 9px 5px; text-align: center; border-bottom: 1px solid var(--line-color); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  tbody tr:nth-child(even) { background: #ffffff; }
  tbody tr:nth-child(odd) { background: var(--accent-soft); }
  tbody tr { page-break-inside: avoid; }
  .c-idx { color: var(--accent-dark); font-weight: 700; }
  .c-name { text-align: right; font-weight: 700; color: #1B2631; font-size: 12px; white-space: normal; overflow: visible; text-overflow: clip; }
  .c-name .name-text { line-height: 1.3; }
  .status-pill { display: inline-block; margin-top: 4px; padding: 2px 11px; border-radius: 999px; background: var(--pill-bg); border: 1px solid var(--pill-border); color: var(--accent-dark); font-size: 9px; font-weight: 600; line-height: 1.6; white-space: normal; }
  .c-mono { font-family: 'Consolas', monospace; direction: ltr; color: var(--accent-dark); font-weight: 600; }
  .c-num { font-weight: 800; color: #1B2631; }
  .c-eligible { color: #196F3D; }
  .c-withheld { color: #A93226; }
  .c-blank { background: #fff !important; }
  .th-blank { white-space: normal !important; font-size: 8.5px !important; line-height: 1.2; }
  .blank-chip { display: block; margin: 3px auto 0; width: 65%; height: 18px; border: 1.2px solid #5D6D7E; border-radius: 6px; background: #fff; }
  .footer { margin-top: 18px; padding-top: 10px; border-top: 1px solid var(--line-color); display: flex; justify-content: space-between; font-size: 11px; color: var(--accent-dark); font-weight: 600; }
  .cover { text-align: center; padding-top: 55px; }
  .cover h1 { font-size: 33px; color: var(--accent-dark); margin-bottom: 16px; }
  .cover .summary-grid { display: flex; flex-wrap: wrap; gap: 12px; justify-content: center; margin-top: 32px; }
  .cover .summary-card { width: 145px; padding: 16px 8px; border-radius: 18px; background: var(--glass-bg); border: 1px solid var(--glass-border); box-shadow: 0 3px 12px var(--glass-shadow); }
  .cover .summary-card .num { display: block; font-size: 28px; font-weight: 900; color: var(--accent-dark); }
  .cover .summary-card .lbl { font-size: 11.5px; color: var(--accent-dark); font-weight: 600; }
"""

def _wrap_pdf_document(title, body_html, css=None):
    return f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{title}</title>
{_FONT_LINK}
<style>{css or _PDF_CSS}</style>
</head>
<body>
{body_html}
</body>
</html>"""

# -----------------------------------------------------------------------------
# نموذج "كانفا": مستوحى من تصميم بطاقات المقارنة الشائعة (خلفية شبكية خفيفة،
# عنوان كبير عريض، شارة بيضوية عائمة فوق الجدول، وخلفية جدول شفافة بنسبة 5%
# من لون الحالة). يستخدم نفس ألوان كل حالة (CATEGORY_DEFS) بدون أي تغيير.
# -----------------------------------------------------------------------------
_PDF_CSS_CANVA = """
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; font-family: 'Cairo', 'Tajawal', 'Arial', sans-serif; color: #1B2631; background: #fff; }
  @page { size: A4; margin: 14mm 10mm 14mm 10mm; }
  body {
    background-image:
      linear-gradient(var(--grid-line) 1px, transparent 1px),
      linear-gradient(90deg, var(--grid-line) 1px, transparent 1px);
    background-size: 22px 22px;
  }
  .cv-section.with-break { page-break-before: always; }
  .cv-header { text-align: right; padding-bottom: 14px; margin-bottom: 26px; }
  .cv-header h1 { margin: 0 0 6px 0; font-size: 32px; font-weight: 800; color: var(--accent-dark); }
  .cv-header p { margin: 0; font-size: 13px; color: #5D6D7E; font-weight: 600; }
  .cv-header .cv-agent { display: inline-block; margin-top: 10px; background: #fff; border: 1.5px solid var(--accent); color: var(--accent-dark); border-radius: 999px; padding: 5px 18px; font-size: 12px; font-weight: 700; }

  .cv-stats { display: flex; gap: 10px; margin-bottom: 18px; }
  .cv-stat-card { flex: 1; text-align: center; padding: 13px 8px; border-radius: 14px; background: var(--accent-soft); border: 1px solid var(--card-border); }
  .cv-stat-card .num { font-size: 26px; font-weight: 900; color: var(--accent-dark); display: block; line-height: 1.2; }
  .cv-stat-card .lbl { font-size: 11.5px; color: #34495E; font-weight: 600; }

  .cv-table-frame { position: relative; margin-top: 30px; }
  .cv-pill-tab { position: absolute; top: -18px; right: 20px; background: var(--accent); color: #fff; border-radius: 999px; padding: 7px 22px; font-size: 12.5px; font-weight: 800; box-shadow: 0 3px 8px var(--card-border); z-index: 2; }
  .cv-table-wrap { border-radius: 24px; overflow: hidden; border: 1px solid var(--accent); background: #fff; }
  table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 13.5px; }
  thead th { background: #fff; color: var(--accent-dark); font-weight: 800; padding: 16px 4px 12px; text-align: center; font-size: 11px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; border-bottom: 2px solid var(--accent); }
  tbody td { padding: 9px 5px; text-align: center; border-bottom: 1px solid var(--card-border); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; background: var(--table-tint); }
  tbody tr:last-child td { border-bottom: none; }
  .cv-idx { color: var(--accent-dark); font-weight: 700; }
  .cv-name { text-align: right; font-weight: 700; color: #1B2631; font-size: 12px; white-space: normal; overflow: visible; text-overflow: clip; }
  .cv-name .name-text { line-height: 1.3; }
  .status-pill { display: inline-block; margin-top: 4px; padding: 2px 11px; border-radius: 999px; background: var(--pill-bg); border: 1px solid var(--pill-border); color: var(--accent-dark); font-size: 9px; font-weight: 600; line-height: 1.6; white-space: normal; }
  .cv-mono { font-family: 'Consolas', monospace; direction: ltr; color: var(--accent-dark); font-weight: 600; }
  .cv-num { font-weight: 800; color: #1B2631; }
  .cv-eligible { color: #196F3D; }
  .cv-withheld { color: #A93226; }
  .cv-blank { background: #fff !important; }
  .th-blank { white-space: normal !important; font-size: 8.5px !important; line-height: 1.2; }
  .blank-chip { display: block; margin: 3px auto 0; width: 65%; height: 18px; border: 1.2px solid #5D6D7E; border-radius: 6px; background: #fff; }
  tbody tr { page-break-inside: avoid; }

  .cv-footer { margin-top: 16px; display: flex; justify-content: space-between; font-size: 10.5px; color: #85929E; font-weight: 600; }

  .cv-cover { text-align: center; padding-top: 60px; }
  .cv-cover h1 { font-size: 36px; color: var(--accent-dark); font-weight: 800; margin-bottom: 18px; }
  .cv-cover .cv-summary-grid { display: flex; flex-wrap: wrap; gap: 12px; justify-content: center; margin-top: 32px; }
  .cv-cover .cv-summary-card { width: 145px; padding: 16px 8px; border-radius: 16px; background: var(--accent-soft); border: 1px solid var(--card-border); }
  .cv-cover .cv-summary-card .num { display: block; font-size: 26px; font-weight: 900; color: var(--accent-dark); }
  .cv-cover .cv-summary-card .lbl { font-size: 11.5px; color: #34495E; font-weight: 600; }
"""

def _colgroup_html_canva(show_referral=None):
    return _colgroup_html(show_referral)

def _category_section_html_canva(rows, cat, card_col_name, agent_label, with_break=False):
    agent_label = esc(agent_label)
    title, subtitle = esc(cat["title"]), esc(cat["subtitle"].format(agent=agent_label))
    accent, accent_soft, accent_dark = cat["accent"], cat["accent_soft"], cat["accent_dark"]
    badge_label, show_referral = esc(cat["badge_label"]), cat["show_referral"]

    # نفس تخفيف الـ25% هنا للحدود والشبكة الخلفية — عدا شريط العنوان
    # والشارات (cv-pill-tab) اللي تبقى بلون var(--accent) الصافي. خلفية
    # الجدول (--table-tint 5%) تبقى بدون تغيير كما هي (طلب محدد سابق).
    style_vars = (
        f"--accent:{accent}; --accent-soft:{accent_soft}; --accent-dark:{accent_dark};"
        f"--table-tint:{_rgba(accent, 0.05)}; --card-border:{_rgba(accent, 0.225)}; --grid-line:{_rgba(accent_dark, 0.034)};"
        f"--pill-bg:{_rgba(accent, 0.16)}; --pill-border:{_rgba(accent, 0.35)};"
    )

    total_people = sum(int(r.get("الأفراد الكلية", 0) or 0) for r in rows)
    total_eligible = sum(int(r.get("الأفراد المستحقة", 0) or 0) for r in rows)
    total_withheld = sum(int(r.get("الأفراد المحجوبين", 0) or 0) for r in rows)

    rows_html = ""
    for i, r in enumerate(rows, start=1):
        referral_text = esc(r.get('الإحالة', '')) if show_referral else ""
        status_pill = f"<div class='status-pill'>{referral_text}</div>" if referral_text else ""
        rows_html += f"""
        <tr>
          <td class="cv-idx">{i}</td>
          <td class="cv-mono">{esc(r.get(card_col_name, ''))}</td>
          <td class="cv-name"><div class="name-text">{esc(r.get('اسم رب الأسرة', ''))}</div>{status_pill}</td>
          <td class="cv-blank"><span class="blank-chip"></span></td>
          <td class="cv-num">{esc(r.get('الأفراد الكلية', ''))}</td>
          <td class="cv-num cv-eligible">{esc(r.get('الأفراد المستحقة', ''))}</td>
          <td class="cv-num cv-withheld">{esc(r.get('الأفراد المحجوبين', ''))}</td>
        </tr>"""

    section_class = "cv-section with-break" if with_break else "cv-section"
    return f"""
    <section class="{section_class}" style="{style_vars}">
      <div class="cv-header">
        <h1>{title}</h1>
        <p>{subtitle}</p>
        <div class="cv-agent">الوكيل: {agent_label}</div>
      </div>
      <div class="cv-stats">
        <div class="cv-stat-card"><span class="num">{len(rows)}</span><span class="lbl">عدد العوائل ({badge_label})</span></div>
        <div class="cv-stat-card"><span class="num">{total_people}</span><span class="lbl">إجمالي الأفراد</span></div>
        <div class="cv-stat-card"><span class="num">{total_eligible}</span><span class="lbl">الأفراد المستحقة</span></div>
        <div class="cv-stat-card"><span class="num">{total_withheld}</span><span class="lbl">الأفراد المحجوبين</span></div>
      </div>
      <div class="cv-table-frame">
        <div class="cv-pill-tab">{badge_label}</div>
        <div class="cv-table-wrap">
          <table>
            {_colgroup_html_canva()}
            <thead><tr><th>ت</th><th>{esc(card_col_name)}</th><th>اسم رب الأسرة</th><th class="th-blank">حقل فارغ</th><th>الكلية</th><th>المستحقة</th><th>المحجوبين</th></tr></thead>
            <tbody>{rows_html}</tbody>
          </table>
        </div>
      </div>
      <div class="cv-footer">
        <span>نظام المقارنة الشامل والذكي — وكيل رقم {agent_label}</span>
        <span>عدد السجلات: {len(rows)}</span>
      </div>
    </section>"""

def _build_category_pdf_html_canva(rows, cat, card_col_name, agent_label):
    section = _category_section_html_canva(rows, cat, card_col_name, agent_label, with_break=False)
    return _wrap_pdf_document(cat["title"], section, css=_PDF_CSS_CANVA)

def _colgroup_html(show_referral=None):
    # الترتيب: ت، رقم البطاقة، الاسم، عمود فاصل فارغ (خلفية بيضاء دائماً)،
    # ثم باقي البيانات (كلي/مستحق/محجوب). عمود الاسم واسع يكفي الاسم
    # الرباعي الكامل + فقاعة الحالة تحته بسطر واحد متوازي بدون قص "...".
    widths = [4, 14, 35, 8, 13, 13, 13]
    return "<colgroup>" + "".join(f'<col style="width:{w}%">' for w in widths) + "</colgroup>"

def _category_section_html(rows, cat, card_col_name, agent_label, with_break=False):
    agent_label = esc(agent_label)
    title, subtitle = esc(cat["title"]), esc(cat["subtitle"].format(agent=agent_label))
    accent, accent_soft, accent_dark = cat["accent"], cat["accent_soft"], cat["accent_dark"]
    badge_label, icon, show_referral = esc(cat["badge_label"]), cat["icon"], cat["show_referral"]

    # تخفيف تركيز كل الألوان الزخرفية (خلفيات، حدود، ظلال، خطوط فاصلة)
    # بنسبة 25% تجاه الأبيض — ما عدا شريط العنوان والشارات (icon-badge،
    # agent-pill) اللي تستخدم var(--accent)/var(--accent-dark) الصافية
    # بدون أي تخفيف، عشان تبقى العناوين واضحة وقوية دائماً.
    style_vars = (
        f"--accent:{accent}; --accent-soft:{accent_soft}; --accent-dark:{accent_dark};"
        f"--glass-bg:{_rgba(accent, 0.105)}; --glass-border:{_rgba(accent, 0.30)};"
        f"--glass-shadow:{_rgba(accent_dark, 0.15)}; --line-color:{_rgba(accent_dark, 0.21)};"
        f"--pill-bg:{_rgba(accent, 0.16)}; --pill-border:{_rgba(accent, 0.35)};"
    )

    total_people = sum(int(r.get("الأفراد الكلية", 0) or 0) for r in rows)
    total_eligible = sum(int(r.get("الأفراد المستحقة", 0) or 0) for r in rows)
    total_withheld = sum(int(r.get("الأفراد المحجوبين", 0) or 0) for r in rows)

    rows_html = ""
    for i, r in enumerate(rows, start=1):
        referral_text = esc(r.get('الإحالة', '')) if show_referral else ""
        status_pill = f"<div class='status-pill'>{referral_text}</div>" if referral_text else ""
        rows_html += f"""
        <tr>
          <td class="c-idx">{i}</td>
          <td class="c-mono">{esc(r.get(card_col_name, ''))}</td>
          <td class="c-name"><div class="name-text">{esc(r.get('اسم رب الأسرة', ''))}</div>{status_pill}</td>
          <td class="c-blank"><span class="blank-chip"></span></td>
          <td class="c-num">{esc(r.get('الأفراد الكلية', ''))}</td>
          <td class="c-num c-eligible">{esc(r.get('الأفراد المستحقة', ''))}</td>
          <td class="c-num c-withheld">{esc(r.get('الأفراد المحجوبين', ''))}</td>
        </tr>"""

    section_class = "report-section with-break" if with_break else "report-section"
    return f"""
    <section class="{section_class}" style="{style_vars}">
      <div class="header">
        <div class="icon-badge">{icon}</div>
        <h1>{title}</h1>
        <div class="pills">
          <div class="desc-pill">{subtitle}</div>
          <div class="agent-pill">الوكيل: {agent_label}</div>
        </div>
      </div>
      <div class="stats">
        <div class="stat-card"><span class="num">{len(rows)}</span><span class="lbl">عدد العوائل ({badge_label})</span></div>
        <div class="stat-card"><span class="num">{total_people}</span><span class="lbl">إجمالي الأفراد</span></div>
        <div class="stat-card"><span class="num">{total_eligible}</span><span class="lbl">الأفراد المستحقة</span></div>
        <div class="stat-card"><span class="num">{total_withheld}</span><span class="lbl">الأفراد المحجوبين</span></div>
      </div>
      <div class="table-wrap">
        <table>
          {_colgroup_html()}
          <thead><tr><th>ت</th><th>{esc(card_col_name)}</th><th>اسم رب الأسرة</th><th class="th-blank">حقل فارغ</th><th>الكلية</th><th>المستحقة</th><th>المحجوبين</th></tr></thead>
          <tbody>{rows_html}</tbody>
        </table>
      </div>
      <div class="footer">
        <span>نظام المقارنة الشامل والذكي — وكيل رقم {agent_label}</span>
        <span>عدد السجلات: {len(rows)}</span>
      </div>
    </section>"""

def _build_category_pdf_html(rows, cat, card_col_name, agent_label):
    section = _category_section_html(rows, cat, card_col_name, agent_label, with_break=False)
    return _wrap_pdf_document(cat["title"], section)

def _derive_agent_label(new_file_name):
    agent_label = new_file_name.replace(".docx", "").replace(".xlsx", "")
    agent_label = re.sub(r'(FOOD|FLOUR)', '', agent_label, flags=re.IGNORECASE)
    agent_label = re.sub(r'[._-]?pdf[_-]?\d*$', '', agent_label, flags=re.IGNORECASE)
    return agent_label.strip("- ").strip()

def _matched_categories(df_results_full):
    all_rows = df_results_full.to_dict("records")
    matched = []
    for cat in CATEGORY_DEFS:
        rows = [r for r in all_rows if cat["match"](r)]
        if rows:
            matched.append((cat, rows))
    return matched

def create_category_pdf_reports(df_results_full, card_col_name, new_file_name, template="glass"):
    """يبني تقرير PDF أنيق مستقل لكل حالة من حالات المتغيرات المكتشفة
    (مضافة، منقولة، حجب كلي/جزئي، رفع حجب، زيادة/نقصان أفراد أو مستحقين،
    تغيير اسم، تحديث عام)، ويُرجع فقط الحالات التي فعلاً لها سجلات ضمن
    نتيجة المقارنة الحالية. template: "glass" (الافتراضي) أو "canva"."""
    agent_label = _derive_agent_label(new_file_name)
    builder = _build_category_pdf_html_canva if template == "canva" else _build_category_pdf_html
    reports = []
    for cat, rows in _matched_categories(df_results_full):
        pdf_bytes = WeasyHTML(string=builder(rows, cat, card_col_name, agent_label)).write_pdf()
        pdf_buffer = BytesIO(pdf_bytes)
        pdf_buffer.seek(0)
        reports.append({
            "key": cat["key"],
            "button_label": f"📄 تحميل PDF - {cat['title'].replace('تقرير ', '')}",
            "file_name": f"{cat['title'].replace('تقرير ', '')} لـ الوكيل {agent_label}.pdf",
            "pdf": pdf_buffer,
        })
    return reports, agent_label

def create_combined_pdf_report(df_results_full, card_col_name, new_file_name, template="glass"):
    """يبني ملف PDF واحد يجمع كل حالات المتغيرات المكتشفة معاً: صفحة غلاف
    تلخّص أعداد كل حالة، تليها كل حالة بقسمها المستقل بنفس تصميمها ولونها
    (كل حالة تبدأ بصفحة جديدة). template: "glass" (الافتراضي) أو "canva"."""
    agent_label = _derive_agent_label(new_file_name)
    matched = _matched_categories(df_results_full)
    if not matched:
        return None, agent_label

    cover_agent_label = esc(agent_label)
    cover_accent, cover_soft, cover_dark = "#154360", "#EBF5FB", "#0B2E4F"

    if template == "canva":
        cover_style = (
            f"--accent:{cover_accent}; --accent-soft:{cover_soft}; --accent-dark:{cover_dark};"
            f"--card-border:{_rgba(cover_accent, 0.30)}; --grid-line:{_rgba(cover_dark, 0.045)};"
        )
        summary_cards = "".join(
            f'<div class="cv-summary-card"><span class="num">{len(rows)}</span><span class="lbl">{cat["badge_label"]}</span></div>'
            for cat, rows in matched
        )
        cover_html = f"""
        <section class="cv-section" style="{cover_style}">
          <div class="cv-cover">
            <h1>التقرير الشامل لكل حالات المتغيرات</h1>
            <div class="cv-agent" style="position:static;">الوكيل: {cover_agent_label}</div>
            <div class="cv-summary-grid">{summary_cards}</div>
          </div>
        </section>"""
        sections_html = "".join(
            _category_section_html_canva(rows, cat, card_col_name, agent_label, with_break=True)
            for cat, rows in matched
        )
        pdf_bytes = WeasyHTML(string=_wrap_pdf_document("التقرير الشامل", cover_html + sections_html, css=_PDF_CSS_CANVA)).write_pdf()
    else:
        cover_style = (
            f"--accent:{cover_accent}; --accent-soft:{cover_soft}; --accent-dark:{cover_dark};"
            f"--glass-bg:{_rgba(cover_accent, 0.14)}; --glass-border:{_rgba(cover_accent, 0.40)};"
            f"--glass-shadow:{_rgba(cover_dark, 0.20)};"
        )
        summary_cards = "".join(
            f'<div class="summary-card"><span class="num">{len(rows)}</span><span class="lbl">{cat["badge_label"]}</span></div>'
            for cat, rows in matched
        )
        cover_html = f"""
        <section class="report-section" style="{cover_style}">
          <div class="cover">
            <div class="icon-badge" style="margin-bottom:16px;">★</div>
            <h1>التقرير الشامل لكل حالات المتغيرات</h1>
            <div class="pills"><div class="agent-pill">الوكيل: {cover_agent_label}</div></div>
            <div class="summary-grid">{summary_cards}</div>
          </div>
        </section>"""
        sections_html = "".join(
            _category_section_html(rows, cat, card_col_name, agent_label, with_break=True)
            for cat, rows in matched
        )
        pdf_bytes = WeasyHTML(string=_wrap_pdf_document("التقرير الشامل", cover_html + sections_html)).write_pdf()

    pdf_buffer = BytesIO(pdf_bytes)
    pdf_buffer.seek(0)
    return pdf_buffer, agent_label


def main():
    # =============================================================================
    # إعدادات واجهة المستخدم وتنسيقات الـ CSS للويب
    # =============================================================================
    st.set_page_config(page_title="نظام المقارنة المتطور للوكلاء", layout="wide")
    st.markdown("""
        <style>
        th, td { text-align: right !important; dir: rtl !important; white-space: nowrap !important; }
        div.stButton > button { background-color: #2C3E50; color: white; width: 100%; font-weight: bold; border-radius: 8px;}
        .report-box { background-color: #ECF0F1; padding: 15px; border-radius: 8px; border-right: 5px solid #2C3E50; text-align: right; margin-bottom: 10px;}
        div[data-testid="stRadio"] > label { font-weight: bold; color: #2C3E50; font-size: 16px; }
        .date-badge { display: inline-block; padding: 8px 12px; background-color: #F8F9F9; color: #2C3E50; border-radius: 5px; font-weight: bold; font-size: 15px; border: 1px solid #BDC3C7; margin-bottom: 5px; direction: rtl; width: 100%; text-align: center;}
        .date-badge span.old { color: #C0392B; }
        .date-badge span.new { color: #27AE60; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown("<h1 style='text-align: right;'>نظام المقارنة الشامل والذكي 📄🔎</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: right;'>تمت إعادة صياغة وهيكلة ملف الـ Word الناتج برمجياً وتدوير العناوين وتنسيق الصفوف التبادلية بدقة فائقة.</p>", unsafe_allow_html=True)
    # -----------------------------------------------------------------------------
    # 6. الواجهة الرئيسية
    # -----------------------------------------------------------------------------
    st.markdown("<h3 style='text-align: right;'>📂 منطقة الرفع والمطابقة</h3>", unsafe_allow_html=True)
    uploaded_files = st.file_uploader("ارفع ملفي الشهر السابق والحالي معاً", type=['docx', 'xlsx'], accept_multiple_files=True)

    if uploaded_files and len(uploaded_files) == 2:
        st.markdown("<h4 style='text-align: right;'>👁️ معاينة الأعمدة المكتشفة</h4>", unsafe_allow_html=True)
        preview_cols = st.columns(2)
        previews_ok = []
        for pf, pcol in zip(uploaded_files, preview_cols):
            with pcol:
                st.markdown(f"**{pf.name}**")
                preview = preview_columns_for_file(pf)
                pf.seek(0)
                if preview is None:
                    st.warning("ما قدرنا نكتشف جدول بعناوين واضحة بهذا الملف — راح يعتمد على محرك احتياطي أقدم عند المقارنة.")
                    previews_ok.append(False)
                else:
                    if preview.get("merged_header"):
                        st.caption("🗂️ عناوين هذا الملف كلها بخلية واحدة (صيغة قديمة شائعة) — استخدمنا محرك متخصص بهذا الشكل وتم تفكيك العناوين والبيانات بنجاح:")
                    else:
                        st.caption("الأعمدة المكتشفة ← نص العنوان بالملف:")
                    st.dataframe(pd.DataFrame(list(preview["detected"].items()), columns=["الدور", "العنوان بالملف"]), hide_index=True, use_container_width=True)
                    if preview["sample_records"]:
                        st.caption("عيّنة (أول سجلين):")
                        st.dataframe(pd.DataFrame(preview["sample_records"]), hide_index=True, use_container_width=True)
                    previews_ok.append(True)

        st.info("🔧 ما تحتاج تأكيد يدوي — النظام يفحص ويصحح اكتشاف الأعمدة تلقائياً لحظة الضغط على زر المقارنة، ويوقف نفسه بس لو لقى تناقض حقيقي بالأرقام (مو مجرد شك شكلي).")

    col_opts1, col_opts2, col_opts3 = st.columns(3)
    with col_opts1: comparison_mode = st.radio("🎯 نوع المقارنة:", ["النوع الأول", "النوع الثاني", "النوع الثالث", "النموذج الرابع (المستحق فقط)", "النموذج الخامس (كشف تلقائي بالعناوين)"], horizontal=True)
    with col_opts2: card_choice_ui = st.radio("💳 البطاقة المعتمدة:", ["تلقائي (الأنسب للمطابقة)", "رقم البطاقة القديم", "رقم البطاقة الحديث"], horizontal=True)
    with col_opts3: matching_engine = st.radio("⚙️ محرك المطابقة المستهدف:", ["المحرك القياسي", "محرك تخطي التسلسل (بطاقة فقط)"], horizontal=True)

    card_type_auto = (card_choice_ui == "تلقائي (الأنسب للمطابقة)")
    card_type_param = "old" if card_choice_ui == "رقم البطاقة القديم" else "new"
    card_col_name = card_choice_ui if not card_type_auto else "رقم البطاقة القديم"
    swap_files = st.checkbox("🔄 **عكس الملفين يدوياً (القديم يصبح حديثاً والحديث قديماً)**")
    pdf_template_ui = st.radio("🎨 نمط تصميم تقارير PDF:", ["الافتراضي (زجاجي)", "كانفا"], horizontal=True)
    pdf_template = "canva" if pdf_template_ui == "كانفا" else "glass"

    if st.button("🔧 تصحيح الأعمدة تلقائياً وبدء المقارنة الذكية"):
        if len(uploaded_files) == 2:
            with st.spinner('جاري التحليل وعزل الحالات تلقائياً...'):
                file1, file2 = uploaded_files[0], uploaded_files[1]
                ext1 = file1.name.split('.')[-1].lower()
                ext2 = file2.name.split('.')[-1].lower()

                if {ext1, ext2} == {"xlsx", "docx"}:
                    # قاعدة ثابتة: عند رفع ملف إكسل وملف وورد معاً، يُعتمد الإكسل دائماً كالملف
                    # السابق (القديم) والوورد دائماً كالملف الحديث، بغض النظر عن التاريخ المستشعر
                    file_a_is_older = (ext1 == "xlsx")
                else:
                    date1, date2 = extract_document_date(file1), extract_document_date(file2)
                    file_a_is_older = (date1 < date2) if (date1 and date2) else True

                if swap_files: file_a_is_older = not file_a_is_older

                if file_a_is_older:
                    file_old, file_new = file1, file2
                    old_name, new_name = file1.name, file2.name
                else:
                    file_old, file_new = file2, file1
                    old_name, new_name = file2.name, file1.name

                st.markdown(f"<div class='date-badge'>الملف المعتمد كـ <span class='old'>السابق: ({esc(old_name)})</span> | الملف المعتمد كـ <span class='new'>الحديث: ({esc(new_name)})</span></div>", unsafe_allow_html=True)

                old_duplicates, new_duplicates = [], []
                is_eligible_only_mode = (comparison_mode == "النموذج الرابع (المستحق فقط)")

                # توجيه النظام حسب نوع النموذج
                if is_eligible_only_mode:
                    st.caption("ℹ️ نموذج \"المستحق فقط\" يقرأ ترتيب أعمدة ثابت مسبقاً (مو حسب العناوين المكتشفة بالمعاينة أعلاه) — تأكد إن ترتيب أعمدة ملفيك يطابق: تسلسل، (عمود)، بطاقة، اسم، (عمود)، مستحق.")
                    old_data, old_duplicates = extract_eligible_only_records(file_old)
                    new_data, new_duplicates = extract_eligible_only_records(file_new)
                    card_col_name = "رقم البطاقة القديم"
                else:
                    # المحرك الذكي بالتعرف على العناوين هو الأدق (يقرأ عناوين
                    # الجدول الفعلية بدل تخمين ترتيب الأعمدة)، فنجربه أولاً لكل
                    # ملف بشكل مستقل تماماً — لو فشل بملف واحد بس (مثلاً عناوينه
                    # مدمجة بخلية وحدة)، يرجع للمحرك القديم لهذا الملف تحديداً
                    # فقط، ولا يسحب معه الملف الثاني اللي نجح فيه الذكي. هذا يمنع
                    # سيناريو فعلي شوهد: فشل الذكي بملف واحد يهبط الملفين مع بعض
                    # للمحرك القديم، والقديم يسرّب رقم بطاقة لعمود عدد بالملف
                    # الثاني اللي كان سليماً تماماً بالذكي.
                    fallback_card_type = "old" if card_type_auto else card_type_param
                    old_data, old_duplicates, old_used_fallback = _extract_with_smart_fallback(file_old, fallback_card_type)
                    new_data, new_duplicates, new_used_fallback = _extract_with_smart_fallback(file_new, fallback_card_type)
                    used_fallback_engine = old_used_fallback or new_used_fallback

                    if used_fallback_engine:
                        fallback_files = [esc(n) for n, used in ((old_name, old_used_fallback), (new_name, new_used_fallback)) if used]
                        st.caption(f"⚠️ المحرك الذكي ما لقى جدول بعناوين واضحة بـ: {'، '.join(fallback_files)} — استخدمنا المحرك الاحتياطي لهذا الملف تحديداً فقط.")

                    if card_type_auto:
                        old_data, new_data = merge_records_by_either_card(old_data, new_data)
                    card_col_name = "رقم البطاقة" if card_type_auto else card_choice_ui

                    if not old_data or not new_data:
                        st.error("❌ ما قدرنا نستخرج أي سجل من ملف واحد أو أكثر (حتى بالمحرك الاحتياطي). تأكد إن الملفات تحتوي فعلاً جدول بيانات، مو ملف فارغ أو بصيغة غير مدعومة.")
                        st.stop()

                if old_duplicates or new_duplicates:
                    with st.expander(f"⚠️ {len(old_duplicates) + len(new_duplicates)} رقم بطاقة مكرر داخل نفس الملف (اعتُمد أول ظهور فقط، تجاهلنا الباقي)"):
                        for card, name in old_duplicates:
                            st.markdown(f"- [{esc(old_name)}] بطاقة {esc(card)} مكررة — '{esc(name)}'")
                        for card, name in new_duplicates:
                            st.markdown(f"- [{esc(new_name)}] بطاقة {esc(card)} مكررة — '{esc(name)}'")

                if card_type_auto and not is_eligible_only_mode:
                    st.caption("🔎 تم استخدام رقم البطاقة القديم والحديث معاً تلقائياً لتقوية المطابقة بين الملفين.")

                # حارس سلامة البيانات: نفحص وننظّف قبل لا نكمل، مو بعد ما تطلع
                # أرقام غلط بالتقارير. أي سجل فاسد يُستبعد نهائياً من الحساب.
                # لو أول محاولة فيها خلل خطير (فساد منهجي، مو سجل معزول) وما
                # جربنا المحرك القديم بعد، نجربه كفرصة أخيرة قبل ما نوقف نهائياً.
                # نموذج "المستحق فقط" يُصفّر الكلي/المحجوب عمداً فنتجاوز فحص
                # الاتساق الحسابي بينهم (skip_consistency_check).
                is_safe, clean_old, clean_new, validation_errors = validate_and_clean_pair(old_data, new_data, old_name, new_name, skip_consistency_check=is_eligible_only_mode)
                if not is_safe and not is_eligible_only_mode and not used_fallback_engine:
                    if card_type_auto:
                        old_data, new_data, card_col_name, _, _ = extract_matched_by_either_card(extract_clean_records, file_old, file_new)
                    else:
                        old_data, _ = extract_clean_records(file_old, card_type=card_type_param)
                        new_data, _ = extract_clean_records(file_new, card_type=card_type_param)
                    is_safe, clean_old, clean_new, validation_errors = validate_and_clean_pair(old_data, new_data, old_name, new_name, skip_consistency_check=is_eligible_only_mode)

                if not is_safe:
                    st.error("❌ توقفت المقارنة: القيم المستخرجة من الملفات غير منطقية (على الأغلب خلل بقراءة الأعمدة)، ولن أكمل الحساب عليها. تفاصيل أول 10 أخطاء:")
                    for err in validation_errors[:10]:
                        st.markdown(f"- {esc(err)}")
                    st.info("راجع ترتيب/عناوين أعمدة الملفين، أو جرب النموذج الخامس (كشف تلقائي بالعناوين) يدوياً.")
                    st.stop()

                old_data, new_data = clean_old, clean_new
                if validation_errors:
                    with st.expander(f"⚠️ {len(validation_errors)} سجل مستبعد لعدم منطقية قيمه (لن يدخل أي حساب أو تقرير)"):
                        for err in validation_errors[:20]:
                            st.markdown(f"- {esc(err)}")

                results, results_ref, counters = process_comparison(old_data, new_data, comparison_mode, card_col_name, matching_engine)

                if results:
                    results = sorted(results, key=lambda x: (str(x.get("اسم رب الأسرة", "")), x.get("meta_sort", 0)))
                    results_ref = sorted(results_ref, key=lambda x: str(x.get("اسم رب الأسرة", "")))

                    df_results = pd.DataFrame(results)
                    df_results_full = df_results.copy()
                    df_display = df_results.copy()

                    if comparison_mode == "النوع الثاني":
                        for idx, row in df_display.iterrows():
                            if row.get("meta_status") == "type2_new":
                                df_display.at[idx, "التسلسل"], df_display.at[idx, "اسم رب الأسرة"], df_display.at[idx, card_col_name], df_display.at[idx, "الإحالة"] = "", "", "", ""

                    st.markdown(f"<h3 style='text-align: right;'>📋 المخرجات الشاشاتية ({comparison_mode})</h3>", unsafe_allow_html=True)

                    styled_df = df_display.style.apply(lambda d: style_all_types(d, old_data, new_data, card_col_name, comparison_mode), axis=None)
                    if comparison_mode == "النوع الثاني": cols_order = ["التسلسل", "اسم رب الأسرة", card_col_name, "الحالة", "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين", "الإحالة"]
                    else: cols_order = ["التسلسل", "اسم رب الأسرة", card_col_name, "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين", "الإحالة"]

                    st.dataframe(styled_df, use_container_width=True, hide_index=True, column_order=cols_order)

                    base_name = new_name.rsplit('.', 1)[0]
                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        word_report = create_word_table_report(df_results_full, f"تقرير - {comparison_mode}", comparison_mode, card_col_name, old_data, new_data, new_name)
                        st.download_button(label="📥 تحميل المخرجات Word بالتصميم الجديد المطور والمقفل", data=word_report, file_name=f"تقرير_{base_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    with col_dl2:
                        word_stats = create_word_stats_report(counters, base_name)
                        st.download_button(label="📊 تحميل تقرير الإحصاء Word", data=word_stats, file_name=f"احصائيات_{base_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

                    category_reports, agent_label = create_category_pdf_reports(df_results_full, card_col_name, new_name, template=pdf_template)
                    if category_reports:
                        st.markdown("<h4 style='text-align: right;'>📁 تقارير PDF منفصلة لكل حالة من حالات المتغيرات</h4>", unsafe_allow_html=True)

                        combined_pdf, _ = create_combined_pdf_report(df_results_full, card_col_name, new_name, template=pdf_template)
                        if combined_pdf:
                            st.download_button(label="📚 تحميل تقرير PDF شامل يجمع كل الحالات", data=combined_pdf, file_name=f"التقرير الشامل لـ الوكيل {agent_label}.pdf", mime="application/pdf", key="pdf_combined")

                        pdf_cols = st.columns(2)
                        for idx, rep in enumerate(category_reports):
                            with pdf_cols[idx % 2]:
                                st.download_button(label=rep["button_label"], data=rep["pdf"], file_name=rep["file_name"], mime="application/pdf", key=f"pdf_{rep['key']}")

                else:
                    st.success("🎉 تطابق تام! لا توجد فروقات بين الملفين.")
        else:
            st.warning("⚠️ يرجى رفع ملفين اثنين بالضبط للتمكن من بدء المقارنة.")


if __name__ == "__main__":
    main()
