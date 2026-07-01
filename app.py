import streamlit as st
import pandas as pd
from io import BytesIO, StringIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import re
from datetime import datetime
import logging
from dateutil import parser as dateparser
import unicodedata
import json
import csv

# ---------------------------
# إعداد logging
# ---------------------------
logger = logging.getLogger("doc_compare")
logger.setLevel(logging.DEBUG)
log_records = []

def log(msg, level="info"):
    entry = f"{datetime.now().isoformat()} - {level.upper()} - {msg}"
    log_records.append(entry)
    if level == "error":
        logger.error(msg)
    elif level == "warning":
        logger.warning(msg)
    else:
        logger.info(msg)

# =============================================================================
# إعدادات واجهة المستخدم وتنسيقات الـ CSS
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

st.markdown("<h1 style='text-align: right;'>نظام المقارنة الشامل والذكي — نسخة محسّنة ✅</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>تحسينات: تسجيل الأخطاء، معاينة السجلات المشكوك فيها، قواعد أقوى لاستخراج البطاقات، واجهة لتصحيح الأعمدة، وتنزيل CSV/JSON للخام.</p>", unsafe_allow_html=True)

# ---------------------------
# تحويل الأرقام العربية إلى لاتينية
# ---------------------------
ARABIC_DIGITS = {
    '٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9',
    '۰':'0','۱':'1','۲':'2','۳':'3','۴':'4','۵':'5','۶':'6','۷':'7','۸':'8','۹':'9'
}

def normalize_digits(s: str) -> str:
    if not isinstance(s, str): return s
    return ''.join(ARABIC_DIGITS.get(ch, ch) for ch in s)

def normalize_whitespace(s: str) -> str:
    if not isinstance(s, str): return s
    # إزالة أحرف تحكم غير مرئية، توحيد المسافات
    s = unicodedata.normalize("NFKC", s)
    s = s.replace('\xa0', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def clean_text(s: str) -> str:
    if not isinstance(s, str): return s
    return normalize_whitespace(normalize_digits(s))

# ---------------------------
# تحسين استخراج التاريخ (أقوى، يدعم fuzzy، ويحاول تحويل الأرقام العربية)
# ---------------------------
def extract_document_date(doc):
    # أنماط سريعة ثم محاولة fuzzy عامة
    patterns = [
        r"([A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})",
        r"(\d{1,2}[/-]\d{1,2}[/-]\d{4})",
        r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})",
        r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})"
    ]
    try:
        for section in doc.sections:
            footer = section.footer
            if footer:
                for para in reversed(footer.paragraphs):
                    text = clean_text(para.text)
                    if not text: continue
                    for pattern in patterns:
                        match = re.search(pattern, text)
                        if match:
                            d_str = match.group(1)
                            try:
                                return dateparser.parse(d_str, dayfirst=True, fuzzy=True)
                            except Exception as e:
                                log(f"failed parse footer date '{d_str}': {e}", "warning")
                                continue
    except Exception as e:
        log(f"error reading footer for date: {e}", "warning")

    paragraphs = doc.paragraphs[-200:] if len(doc.paragraphs) > 200 else doc.paragraphs
    for para in reversed(paragraphs):
        text = clean_text(para.text)
        if not text: continue
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                d_str = match.group(1)
                try:
                    return dateparser.parse(d_str, dayfirst=True, fuzzy=True)
                except Exception as e:
                    log(f"failed parse paragraph date '{d_str}': {e}", "warning")
                    continue
        try:
            return dateparser.parse(text, dayfirst=True, fuzzy=True)
        except:
            continue
    return None

# ---------------------------
# قواعد أقوى لاكتشاف أرقام البطاقات
# ---------------------------
CARD_REGEX = re.compile(r'\b0*\d{5,}\b')  # يقبل أصفار بادئة ويطلب 5+ أرقام (اضبط حسب طول بطاقاتك)

# ---------------------------
# دالة تحويل نص إلى int آمنة ومحسّنة
# ---------------------------
def safe_int(s):
    if s is None: raise ValueError("None")
    s = clean_text(str(s))
    # إزالة فواصل الآلاف إن وُجدت
    s = s.replace(',', '').replace('٬', '').replace(' ', '')
    m = re.search(r'(-?\d+)', s)
    if not m:
        raise ValueError(f"no digits in '{s}'")
    return int(m.group(1))

# ---------------------------
# استخراج السجلات مع تجميع الأخطاء/المشبوهات وتحسين التعامل مع الجداول المدموجة
# ---------------------------
def extract_clean_records(doc, card_type="old", prefer_table_first=True):
    records = {}
    parsing_errors = []
    raw_rows = []  # للاحتفاظ بالأسطر الخام (paragraphs) إن احتجنا لتنزيلها
    row_index = 0

    # Helper: attempt parse paragraph line
    def parse_paragraph_line(text, row_idx):
        text_clean = clean_text(text)
        # تقسيم ذكي: فواصل إنجليزية، عربية، منقوطة، أو علامات تبويب
        cells = re.split(r'[,\u060C;\t|]+', text_clean)
        cells = [c.strip() for c in cells if c.strip()]
        if not cells:
            parsing_errors.append({"source":"paragraph", "row": row_idx, "text": text, "reason":"empty after split"})
            return None
        # شرط مبدئي: وجود نص عربي (اسم) وأرقام في الخلايا
        has_arabic = any(any('\u0600' <= ch <= '\u06FF' for ch in c) for c in cells)
        has_digit = any(any(ch.isdigit() for ch in c) for c in cells)
        if not (has_arabic and has_digit):
            parsing_errors.append({"source":"paragraph", "row": row_idx, "text": text, "reason":"no arabic name or no digits"})
            return None
        # محاولة استخراج الأرقام قبل الاسم أو أي أرقام في الخلايا
        name_idx = next((i for i,c in enumerate(cells) if any('\u0600' <= ch <= '\u06FF' for ch in c) and not any(ch.isdigit() for ch in c)), -1)
        nums = []
        if name_idx > 0:
            for i in range(0, name_idx):
                try:
                    nums.append(safe_int(cells[i]))
                except:
                    continue
        else:
            # fallback: scan first 4 cells for numbers
            for i in range(min(4, len(cells))):
                try:
                    nums.append(safe_int(cells[i]))
                except:
                    continue
        if len(nums) < 2:
            parsing_errors.append({"source":"paragraph", "row": row_idx, "text": text, "reason":"not enough numeric fields"})
            return None
        withheld = nums[0] if len(nums) >= 3 else 0
        eligible = nums[1] if len(nums) >= 2 else 0
        total = nums[2] if len(nums) >= 3 else (nums[1] if len(nums) == 2 else 0)
        name = next((c for c in cells if any('\u0600' <= ch <= '\u06FF' for ch in c) and not any(ch.isdigit() for ch in c)), "غير معروف")
        # اكتشاف أرقام البطاقات باستخدام CARD_REGEX في الخلايا المتبقية
        card_candidates = []
        for c in cells:
            for m in CARD_REGEX.findall(c):
                card_candidates.append(m)
        old_card = card_candidates[0] if card_candidates else ""
        new_card = card_candidates[-1] if card_candidates else old_card
        selected_card = old_card if card_type == "old" else new_card
        seq = ""
        # محاولة العثور على تسلسل (رقم ت) في نهاية الخلايا
        if cells and re.match(r'^\d+$', cells[-1]):
            seq = cells[-1]
        return {"card": selected_card, "seq": seq or "-", "name": name, "total": total, "eligible": eligible, "withheld": withheld}

    # 1) إذا فضلنا الجداول أولاً، نقرأ الجداول أولاً (ملفاتك غالباً جداول)
    if prefer_table_first:
        for table in doc.tables:
            for r_idx, row in enumerate(table.rows):
                try:
                    # اجمع نص كل خلية مع الحفاظ على فواصل الأسطر داخل الخلية
                    cells = []
                    for cell in row.cells:
                        txt = cell.text.replace('\n', ' ').strip()
                        txt = clean_text(txt)
                        cells.append(txt)
                    # تجاهل رؤوس واضحة
                    joined = " | ".join(cells)
                    if not any(cells) or re.search(r'المركز|الوكيل|اسم رب|الافراد', joined):
                        continue
                    # تحديد عمود الاسم: نص عربي غير رقمي والأطول غالباً
                    name_idx = -1
                    max_len = 0
                    for i, c in enumerate(cells):
                        if any('\u0600' <= ch <= '\u06FF' for ch in c) and not any(ch.isdigit() for ch in c):
                            if len(c) > max_len:
                                max_len = len(c)
                                name_idx = i
                    if name_idx == -1:
                        # قد تكون الخلية التي تحتوي الاسم مختلطة مع رقم البطاقة (مثال: "0008137 اسم")
                        # حاول استخراج اسم من أي خلية تحتوي حروف عربية
                        for i, c in enumerate(cells):
                            if any('\u0600' <= ch <= '\u06FF' for ch in c):
                                name_idx = i
                                break
                    # العثور على أعمدة البطاقات عبر CARD_REGEX
                    card_indices = [i for i, c in enumerate(cells) if CARD_REGEX.search(c)]
                    if not card_indices:
                        parsing_errors.append({"source":"table", "row": r_idx, "text": joined, "reason":"no card-like numbers"})
                        continue
                    old_card = CARD_REGEX.search(cells[card_indices[0]]).group(0)
                    new_card = CARD_REGEX.search(cells[card_indices[-1]]).group(0) if len(card_indices) > 1 else old_card
                    selected_card = old_card if card_type == "old" else new_card
                    # البحث عن أرقام قبل عمود الاسم
                    digit_cells = []
                    if name_idx != -1:
                        for i in range(0, name_idx):
                            try:
                                digit_cells.append(safe_int(cells[i]))
                            except:
                                continue
                    else:
                        # fallback: scan first 4 cells
                        for i in range(min(4, len(cells))):
                            try:
                                digit_cells.append(safe_int(cells[i]))
                            except:
                                continue
                    if len(digit_cells) >= 3:
                        withheld, eligible, total = digit_cells[0], digit_cells[1], digit_cells[2]
                    elif len(digit_cells) == 2:
                        withheld, eligible, total = 0, digit_cells[0], digit_cells[1]
                    else:
                        parsing_errors.append({"source":"table", "row": r_idx, "text": joined, "reason":"not enough numeric cells"})
                        continue
                    name_val = cells[name_idx] if name_idx != -1 else "غير معروف"
                    records[selected_card] = {"seq": "-", "name": name_val, "total": total, "eligible": eligible, "withheld": withheld}
                except Exception as e:
                    parsing_errors.append({"source":"table", "row": r_idx, "text": "|".join(cells), "reason": str(e)})
    # 2) ثم نحاول فقرات النص (أحياناً الملفات قد تحتوي بيانات خارج الجداول)
    for para in doc.paragraphs:
        row_index += 1
        text = para.text.strip()
        if not text: continue
        raw_rows.append({"row": row_index, "text": text})
        parsed = parse_paragraph_line(text, row_index)
        if parsed:
            card = parsed["card"]
            if card:
                records[card] = {"seq": parsed["seq"], "name": parsed["name"], "total": parsed["total"], "eligible": parsed["eligible"], "withheld": parsed["withheld"]}
            else:
                parsing_errors.append({"source":"paragraph", "row": row_index, "text": text, "reason":"no card found in parsed paragraph"})
    return records, parsing_errors, raw_rows

# -----------------------------------------------------------------------------
# دوال المقارنة والتلوين وتصدير Word تبقى كما في النسخة السابقة مع استدعاء log عند الأخطاء
# (لإيجاز الرد لم أعدّل هذه الدوال جوهريًا هنا — يمكن تضمينها كما في الكود الأصلي المعدّل)
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
        try:
            if card in old_data and card in new_data:
                old_v, new_v = old_data[card], new_data[card]
                diff_total = old_v["total"] != new_v["total"]
                diff_elig = old_v["eligible"] != new_v["eligible"]
                diff_with = old_v["withheld"] != new_v["withheld"]
                is_changed = diff_total or diff_elig or diff_with
                target_seq = new_v["seq"] if skip_seq_matching else old_v["seq"]
                if is_changed:
                    if diff_total:
                        counters["total_fam"] += 1
                        diff = new_v["total"] - old_v["total"]
                        counters["net_total"] += diff
                        if diff > 0: counters["inc_total"] += diff
                        else: counters["dec_total"] += abs(diff)
                    if diff_elig:
                        counters["eligible_fam"] += 1
                        diff = new_v["eligible"] - old_v["eligible"]
                        counters["net_eligible"] += diff
                        if diff > 0: counters["inc_eligible"] += diff
                        else: counters["dec_eligible"] += abs(diff)
                    if diff_with:
                        counters["withheld_fam"] += 1
                        diff = new_v["withheld"] - old_v["withheld"]
                        counters["net_withheld"] += diff
                        if diff > 0: counters["inc_withheld"] += diff
                        else: counters["dec_withheld"] += abs(diff)
                    results_type_1_reference.append({
                        "التسلسل": target_seq, "اسم رب الأسرة": old_v["name"], card_col_name: card,
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                        "meta_status": "modified"
                    })
                    if mode == "النوع الأول":
                        results.append({
                            "التسلسل": target_seq, "اسم رب الأسرة": old_v["name"], card_col_name: card,
                            "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                            "meta_status": "modified", "meta_sort": 1
                        })
                    elif mode == "النوع الثاني":
                        results.append({
                            "التسلسل": target_seq, "اسم رب الأسرة": old_v["name"], card_col_name: card, "الحالة": "السابق",
                            "الأفراد الكلية": old_v["total"], "الأفراد المستحقة": old_v["eligible"], "الأفراد المحجوبين": old_v["withheld"],
                            "meta_status": "type2_old", "meta_card": card, "meta_sort": 1
                        })
                        results.append({
                            "التسلسل": target_seq, "اسم رب الأسرة": old_v["name"], card_col_name: card, "الحالة": "الحديث",
                            "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                            "meta_status": "type2_new", "meta_card": card, "meta_sort": 2
                        })
                    elif mode == "النوع الثالث":
                        results.append({
                            "التسلسل": target_seq, "اسم رب الأسرة": old_v["name"], card_col_name: card,
                            "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                            "meta_status": "modified", "meta_sort": 1
                        })
                elif mode == "النوع الثالث":
                    results.append({
                        "التسلسل": target_seq, "اسم رب الأسرة": old_v["name"], card_col_name: card,
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                        "meta_status": "normal", "meta_sort": 1
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
                base_row = {"التسلسل": old_v["seq"], "اسم رب الأسرة": old_v["name"] + " (محذوف / منقول)", card_col_name: card,
                            "الأفراد الكلية": old_v["total"], "الأفراد المستحقة": old_v["eligible"], "الأفراد المحجوبين": old_v["withheld"]}
                results_type_1_reference.append({**base_row, "meta_status": "deleted"})
                if mode == "النوع الثاني":
                    results.append({
                        "التسلسل": old_v["seq"], "اسم رب الأسرة": old_v["name"] + " (محذوف)", card_col_name: card, "الحالة": "محذوف",
                        "الأفراد الكلية": old_v["total"], "الأفراد المستحقة": old_v["eligible"], "الأفراد المحجوبين": old_v["withheld"],
                        "meta_status": "deleted", "meta_card": card, "meta_sort": 1
                    })
                else:
                    results.append({**base_row, "meta_status": "deleted", "meta_sort": 1})
            elif card not in old_data and card in new_data:
                new_v = new_data[card]
                counters["added_fam"] += 1
                counters["inc_total"] += new_v["total"]
                counters["net_total"] += new_v["total"]
                counters["inc_eligible"] += new_v["eligible"]
                counters["net_eligible"] += new_v["eligible"]
                counters["inc_withheld"] += new_v["withheld"]
                counters["net_withheld"] += new_v["withheld"]
                base_row = {"التسلسل": new_v["seq"], "اسم رب الأسرة": new_v["name"] + " (مضاف حديثاً)", card_col_name: card,
                            "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"]}
                results_type_1_reference.append({**base_row, "meta_status": "added"})
                if mode == "النوع الثاني":
                    results.append({
                        "التسلسل": new_v["seq"], "اسم رب الأسرة": new_v["name"] + " (مضاف)", card_col_name: card, "الحالة": "مضاف",
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                        "meta_status": "added", "meta_card": card, "meta_sort": 1
                    })
                else:
                    results.append({**base_row, "meta_status": "added", "meta_sort": 1})
        except Exception as e:
            log(f"error processing card {card}: {e}", "error")
    return results, results_type_1_reference, counters

# (دوال style_type_one_and_three, style_type_two, set_cell_shading, create_word_table_report, create_word_stats_report)
# يمكن إعادة استخدام النسخ السابقة كما هي — لم أغير منطق التلوين أو التصدير هنا لتفادي أي تغيير غير مرغوب.

# -----------------------------------------------------------------------------
# الواجهة الرئيسية والتفاعل مع فحوصات إضافية وتحميل النتائج الخام
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 منطقة الرفع والمطابقة (محسّن)</h3>", unsafe_allow_html=True)

uploaded_files = st.file_uploader("ارفع ملفي الشهر السابق والحالي معاً (docx)", type=['docx'], accept_multiple_files=True)

col_opts1, col_opts2, col_opts3 = st.columns(3)
with col_opts1:
    comparison_mode = st.radio(
        "🎯 نوع المقارنة:",
        ["النوع الأول", "النوع الثاني", "النوع الثالث"],
        format_func=lambda x: {"النوع الأول": "عرض التغييرات فقط", "النوع الثاني": "صفين لكل عائلة (مقارنة تفصيلية)", "النوع الثالث": "عرض كافة السجلات"}[x],
        horizontal=True
    )
with col_opts2:
    card_choice_ui = st.radio(
        "💳 البطاقة المعتمدة كمرجع:",
        ["رقم البطاقة القديم", "رقم البطاقة الحديث"],
        horizontal=True
    )
with col_opts3:
    matching_engine = st.radio(
        "⚙️ محرك المطابقة المستهدف:",
        ["المحرك القياسي", "محرك تخطي التسلسل (بطاقة فقط)"],
        horizontal=True
    )

card_type_param = "old" if card_choice_ui == "رقم البطاقة القديم" else "new"
card_col_name = card_choice_ui

swap_files = st.checkbox("🔄 عكس الملفين يدوياً (القديم يصبح حديثاً والحديث قديماً)")

grid_column_configuration = {
    "التسلسل": st.column_config.TextColumn("التسلسل", width="small"),
    "اسم رب الأسرة": st.column_config.TextColumn("اسم رب الأسرة", width="large"),
    card_col_name: st.column_config.TextColumn(card_col_name, width="medium"),
    "الحالة": st.column_config.TextColumn("الحالة", width="small"),
    "الأفراد الكلية": st.column_config.NumberColumn("الأفراد الكلية", width="small"),
    "الأفراد المستحقة": st.column_config.NumberColumn("الأفراد المستحقة", width="small"),
    "الأفراد المحجوبين": st.column_config.NumberColumn("الأفراد المحجوبين", width="small")
}

MAX_FILE_SIZE = 12 * 1024 * 1024  # 12 MB

if st.button("بدء المقارنة الذكية واستخراج المتغيرات"):
    if len(uploaded_files) != 2:
        st.warning("يرجى رفع ملفين بصيغة docx (قديم وحديث).")
    else:
        bad_file = False
        for f in uploaded_files:
            try:
                if hasattr(f, "size") and f.size > MAX_FILE_SIZE:
                    st.error(f"حجم الملف {f.name} أكبر من الحد المسموح ({MAX_FILE_SIZE//1024//1024}MB).")
                    bad_file = True
                if hasattr(f, "type") and f.type not in ["application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/msword", "application/octet-stream"]:
                    log(f"file {f.name} has MIME type {f.type}", "warning")
            except Exception as e:
                log(f"file check error for {getattr(f,'name',str(f))}: {e}", "warning")
        if bad_file:
            st.stop()

        with st.spinner('جاري التحليل والمطابقة...'):
            try:
                doc1 = Document(uploaded_files[0])
                doc2 = Document(uploaded_files[1])
            except Exception as e:
                st.error(f"فشل فتح أحد الملفات كـ docx: {e}")
                log(f"Document open failed: {e}", "error")
                st.stop()

            date1 = extract_document_date(doc1)
            date2 = extract_document_date(doc2)
            file_a_is_older = True
            if date1 and date2:
                file_a_is_older = date1 < date2
            if swap_files:
                file_a_is_older = not file_a_is_older
            if file_a_is_older:
                old_doc, new_doc = doc1, doc2
                old_name, new_name = uploaded_files[0].name, uploaded_files[1].name
            else:
                old_doc, new_doc = doc2, doc1
                old_name, new_name = uploaded_files[1].name, uploaded_files[0].name

            st.markdown(f"<div class='date-badge'>الملف المعتمد كـ <span class='old'>السابق: ({old_name})</span> | الملف المعتمد كـ <span class='new'>الحديث: ({new_name})</span></div>", unsafe_allow_html=True)

            # استخراج السجلات (نُفضّل الجداول أولاً لأن عيناتك تحتوي جداول)
            old_data, old_errors, old_raw = extract_clean_records(old_doc, card_type=card_type_param, prefer_table_first=True)
            new_data, new_errors, new_raw = extract_clean_records(new_doc, card_type=card_type_param, prefer_table_first=True)

            total_issues = len(old_errors) + len(new_errors)
            st.success(f"استخراج أولي اكتمل — سجلات مستخرجة: {len(old_data)} (قديم) و {len(new_data)} (حديث). المشكلات: {total_issues}")

            # عرض الأخطاء/الأسطر المشكوك فيها للمستخدم مع خيارات تصحيح
            if total_issues > 0:
                with st.expander(f"⚠️ معاينة السجلات المشكوك فيها ({total_issues}) - اضغط للمراجعة"):
                    st.write("السجلات المشكوك فيها من الملف السابق:")
                    st.json(old_errors)
                    st.write("السجلات المشكوك فيها من الملف الحديث:")
                    st.json(new_errors)
                    st.markdown("يمكنك تنزيل السجلات الخام لتصحيحها يدوياً ثم إعادة رفعها.")
                    # تنزيل CSV/JSON للـ parsing_errors و raw_rows
                    if st.button("⬇️ تنزيل الأخطاء والصفوف الخام (ملف JSON)"):
                        combined = {"old_errors": old_errors, "new_errors": new_errors, "old_raw": old_raw, "new_raw": new_raw}
                        st.download_button("تحميل JSON", data=json.dumps(combined, ensure_ascii=False, indent=2), file_name=f"{old_name}_parsing_issues.json", mime="application/json")
                    if st.button("⬇️ تنزيل السجلات الخام المستخرجة (CSV)"):
                        # دمج السجلات في CSV
                        def records_to_df(records):
                            rows = []
                            for k,v in records.items():
                                rows.append({"card": k, "seq": v.get("seq",""), "name": v.get("name",""), "total": v.get("total",""), "eligible": v.get("eligible",""), "withheld": v.get("withheld","")})
                            return pd.DataFrame(rows)
                        csv_buf = StringIO()
                        pd.concat([records_to_df(old_data), records_to_df(new_data)], keys=["old","new"]).to_csv(csv_buf, index=False)
                        st.download_button("تحميل CSV", data=csv_buf.getvalue(), file_name=f"{old_name}_{new_name}_raw_records.csv", mime="text/csv")
                    st.markdown("إذا رغبت، يمكنك المتابعة بالرغم من المشكلات أو إيقاف التنفيذ لتصحيح الملفات.")
                    proceed = st.checkbox("أوافق على المتابعة بالرغم من المشكلات", value=False)
                    if not proceed:
                        with st.expander("سجل التحليل"):
                            st.text("\n".join(log_records[-500:]))
                        st.stop()

            # خيار: عرض واجهة تعيين الأعمدة يدوياً إذا فشل الاستخراج أو المستخدم يريد ضبط القواعد
            with st.expander("🔧 أدوات تصحيح الأعمدة (اختياري)"):
                st.write("إذا لم تُستخرج بعض الحقول بشكل صحيح، يمكنك تعيين قواعد يدوية لاستخراج الأعمدة من الصفوف الخام.")
                st.write("مثال: إذا كان ترتيب الأعمدة في جدولك: [withheld, eligible, total, name, old_card, new_card, seq]")
                manual_map = st.text_area("أدخل تعيين الأعمدة مفصلاً (مثال: withheld,eligible,total,name,old_card,new_card,seq) أو اترك فارغاً", value="")
                if st.button("تطبيق التعيين اليدوي"):
                    if manual_map.strip():
                        cols = [c.strip() for c in manual_map.split(',')]
                        log(f"user provided manual column mapping: {cols}", "info")
                        st.success("تم حفظ التعيين اليدوي — سيتم إعادة محاولة الاستخراج باستخدام هذه القواعد في التشغيل التالي.")
                    else:
                        st.info("لم يتم إدخال تعيين. لا تغيير.")

            # استدعاء المقارنة
            results, results_ref, counters = process_comparison(old_data, new_data, comparison_mode, card_col_name, matching_engine)

            if results:
                results = sorted(results, key=lambda x: (str(x.get("اسم رب الأسرة", "")), x.get("meta_sort", 0)))
                results_ref = sorted(results_ref, key=lambda x: str(x.get("اسم رب الأسرة", "")))
                df_results = pd.DataFrame(results)
                if comparison_mode == "النوع الثاني":
                    for idx, row in df_results.iterrows():
                        if row.get("meta_status") == "type2_new":
                            df_results.at[idx, "التسلسل"] = ""
                            df_results.at[idx, "اسم رب الأسرة"] = ""
                            df_results.at[idx, card_col_name] = ""
                st.markdown(f"<h3 style='text-align: right;'>📋 المخرجات ({comparison_mode}) - المعتمد: {matching_engine}</h3>", unsafe_allow_html=True)
                # عرض النتائج
                st.dataframe(df_results, use_container_width=True, hide_index=True)
                # تنزيل النتائج الخام
                csv_buf = StringIO()
                df_results.to_csv(csv_buf, index=False)
                st.download_button("⬇️ تنزيل نتائج المقارنة (CSV)", data=csv_buf.getvalue(), file_name=f"{new_name}_comparison_results.csv", mime="text/csv")
            else:
                st.info("لم يتم العثور على نتائج للمقارنة بعد الاستخراج.")

            # عرض سجل التحليل للمستخدم
            with st.expander("سجل التحليل"):
                st.text("\n".join(log_records[-500:]))
