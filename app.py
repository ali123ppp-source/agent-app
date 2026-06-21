import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import re
from datetime import datetime

# =============================================================================
# إعدادات واجهة المستخدم (CSS لتنسيق الاتجاهات والألوان المتقدمة)
# =============================================================================
st.set_page_config(page_title="نظام المقارنة المتطور للوكلاء", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2C3E50; color: white; width: 100%; font-weight: bold; border-radius: 8px;}
    .report-box { background-color: #ECF0F1; padding: 15px; border-radius: 8px; border-right: 5px solid #2C3E50; text-align: right; margin-bottom: 10px;}
    div[data-testid="stRadio"] > label { font-weight: bold; color: #2C3E50; font-size: 16px; }
    .date-badge { display: inline-block; padding: 8px 12px; background-color: #E8F8F5; color: #16A085; border-radius: 5px; font-weight: bold; font-size: 14px; border: 1px solid #1ABC9C; margin-bottom: 15px; direction: rtl;}
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام المقارنة الشامل والذكي 📄🔎</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>نظام ذكي يقوم بقراءة الملفات، تحديد التواريخ تلقائياً، واعتماد رقم البطاقة القديم كمحور للمقارنة.</p>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 1. محرك الاستشعار الزمني ثلاثي الطبقات (الذكي والآمن)
# -----------------------------------------------------------------------------
def extract_document_date(doc):
    """يبحث عن التاريخ النصي بصيغ متعددة، وإذا لم يجده يعتمد على خصائص الملف الوصفية المخفية"""
    patterns = [
        r"([A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})", # Sunday, May 17, 2026
        r"(\d{1,2}[/-]\d{1,2}[/-]\d{4})",               # 17/05/2026 أو 17-05-2026
        r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})"                # 2026/05/17
    ]
    
    # الطبقة الأولى: البحث في تذييل الصفحة (Footer)
    for section in doc.sections:
        footer = section.footer
        if footer:
            for para in reversed(footer.paragraphs):
                for pattern in patterns:
                    match = re.search(pattern, para.text)
                    if match:
                        try:
                            date_str = match.group(1)
                            if "-" in date_str or "/" in date_str:
                                return pd.to_datetime(date_str, dayfirst=True).to_pydatetime()
                            return datetime.strptime(date_str, "%A, %B %d, %Y")
                        except: continue
                        
    # الطبقة الثانية: البحث في الأسطر الأخيرة من متن النص
    for para in reversed(doc.paragraphs):
        for pattern in patterns:
            match = re.search(pattern, para.text)
            if match:
                try:
                    date_str = match.group(1)
                    if "-" in date_str or "/" in date_str:
                        return pd.to_datetime(date_str, dayfirst=True).to_pydatetime()
                    return datetime.strptime(date_str, "%A, %B %d, %Y")
                except: continue
                
    # الطبقة الثالثة (الحل السحري): استخراج تاريخ التعديل والحفظ الفعلي للملف من الـ Metadata
    try:
        if doc.core_properties.modified:
            return doc.core_properties.modified.replace(tzinfo=None)
    except:
        pass

    return None

# -----------------------------------------------------------------------------
# 2. محرك الاستخراج المرن (يدعم الجداول والنصوص المفصولة بفواصل)
# -----------------------------------------------------------------------------
def extract_clean_records(doc):
    records = {}
    
    # الإستراتيجية الأولى: فحص أسطر النص العادية (دعم البيانات المفصولة بفواصل CSV كالموجودة بملفك)
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text: continue
        cells = [c.strip() for c in text.split(',')]
        
        # التأكد من مطابقة السطر لبنية البيانات المطلوبة (الرقم الإحصائي والاسم)
        if len(cells) >= 6 and any(char.isdigit() for char in cells[0]) and any('\u0600' <= char <= '\u06FF' for char in cells[3]):
            try:
                withheld = int(cells[0])
                eligible = int(cells[1])
                total = int(cells[2])
                name = cells[3]
                old_card = cells[4] # اعتماد حقل رقم البطاقة القديم دائماً كمعرف أساسي
                seq = cells[6] if len(cells) > 6 else "-"
                
                if old_card:
                    records[old_card] = {
                        "seq": seq, "name": name, "total": total, 
                        "eligible": eligible, "withheld": withheld
                    }
            except ValueError:
                continue

    # الإستراتيجية الثانية: فحص الجداول (إذا كان الملف يحتوي على جداول حقيقية منظمّة)
    if not records:
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "اسم رب" in "".join(cells):
                    continue
                
                name_idx = -1
                max_len = 0
                for i, c in enumerate(cells):
                    if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
                        if len(c) > max_len:
                            max_len = len(c)
                            name_idx = i
                if name_idx == -1: continue
                
                card_indices = [i for i, c in enumerate(cells) if c.isdigit() and len(c) >= 5]
                if not card_indices: continue
                
                # اعتماد رقم البطاقة القديم (الفهرس الأول عادةً في بنية الملفات)
                old_card = cells[card_indices[0]]
                    
                seq = "-"
                for i in range(len(cells)-1, card_indices[-1], -1):
                    if cells[i].isdigit():
                        seq = cells[i]
                        break
                
                digit_cells = [int(cells[i]) for i in range(name_idx) if cells[i].isdigit()]
                if len(digit_cells) >= 3:
                    withheld, eligible, total = digit_cells[0], digit_cells[1], digit_cells[2]
                elif len(digit_cells) == 2:
                    withheld, eligible, total = 0, digit_cells[0], digit_cells[1]
                else:
                    continue
                    
                records[old_card] = {
                    "seq": seq, "name": cells[name_idx], "total": total, 
                    "eligible": eligible, "withheld": withheld
                }
    return records

# -----------------------------------------------------------------------------
# 3. محرك المقارنة وبناء المصفوفات بالاعتماد على رقم البطاقة القديم
# -----------------------------------------------------------------------------
def process_comparison(old_data, new_data, mode):
    results = []
    results_type_1_reference = []
    
    counters = {
        "total_fam": 0, "eligible_fam": 0, "withheld_fam": 0, "added_fam": 0, "deleted_fam": 0,
        "inc_total": 0, "dec_total": 0, "net_total": 0,
        "inc_eligible": 0, "dec_eligible": 0, "net_eligible": 0,
        "inc_withheld": 0, "dec_withheld": 0, "net_withheld": 0
    }
    
    all_cards = set(old_data.keys()).union(set(new_data.keys()))
    
    for card in all_cards:
        if card in old_data and card in new_data:
            old_v, new_v = old_data[card], new_data[card]
            diff_total = old_v["total"] != new_v["total"]
            diff_elig = old_v["eligible"] != new_v["eligible"]
            diff_with = old_v["withheld"] != new_v["withheld"]
            is_changed = diff_total or diff_elig or diff_with
            
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
                
                # بناء الجدول المرجعي الثابت دوماً بأسفل النظام
                results_type_1_reference.append({
                    "التسلسل": old_v["seq"], "رقم البطاقة القديم": card, "اسم رب الأسرة": old_v["name"],
                    "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                    "meta_status": "modified"
                })
                
                if mode == "النوع الأول":
                    results.append({
                        "التسلسل": old_v["seq"], "رقم البطاقة القديم": card, "اسم رب الأسرة": old_v["name"],
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                        "meta_status": "modified", "meta_sort": 1
                    })
                elif mode == "النوع الثاني":
                    # صف معطيات الشهر السابق
                    results.append({
                        "التسلسل": old_v["seq"], "اسم رب الأسرة": old_v["name"], "الحالة": "السابق", "رقم البطاقة القديم": card,
                        "الأفراد الكلية": old_v["total"], "الأفراد المستحقة": old_v["eligible"], "الأفراد المحجوبين": old_v["withheld"],
                        "meta_status": "type2_old", "meta_card": card, "meta_sort": 1
                    })
                    # صف معطيات الشهر الحالي
                    results.append({
                        "التسلسل": old_v["seq"], "اسم رب الأسرة": old_v["name"], "الحالة": "الحديث", "رقم البطاقة القديم": card,
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                        "meta_status": "type2_new", "meta_card": card, "meta_sort": 2
                    })
                elif mode == "النوع الثالث":
                    results.append({
                        "التسلسل": old_v["seq"], "رقم البطاقة القديم": card, "اسم رب الأسرة": old_v["name"],
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                        "meta_status": "modified", "meta_sort": 1
                    })
            elif mode == "النوع الثالث":
                results.append({
                    "التسلسل": old_v["seq"], "رقم البطاقة القديم": card, "اسم رب الأسرة": old_v["name"],
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
            
            base_row = {"التسلسل": old_v["seq"], "رقم البطاقة القديم": card, "اسم رب الأسرة": old_v["name"] + " (محذوف / منقول)",
                        "الأفراد الكلية": old_v["total"], "الأفراد المستحقة": old_v["eligible"], "الأفراد المحجوبين": old_v["withheld"]}
            
            results_type_1_reference.append({**base_row, "meta_status": "deleted"})
            
            if mode == "النوع الثاني":
                results.append({
                    "التسلسل": old_v["seq"], "اسم رب الأسرة": old_v["name"] + " (محذوف)", "الحالة": "محذوف", "رقم البطاقة القديم": card,
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
            
            base_row = {"التسلسل": new_v["seq"], "رقم البطاقة القديم": card, "اسم رب الأسرة": new_v["name"] + " (مضاف حديثاً)",
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"]}
            
            results_type_1_reference.append({**base_row, "meta_status": "added"})
            
            if mode == "النوع الثاني":
                results.append({
                    "التسلسل": new_v["seq"], "اسم رب الأسرة": new_v["name"] + " (مضاف)", "الحالة": "مضاف", "رقم البطاقة القديم": card,
                    "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                    "meta_status": "added", "meta_card": card, "meta_sort": 1
                })
            else:
                results.append({**base_row, "meta_status": "added", "meta_sort": 1})
                
    return results, results_type_1_reference, counters

# -----------------------------------------------------------------------------
# 4. دوال التلوين البصري وتظليل الخلايا المتغيرة (لشاشات العرض المباشر)
# -----------------------------------------------------------------------------
def style_type_two(df, old_data, new_data):
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    for idx, row in df.iterrows():
        status = row.get("meta_status", "")
        card = row.get("meta_card", "")
        
        if status == "type2_old":
            styles.loc[idx, "الحالة"] = 'background-color: #E0E0E0; font-weight: bold;' # تظليل رمادي للسابق
        elif status == "type2_new":
            styles.loc[idx, "الحالة"] = 'background-color: #C8E6C9; font-weight: bold;' # تظليل أخضر للحديث
            
        if status in ["type2_old", "type2_new"] and card in old_data and card in new_data:
            o_val, n_val = old_data[card], new_data[card]
            if o_val["total"] != n_val["total"]: styles.loc[idx, "الأفراد الكلية"] = 'background-color: #FDE0DC;'
            if o_val["eligible"] != n_val["eligible"]: styles.loc[idx, "الأفراد المستحقة"] = 'background-color: #FDE0DC;'
            if o_val["withheld"] != n_val["withheld"]: styles.loc[idx, "الأفراد المحجوبين"] = 'background-color: #FDE0DC;'
    return styles

def style_type_three(df, old_data, new_data):
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    for idx, row in df.iterrows():
        status = row["meta_status"]
        card = row["رقم البطاقة القديم"]
        if status == "modified":
            styles.loc[idx] = 'background-color: #FFEBEE;'
            if card in old_data and card in new_data:
                if old_data[card]["total"] != new_data[card]["total"]: styles.loc[idx, "الأفراد الكلية"] = 'background-color: #FFE082; font-weight: bold;'
                if old_data[card]["eligible"] != new_data[card]["eligible"]: styles.loc[idx, "الأفراد المستحقة"] = 'background-color: #FFE082; font-weight: bold;'
                if old_data[card]["withheld"] != new_data[card]["withheld"]: styles.loc[idx, "الأفراد المحجوبين"] = 'background-color: #FFE082; font-weight: bold;'
        elif status == "added": styles.loc[idx] = 'background-color: #E8F5E9; color: #2E7D32;'
        elif status == "deleted": styles.loc[idx] = 'background-color: #ECEFF1; color: #455A64; text-decoration: line-through;'
    return styles

# -----------------------------------------------------------------------------
# 5. دوال تصدير مستندات Word الملونة والمدمجة بدقة تامة
# -----------------------------------------------------------------------------
def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} fill="{color_hex}"/>')
    tcPr.append(shd)

def create_word_table_report(df, title, mode, old_data=None, new_data=None):
    doc = Document()
    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    display_df = df.drop(columns=["meta_status", "meta_card", "meta_sort"], errors="ignore")
    cols = list(display_df.columns)[::-1] # عكس الحقول لتظهر من اليمين لليسار بالـ Word
    
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    for i, col in enumerate(cols):
        table.rows[0].cells[i].text = str(col)
        table.rows[0].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    prev_cells = None
    for idx, row in df.iterrows():
        row_cells = table.add_row().cells
        status = row.get("meta_status", "normal")
        card = row.get("meta_card") if mode == "النوع الثاني" else row.get("رقم البطاقة القديم")
        
        for i, col in enumerate(cols):
            row_cells[i].text = str(row[col]) if pd.notna(row[col]) and row[col] != "" else ""
            row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            if mode == "النوع الثاني":
                if col == "الحالة":
                    if status == "type2_old": set_cell_shading(row_cells[i], "E0E0E0")
                    elif status == "type2_new": set_cell_shading(row_cells[i], "C8E6C9")
                if status in ["type2_old", "type2_new"] and old_data and new_data and card in old_data and card in new_data:
                    if col == "الأفراد الكلية" and old_data[card]["total"] != new_data[card]["total"]: set_cell_shading(row_cells[i], "FDE0DC")
                    if col == "الأفراد المستحقة" and old_data[card]["eligible"] != new_data[card]["eligible"]: set_cell_shading(row_cells[i], "FDE0DC")
                    if col == "الأفراد المحجوبين" and old_data[card]["withheld"] != new_data[card]["withheld"]: set_cell_shading(row_cells[i], "FDE0DC")
            elif mode == "النوع الثالث" and old_data and new_data:
                if status == "modified":
                    set_cell_shading(row_cells[i], "FFEBEE")
                    if card in old_data and card in new_data:
                        if col == "الأفراد الكلية" and old_data[card]["total"] != new_data[card]["total"]: set_cell_shading(row_cells[i], "FFE082")
                        if col == "الأفراد المستحقة" and old_data[card]["eligible"] != new_data[card]["eligible"]: set_cell_shading(row_cells[i], "FFE082")
                        if col == "الأفراد المحجوبين" and old_data[card]["withheld"] != new_data[card]["withheld"]: set_cell_shading(row_cells[i], "FFE082")
                elif status == "added": set_cell_shading(row_cells[i], "E8F5E9")
                elif status == "deleted": set_cell_shading(row_cells[i], "ECEFF1")

        # معالجة دمج الخلايا الحقيقي والمتطور للنوع الثاني لمطابقة الرسم التوضيحي للقرّاء
        if mode == "النوع الثاني":
            if status == "type2_old":
                prev_cells = row_cells
            elif status == "type2_new" and prev_cells:
                for merge_col in ["التسلسل", "اسم رب الأسرة", "رقم البطاقة القديم"]:
                    if merge_col in cols:
                        m_idx = cols.index(merge_col)
                        prev_cells[m_idx].merge(row_cells[m_idx])
                        row_cells[m_idx].text = ""
                
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def create_word_stats_report(counters, filename_base):
    doc = Document()
    heading = doc.add_heading(f"تقرير الإحصاء الشامل لشهر - {filename_base}", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph().add_run("أولاً: إحصاء حركة الأفراد").bold = True
    stats_individuals = [
        ("إجمالي زيادة الأفراد الكلية (المضافين):", f"+{counters['inc_total']}"),
        ("إجمالي نقصان الأفراد الكلية (المحذوفين):", f"-{counters['dec_total']}"),
        ("صافي التغيير في الأفراد الكلية:", f"{counters['net_total']:+d}"),
        ("إجمالي زيادة الأفراد المستحقة:", f"+{counters['inc_eligible']}"),
        ("إجمالي نقصان الأفراد المستحقة:", f"-{counters['dec_eligible']}"),
        ("صافي التغيير في الأفراد المستحقة:", f"{counters['net_eligible']:+d}"),
        ("إجمالي زيادة الأفراد المحجوبين:", f"+{counters['inc_withheld']}"),
        ("إجمالي نقصان الأفراد المحجوبين:", f"-{counters['dec_withheld']}"),
        ("صافي التغيير في الأفراد المحجوبين:", f"{counters['net_withheld']:+d}")
    ]
    for text, val in stats_individuals:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(f" {val} ").bold = True
        p.add_run(f" : {text}")
        
    doc.add_paragraph().add_run().add_break()
    doc.add_paragraph().add_run("ثانياً: إحصاء القيود والعوائل (إداري)").bold = True
    stats_families = [
        ("عوائل تغيرت أعداد أفرادها الكلية:", counters['total_fam']),
        ("عوائل تغيرت أعداد أفرادها المستحقة:", counters['eligible_fam']),
        ("عوائل تغيرت أعداد أفرادها المحجوبين:", counters['withheld_fam']),
        ("عوائل جديدة تمت إضافتها بالكامل:", counters['added_fam']),
        ("عوائل تم نقلها أو حذفها بالكامل:", counters['deleted_fam'])
    ]
    for text, val in stats_families:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(f" {val} ").bold = True
        p.add_run(f" : {text}")
        
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# 6. واجهة المستخدم الذكية والتشغيل التلقائي الكامل
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 منطقة رفع الملفات الذكية</h3>", unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "ارفع ملفي الشهر السابق والشهر الحالي معاً هنا (سيتولى النظام فرز من هو الأقدم والأحدث تلقائياً بذكاء)", 
    type=['docx'], accept_multiple_files=True
)

comparison_mode = st.radio(
    "🎯 اختر نوع نظام المقارنة المطلوب:",
    ["النوع الأول", "النوع الثاني", "النوع الثالث"],
    format_func=lambda x: {
        "النوع الأول": "النوع الأول: عرض الأسطر المتغيرة فقط",
        "النوع الثاني": "النوع الثاني: صفين لكل عائلة (سابق/حديث) وتظليل الخلية المتغيرة",
        "النوع الثالث": "النوع الثالث: عرض كافة السجلات وتظليل الخلايا"
    }[x], horizontal=True
)

if st.button("بدء المقارنة الذكية واستخراج المتغيرات"):
    if len(uploaded_files) == 2:
        with st.spinner('جاري تحليل الملفات وقراءة التواريخ الذكية والبدء بالمطابقة...'):
            doc1 = Document(uploaded_files[0])
            doc2 = Document(uploaded_files[1])
            
            date1 = extract_document_date(doc1)
            date2 = extract_document_date(doc2)
            
            # ترتيب فرز وتعيين الملفات زمنياً بناءً على محرك الاستشعار
            if date1 and date2:
                if date1 < date2:
                    old_doc, new_doc = doc1, doc2
                    old_name, new_name = uploaded_files[0].name, uploaded_files[1].name
                    st.markdown(f"<div align='right' class='date-badge'>✓ استشعار تلقائي زمني ناجح | الملف السابق: ({date1.strftime('%Y-%m-%d')}) ➔ الملف الحالي: ({date2.strftime('%Y-%m-%d')})</div>", unsafe_allow_html=True)
                else:
                    old_doc, new_doc = doc2, doc1
                    old_name, new_name = uploaded_files[1].name, uploaded_files[0].name
                    st.markdown(f"<div align='right' class='date-badge'>✓ استشعار تلقائي زمني ناجح | الملف السابق: ({date2.strftime('%Y-%m-%d')}) ➔ الملف الحالي: ({date1.strftime('%Y-%m-%d')})</div>", unsafe_allow_html=True)
            else:
                st.warning("⚠️ لم يعثر النظام على تاريخ نصي صريح، تم تطبيق طبقة الحماية الثانية للاعتماد على خصائص المستندات للفرز.")
                old_doc, new_doc = doc1, doc2
                old_name, new_name = uploaded_files[0].name, uploaded_files[1].name
            
            # معالجة القراءة والمطابقة
            old_data = extract_clean_records(old_doc)
            new_data = extract_clean_records(new_doc)
            
            results, results_ref, counters = process_comparison(old_data, new_data, comparison_mode)
            
            if results:
                # ترتيب أبجدي صارم مع تثبيت تسلسل الصفوف التابعة لنفس العائلة بالنوع الثاني
                results = sorted(results, key=lambda x: (str(x.get("اسم رب الأسرة", "")), x.get("meta_sort", 0)))
                results_ref = sorted(results_ref, key=lambda x: str(x.get("اسم رب الأسرة", "")))
                
                df_results = pd.DataFrame(results)
                
                # إخفاء بصري للنصوص المتكررة بالنوع الثاني داخل عرض Streamlit لمحاكاة الدمج (الدمج الفعلي يظهر بالـ Word المطبوع)
                if comparison_mode == "النوع الثاني":
                    for idx, row in df_results.iterrows():
                        if row.get("meta_status") == "type2_new":
                            df_results.at[idx, "التسلسل"] = ""
                            df_results.at[idx, "اسم رب الأسرة"] = ""
                            df_results.at[idx, "رقم البطاقة القديم"] = ""
                
                st.markdown(f"<h3 style='text-align: right;'>📋 جدول المخرجات الرئيسية ({comparison_mode})</h3>", unsafe_allow_html=True)
                
                if comparison_mode == "النوع الثالث":
                    styled_df = df_results.style.apply(lambda d: style_type_three(d, old_data, new_data), axis=None)
                    st.dataframe(styled_df, use_container_width=True, hide_index=True, column_order=[c for c in df_results.columns if not c.startswith("meta_")])
                elif comparison_mode == "النوع الثاني":
                    styled_df = df_results.style.apply(lambda d: style_type_two(d, old_data, new_data), axis=None)
                    cols_order = ["التسلسل", "اسم رب الأسرة", "الحالة", "رقم البطاقة القديم", "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين"]
                    st.dataframe(styled_df, use_container_width=True, hide_index=True, column_order=cols_order)
                else:
                    cols_order = ["التسلسل", "رقم البطاقة القديم", "اسم رب الأسرة", "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين"]
                    st.dataframe(df_results, use_container_width=True, hide_index=True, column_order=cols_order)
                
                base_name = new_name.rsplit('.', 1)[0]
                
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    word_report = create_word_table_report(df_results, f"تقرير مخرجات - {comparison_mode}", comparison_mode, old_data, new_data)
                    st.download_button(label=f"📥 تحميل المخرجات ({comparison_mode}) Word", data=word_report, file_name=f"تقرير_{comparison_mode}_{base_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                with col_dl2:
                    word_stats = create_word_stats_report(counters, base_name)
                    st.download_button(label="📊 تحميل تقرير الإحصاء الوزاري Word", data=word_stats, file_name=f"احصائيات_{base_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                
                # دمج وإضافة الجدول المرجعي الثابت (المماثل للنوع الأول باعتماد رقم البطاقة القديم) دائماً بالأسفل
                st.markdown("---")
                st.markdown("<h3 style='text-align: right; color: #7F8C8D;'>📌 الجدول المرجعي الثابت (مماثل للنوع الأول معتمد على رقم البطاقة القديم)</h3>", unsafe_allow_html=True)
                df_ref = pd.DataFrame(results_ref)
                cols_order_ref = ["التسلسل", "رقم البطاقة القديم", "اسم رب الأسرة", "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين"]
                st.dataframe(df_ref, use_container_width=True, hide_index=True, column_order=cols_order_ref)
                
            else:
                st.success("🎉 تطابق تام! لا توجد فروقات أو متغيرات بين الملفين المرفوعين.")
    else:
        st.warning("⚠️ يرجى رفع ملفين اثنين بالضبط في صندوق الرفع للتمكن من بدء إجراء المقارنة السليمة.")
