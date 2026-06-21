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

# إعدادات الواجهة
st.set_page_config(page_title="نظام المقارنة المتطور للوكلاء", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2C3E50; color: white; width: 100%; font-weight: bold; border-radius: 8px;}
    .report-box { background-color: #ECF0F1; padding: 15px; border-radius: 8px; border-right: 5px solid #2C3E50; text-align: right; margin-bottom: 10px;}
    .net-diff { font-size: 16px; font-weight: bold; margin-top: 5px; color: #2C3E50; border-top: 1px solid #BDC3C7; padding-top: 5px;}
    .stat-inc { font-size: 14px; color: #27AE60; font-weight: bold; }
    .stat-dec { font-size: 14px; color: #C0392B; font-weight: bold; }
    div[data-testid="stRadio"] > label { font-weight: bold; color: #2C3E50; font-size: 16px; }
    .date-badge { display: inline-block; padding: 5px 10px; background-color: #E8F8F5; color: #16A085; border-radius: 5px; font-weight: bold; font-size: 14px; border: 1px solid #1ABC9C; margin-bottom: 15px;}
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام المقارنة الشامل والذكي 📄🔎</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>مدمج بمحرك الفرز الزمني التلقائي لتحديد الشهر السابق والحالي بدقة.</p>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# محرك الاستشعار الزمني (الجديد)
# -----------------------------------------------------------------------------
def extract_document_date(doc):
    """يبحث عن التاريخ الإنجليزي في نهاية الملف أو التذييل لتحديد عمر الملف بدقة"""
    pattern = r"([A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})"
    
    # 1. البحث في التذييل (Footer) أولاً لأنه المكان المعتاد
    for section in doc.sections:
        footer = section.footer
        if footer:
            for para in reversed(footer.paragraphs):
                match = re.search(pattern, para.text)
                if match:
                    try: return datetime.strptime(match.group(1), "%A, %B %d, %Y")
                    except ValueError: continue
                    
    # 2. البحث في النصوص العادية من الأسفل للأعلى كبديل احتياطي
    for para in reversed(doc.paragraphs):
        match = re.search(pattern, para.text)
        if match:
            try: return datetime.strptime(match.group(1), "%A, %B %d, %Y")
            except ValueError: continue
            
    return None # لم يتم العثور على تاريخ

# -----------------------------------------------------------------------------
# دوال التلوين لملفات الـ Word
# -----------------------------------------------------------------------------
def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} fill="{color_hex}"/>')
    tcPr.append(shd)

# -----------------------------------------------------------------------------
# محرك الاستخراج
# -----------------------------------------------------------------------------
def extract_clean_records(doc):
    records = {}
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
            card_num = cells[card_indices[1]] if len(card_indices) >= 2 else cells[card_indices[0]]
                
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
                
            records[card_num] = {
                "seq": seq, "name": cells[name_idx], "total": total, 
                "eligible": eligible, "withheld": withheld
            }
    return records

# -----------------------------------------------------------------------------
# محرك المقارنة المطور
# -----------------------------------------------------------------------------
def process_comparison(old_data, new_data, mode):
    results = []
    results_type_1_reference = []
    
    counters = {
        "total_fam": 0, "eligible_fam": 0, "withheld_fam": 0, 
        "added_fam": 0, "deleted_fam": 0,
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
                
                results_type_1_reference.append({
                    "التسلسل": old_v["seq"], "رقم البطاقة": card, "اسم رب الأسرة": old_v["name"],
                    "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                    "meta_status": "modified"
                })
                
                if mode == "النوع الأول":
                    results.append({
                        "التسلسل": old_v["seq"], "رقم البطاقة": card, "اسم رب الأسرة": old_v["name"],
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                        "meta_status": "modified", "meta_sort": 1
                    })
                elif mode == "النوع الثاني":
                    results.append({
                        "التسلسل": old_v["seq"], "اسم رب الأسرة": old_v["name"], "الحالة": "السابق", "رقم البطاقة": card,
                        "الأفراد الكلية": old_v["total"], "الأفراد المستحقة": old_v["eligible"], "الأفراد المحجوبين": old_v["withheld"],
                        "meta_status": "type2_old", "meta_card": card, "meta_sort": 1
                    })
                    results.append({
                        "التسلسل": old_v["seq"], "اسم رب الأسرة": old_v["name"], "الحالة": "الحديث", "رقم البطاقة": card,
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                        "meta_status": "type2_new", "meta_card": card, "meta_sort": 2
                    })
                elif mode == "النوع الثالث":
                    results.append({
                        "التسلسل": old_v["seq"], "رقم البطاقة": card, "اسم رب الأسرة": old_v["name"],
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                        "meta_status": "modified", "meta_sort": 1
                    })
            elif mode == "النوع الثالث":
                results.append({
                    "التسلسل": old_v["seq"], "رقم البطاقة": card, "اسم رب الأسرة": old_v["name"],
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
            
            base_row = {"التسلسل": old_v["seq"], "رقم البطاقة": card, "اسم رب الأسرة": old_v["name"] + " (محذوف / منقول)",
                        "الأفراد الكلية": old_v["total"], "الأفراد المستحقة": old_v["eligible"], "الأفراد المحجوبين": old_v["withheld"]}
            
            results_type_1_reference.append({**base_row, "meta_status": "deleted"})
            
            if mode == "النوع الثاني":
                results.append({
                    "التسلسل": old_v["seq"], "اسم رب الأسرة": old_v["name"] + " (محذوف)", "الحالة": "محذوف", "رقم البطاقة": card,
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
            
            base_row = {"التسلسل": new_v["seq"], "رقم البطاقة": card, "اسم رب الأسرة": new_v["name"] + " (مضاف حديثاً)",
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"]}
            
            results_type_1_reference.append({**base_row, "meta_status": "added"})
            
            if mode == "النوع الثاني":
                results.append({
                    "التسلسل": new_v["seq"], "اسم رب الأسرة": new_v["name"] + " (مضاف)", "الحالة": "مضاف", "رقم البطاقة": card,
                    "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
                    "meta_status": "added", "meta_card": card, "meta_sort": 1
                })
            else:
                results.append({**base_row, "meta_status": "added", "meta_sort": 1})
                
    return results, results_type_1_reference, counters

# -----------------------------------------------------------------------------
# دوال التظليل البصري لجداول Pandas
# -----------------------------------------------------------------------------
def style_type_two(df, old_data, new_data):
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    for idx, row in df.iterrows():
        status = row.get("meta_status", "")
        card = row.get("meta_card", "")
        
        if status == "type2_old":
            styles.loc[idx, "الحالة"] = 'background-color: #E0E0E0; font-weight: bold;'
        elif status == "type2_new":
            styles.loc[idx, "الحالة"] = 'background-color: #C8E6C9; font-weight: bold;'
            
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
        card = row["رقم البطاقة"]
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
# دالة تصدير Word الشاملة لجميع الأنواع
# -----------------------------------------------------------------------------
def create_word_table_report(df, title, mode, old_data=None, new_data=None):
    doc = Document()
    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    display_df = df.drop(columns=["meta_status", "meta_card", "meta_sort"], errors="ignore")
    cols = list(display_df.columns)[::-1]
    
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
        card = row.get("meta_card") if mode == "النوع الثاني" else row.get("رقم البطاقة")
        
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

        if mode == "النوع الثاني":
            if status == "type2_old":
                prev_cells = row_cells
            elif status == "type2_new" and prev_cells:
                for merge_col in ["التسلسل", "اسم رب الأسرة", "رقم البطاقة"]:
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
# واجهة المستخدم الذكية للرفع التلقائي
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 منطقة رفع الملفات المزدوجة</h3>", unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "ارفع ملفي الشهر السابق والشهر الحالي معاً هنا (سيتولى النظام فرز من هو الأقدم والأحدث تلقائياً)", 
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
        with st.spinner('جاري تحليل الملفات، قراءة التواريخ المدمجة، والبدء بالمطابقة...'):
            # تحويل الملفات المرفوعة إلى كائنات Document لقراءتها مرة واحدة
            doc1 = Document(uploaded_files[0])
            doc2 = Document(uploaded_files[1])
            
            # استخراج التواريخ
            date1 = extract_document_date(doc1)
            date2 = extract_document_date(doc2)
            
            # تحديد الملف القديم والجديد بناءً على التاريخ
            if date1 and date2:
                if date1 < date2:
                    old_doc, new_doc = doc1, doc2
                    old_name, new_name = uploaded_files[0].name, uploaded_files[1].name
                    st.markdown(f"<div align='right' class='date-badge'>✓ تم تحديد الملف القديم: ({date1.strftime('%Y-%m-%d')}) | والملف الحديث: ({date2.strftime('%Y-%m-%d')})</div>", unsafe_allow_html=True)
                else:
                    old_doc, new_doc = doc2, doc1
                    old_name, new_name = uploaded_files[1].name, uploaded_files[0].name
                    st.markdown(f"<div align='right' class='date-badge'>✓ تم تحديد الملف القديم: ({date2.strftime('%Y-%m-%d')}) | والملف الحديث: ({date1.strftime('%Y-%m-%d')})</div>", unsafe_allow_html=True)
            else:
                st.warning("⚠️ لم يتمكن النظام من العثور على التواريخ بالصيغة المطلوبة، تم الاعتماد على ترتيب الرفع كإجراء بديل.")
                old_doc, new_doc = doc1, doc2
                old_name, new_name = uploaded_files[0].name, uploaded_files[1].name
            
            # استخراج البيانات والمقارنة
            old_data = extract_clean_records(old_doc)
            new_data = extract_clean_records(new_doc)
            
            results, results_ref, counters = process_comparison(old_data, new_data, comparison_mode)
            
            if results:
                results = sorted(results, key=lambda x: (str(x.get("اسم رب الأسرة", "")), x.get("meta_sort", 0)))
                results_ref = sorted(results_ref, key=lambda x: str(x.get("اسم رب الأسرة", "")))
                
                df_results = pd.DataFrame(results)
                
                if comparison_mode == "النوع الثاني":
                    for idx, row in df_results.iterrows():
                        if row.get("meta_status") == "type2_new":
                            df_results.at[idx, "التسلسل"] = ""
                            df_results.at[idx, "اسم رب الأسرة"] = ""
                            df_results.at[idx, "رقم البطاقة"] = ""
                
                st.markdown(f"<h3 style='text-align: right;'>📋 جدول المخرجات الرئيسية ({comparison_mode})</h3>", unsafe_allow_html=True)
                
                if comparison_mode == "النوع الثالث":
                    styled_df = df_results.style.apply(lambda d: style_type_three(d, old_data, new_data), axis=None)
                    st.dataframe(styled_df, use_container_width=True, hide_index=True, column_order=[c for c in df_results.columns if not c.startswith("meta_")])
                elif comparison_mode == "النوع الثاني":
                    styled_df = df_results.style.apply(lambda d: style_type_two(d, old_data, new_data), axis=None)
                    cols_order = ["التسلسل", "اسم رب الأسرة", "الحالة", "رقم البطاقة", "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين"]
                    st.dataframe(styled_df, use_container_width=True, hide_index=True, column_order=cols_order)
                else:
                    cols_order = ["التسلسل", "رقم البطاقة", "اسم رب الأسرة", "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين"]
                    st.dataframe(df_results, use_container_width=True, hide_index=True, column_order=cols_order)
                
                base_name = new_name.rsplit('.', 1)[0]
                
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    word_report = create_word_table_report(df_results, f"تقرير مخرجات - {comparison_mode}", comparison_mode, old_data, new_data)
                    st.download_button(label=f"📥 تحميل المخرجات ({comparison_mode}) Word", data=word_report, file_name=f"تقرير_{comparison_mode}_{base_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                with col_dl2:
                    word_stats = create_word_stats_report(counters, base_name)
                    st.download_button(label="📊 تحميل تقرير الإحصاء الوزاري Word", data=word_stats, file_name=f"احصائيات_{base_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                
                st.markdown("---")
                st.markdown("<h3 style='text-align: right; color: #7F8C8D;'>📌 الجدول المرجعي (مماثل للنوع الأول باعتماد البطاقة القديمة)</h3>", unsafe_allow_html=True)
                df_ref = pd.DataFrame(results_ref)
                cols_order_ref = ["التسلسل", "رقم البطاقة", "اسم رب الأسرة", "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين"]
                st.dataframe(df_ref, use_container_width=True, hide_index=True, column_order=cols_order_ref)
                
            else:
                st.success("🎉 تطابق تام! لا توجد فروقات بين الشهرين.")
    else:
        st.warning("⚠️ يرجى رفع ملفين اثنين بالضبط في صندوق الرفع بالأعلى للتمكن من إجراء المقارنة.")
