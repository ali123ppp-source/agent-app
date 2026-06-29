import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import RGBColor
import re
from datetime import datetime

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

st.markdown("<h1 style='text-align: right;'>نظام المقارنة الشامل والذكي 📄🔎</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>تمت إضافة ميزة استخراج الحالات في أوراق منفصلة داخل نفس ملف الـ Word مع الحفاظ على التنسيق الاحترافي.</p>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# 1. محرك الاستشعار الزمني
# -----------------------------------------------------------------------------
def extract_document_date(doc):
    patterns = [
        r"([A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})", 
        r"(\d{1,2}[/-]\d{1,2}[/-]\d{4})",               
        r"(\d{4}[/-]\d{1,2}[/-]\d{1,2})"                
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
                    return datetime.strptime(d_str, "%A, %B %d, %Y")
                except: continue
    return None

# -----------------------------------------------------------------------------
# 2. محرك الاستخراج الدقيق
# -----------------------------------------------------------------------------
def extract_clean_records(doc, card_type="old"):
    records = {}
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text: continue
        cells = [c.strip() for c in text.split(',')]
        if len(cells) >= 6 and any(char.isdigit() for char in cells[0]) and any('\u0600' <= char <= '\u06FF' for char in cells[3]):
            try:
                withheld, eligible, total, name = int(cells[0]), int(cells[1]), int(cells[2]), cells[3]
                old_card = cells[4]
                new_card = cells[5] if len(cells) > 5 else old_card
                selected_card = old_card if card_type == "old" else new_card
                seq = cells[6] if len(cells) > 6 else "-"
                if selected_card:
                    records[selected_card] = {"seq": seq, "name": name, "total": total, "eligible": eligible, "withheld": withheld}
            except ValueError: continue

    if not records:
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "اسم رب" in "".join(cells): continue
                name_idx = -1
                max_len = 0
                for i, c in enumerate(cells):
                    if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
                        if len(c) > max_len: max_len, name_idx = len(c), i
                if name_idx == -1: continue
                card_indices = [i for i, c in enumerate(cells) if c.isdigit() and len(c) >= 5]
                if not card_indices: continue
                
                old_card = cells[card_indices[0]]
                new_card = cells[card_indices[-1]] if len(card_indices) > 1 else old_card
                selected_card = old_card if card_type == "old" else new_card
                seq = "-"
                for i in range(len(cells)-1, card_indices[-1], -1):
                    if cells[i].isdigit():
                        seq = cells[i]
                        break
                digit_cells = [int(cells[i]) for i in range(name_idx) if cells[i].isdigit()]
                if len(digit_cells) >= 3: withheld, eligible, total = digit_cells[0], digit_cells[1], digit_cells[2]
                elif len(digit_cells) == 2: withheld, eligible, total = 0, digit_cells[0], digit_cells[1]
                else: continue
                records[selected_card] = {"seq": seq, "name": cells[name_idx], "total": total, "eligible": eligible, "withheld": withheld}
    return records

# -----------------------------------------------------------------------------
# 3. محرك المقارنة ونظام الإحالة الذكي
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
            is_changed = d_tot != 0 or d_elig != 0 or d_with != 0
            
            target_seq = new_v["seq"] if skip_seq_matching else old_v["seq"]
            
            notes = []
            if old_v["name"] != new_v["name"]:
                notes.append(f"تبدل اسم رب الأسرة (السابق: {old_v['name']})")
                is_changed = True
            
            if d_with > 0:
                notes.append(f"حجب {d_with} نفر")
            elif d_with < 0:
                notes.append(f"رفع الحجب عن {abs(d_with)} نفر")
                
            if d_tot > 0 and d_with >= 0:
                notes.append("إضافة طفل")
            
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
                
                base_dict = {
                    "التسلسل": target_seq, "اسم رب الأسرة": new_v["name"], card_col_name: card,
                    "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], 
                    "الأفراد المحجوبين": new_v["withheld"], "الإحالة": referral_text, "meta_card": card
                }
                
                results_type_1_reference.append({**base_dict, "meta_status": "modified"})
                
                if mode == "النوع الأول" or mode == "النوع الثالث":
                    results.append({**base_dict, "meta_status": "modified", "meta_sort": 1})
                elif mode == "النوع الثاني":
                    results.append({
                        "التسلسل": target_seq, "اسم رب الأسرة": old_v["name"], card_col_name: card, "الحالة": "السابق",
                        "الأفراد الكلية": old_v["total"], "الأفراد المستحقة": old_v["eligible"], "الأفراد المحجوبين": old_v["withheld"],
                        "الإحالة": "", "meta_status": "type2_old", "meta_card": card, "meta_sort": 1
                    })
                    results.append({
                        "التسلسل": target_seq, "اسم رب الأسرة": new_v["name"], card_col_name: card, "الحالة": "الحديث",
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], "الأفراد المحجوبين": new_v["withheld"],
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
            
            base_row = {"التسلسل": old_v["seq"], "اسم رب الأسرة": old_v["name"], card_col_name: card,
                        "الأفراد الكلية": old_v["total"], "الأفراد المستحقة": old_v["eligible"], 
                        "الأفراد المحجوبين": old_v["withheld"], "الإحالة": "عائلة محذوفة", "meta_card": card}
            results_type_1_reference.append({**base_row, "meta_status": "deleted"})
            
            if mode == "النوع الثاني":
                results.append({**base_row, "الحالة": "محذوف", "meta_status": "deleted", "meta_card": card, "meta_sort": 1})
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
            
            base_row = {"التسلسل": new_v["seq"], "اسم رب الأسرة": new_v["name"], card_col_name: card,
                        "الأفراد الكلية": new_v["total"], "الأفراد المستحقة": new_v["eligible"], 
                        "الأفراد المحجوبين": new_v["withheld"], "الإحالة": "عائلة مضافة", "meta_card": card}
            results_type_1_reference.append({**base_row, "meta_status": "added"})
            
            if mode == "النوع الثاني":
                results.append({**base_row, "الحالة": "مضاف", "meta_status": "added", "meta_card": card, "meta_sort": 1})
            else:
                results.append({**base_row, "meta_status": "added", "meta_sort": 1})
                
    return results, results_type_1_reference, counters

# -----------------------------------------------------------------------------
# 4. دوال التظليل البصري للويب
# -----------------------------------------------------------------------------
def style_all_types(doc_df, old_data, new_data, card_col_name, mode):
    styles = pd.DataFrame('', index=doc_df.index, columns=doc_df.columns)
    for idx, row in doc_df.iterrows():
        status = row.get("meta_status", "")
        card = row.get("meta_card")
        notes = row.get("الإحالة", "")
        
        if "تبدل" in notes: styles.loc[idx, "الإحالة"] = 'color: #8E44AD; font-weight: bold;'
        elif "إضافة" in notes: styles.loc[idx, "الإحالة"] = 'color: #2980B9; font-weight: bold;'
        elif "حجب" in notes and "رفع" not in notes: styles.loc[idx, "الإحالة"] = 'color: #C0392B; font-weight: bold;'
        elif "رفع" in notes or "مضافة" in notes: styles.loc[idx, "الإحالة"] = 'color: #27AE60; font-weight: bold;'
        
        if status == "type2_old": styles.loc[idx, "الحالة"] = 'background-color: #F5F5F5; font-weight: bold; color: #7F8C8D;'
        elif status == "type2_new": styles.loc[idx, "الحالة"] = 'background-color: #E8F8F5; font-weight: bold; color: #16A085;'
        elif status == "added": styles.loc[idx] = 'background-color: #E8F5E9; color: #2E7D32;'
        elif status == "deleted": styles.loc[idx] = 'background-color: #ECEFF1; color: #455A64; text-decoration: line-through;'

        if status in ["modified", "type2_old", "type2_new"] and card in old_data and card in new_data:
            o_val, n_val = old_data[card], new_data[card]
            if o_val["total"] != n_val["total"]: styles.loc[idx, "الأفراد الكلية"] = 'background-color: #FDE0DC; font-weight: bold; color: #C0392B;'
            if o_val["eligible"] != n_val["eligible"]: styles.loc[idx, "الأفراد المستحقة"] = 'background-color: #FDE0DC; font-weight: bold; color: #C0392B;'
            if o_val["withheld"] != n_val["withheld"]: styles.loc[idx, "الأفراد المحجوبين"] = 'background-color: #FDE0DC; font-weight: bold; color: #C0392B;'
    return styles

# -----------------------------------------------------------------------------
# 5. محرك تصدير مستندات Word المطور (مع ميزة تقسيم الحالات تلقائياً)
# -----------------------------------------------------------------------------
def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} fill="{color_hex}"/>')
    tcPr.append(shd)

def create_word_table_report(doc_df, title, mode, card_col_name, old_data, new_data, new_file_name):
    doc = Document()
    
    # دالة داخلية لبناء الجداول للحفاظ على تماثل الإعدادات والأحجام بدقة احترافية لكل الأوراق
    def append_table_to_doc(target_doc, df_to_write, table_title):
        target_doc.add_heading(f"الملف المعتمد: {new_file_name}", level=2).alignment = WD_ALIGN_PARAGRAPH.CENTER
        target_doc.add_heading(table_title, level=3).alignment = WD_ALIGN_PARAGRAPH.CENTER
        target_doc.add_paragraph()
        
        display_df = df_to_write.drop(columns=["meta_status", "meta_card", "meta_sort"], errors="ignore")
        cols = list(display_df.columns)[::-1]
        
        table = target_doc.add_table(rows=1, cols=len(cols))
        table.style = 'Table Grid'
        table.autofit = True # احترافي لمنع نزول النص إلى سطر ثانٍ وجعله متلائماً مع حجم المحتوى والصفحة
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        for i, col in enumerate(cols):
            table.rows[0].cells[i].text = str(col)
            table.rows[0].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        prev_cells = None
        for idx, row in df_to_write.iterrows():
            row_cells = table.add_row().cells
            status = row.get("meta_status", "normal")
            card = row.get("meta_card")
            
            for i, col in enumerate(cols):
                val_text = str(row[col]) if pd.notna(row[col]) and row[col] != "" else ""
                p = row_cells[i].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # تلوين حقل الإحالة بدقة متناهية بناءً على المعايير المطلوبة
                if col == "الإحالة" and val_text:
                    parts = val_text.split(" | ")
                    for p_idx, part in enumerate(parts):
                        run = p.add_run(part)
                        run.bold = True
                        if "تبدل" in part: run.font.color.rgb = RGBColor(128, 0, 128) # بنفسجي
                        elif "إضافة" in part: run.font.color.rgb = RGBColor(0, 0, 255) # أزرق
                        elif "حجب" in part and "رفع" not in part: run.font.color.rgb = RGBColor(255, 0, 0) # أحمر
                        elif "رفع" in part or "مضافة" in part: run.font.color.rgb = RGBColor(0, 128, 0) # أخضر
                        elif "محذوفة" in part: run.font.color.rgb = RGBColor(255, 0, 0) # أحمر
                        if p_idx < len(parts) - 1:
                            p.add_run(" | ")
                else:
                    p.add_run(val_text)
                
                if mode == "النوع الثاني" and col == "الحالة":
                    if status == "type2_old": set_cell_shading(row_cells[i], "F5F5F5")
                    elif status == "type2_new": set_cell_shading(row_cells[i], "E8F8F5")
                
                if status in ["modified", "type2_old", "type2_new"] and old_data and new_data and card in old_data and card in new_data:
                    if col == "الأفراد الكلية" and old_data[card]["total"] != new_data[card]["total"]: set_cell_shading(row_cells[i], "FDE0DC")
                    if col == "الأفراد المستحقة" and old_data[card]["eligible"] != new_data[card]["eligible"]: set_cell_shading(row_cells[i], "FDE0DC")
                    if col == "الأفراد المحجوبين" and old_data[card]["withheld"] != new_data[card]["withheld"]: set_cell_shading(row_cells[i], "FDE0DC")
                elif status == "added": set_cell_shading(row_cells[i], "E8F5E9")
                elif status == "deleted": set_cell_shading(row_cells[i], "ECEFF1")

            if mode == "النوع الثاني":
                if status == "type2_old":
                    prev_cells = row_cells
                elif status == "type2_new" and prev_cells:
                    for merge_col in ["التسلسل", "اسم رب الأسرة", card_col_name, "الإحالة"]:
                        if merge_col in cols:
                            m_idx = cols.index(merge_col)
                            
                            # اختيار النص المناسب لإبقائه في الخلية المدمجة لضمان عدم اختفاء البيانات
                            if merge_col == "الإحالة": text_to_keep = str(row["الإحالة"])
                            elif merge_col == "اسم رب الأسرة": text_to_keep = str(row["اسم رب الأسرة"])
                            elif merge_col == card_col_name: text_to_keep = str(row[card_col_name])
                            else: text_to_keep = prev_cells[m_idx].text
                                
                            prev_cells[m_idx].merge(row_cells[m_idx])
                            
                            prev_cells[m_idx].text = ""
                            p_merge = prev_cells[m_idx].paragraphs[0]
                            p_merge.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            
                            if merge_col == "الإحالة" and text_to_keep:
                                parts = text_to_keep.split(" | ")
                                for p_idx, part in enumerate(parts):
                                    run = p_merge.add_run(part)
                                    run.bold = True
                                    if "تبدل" in part: run.font.color.rgb = RGBColor(128, 0, 128)
                                    elif "إضافة" in part: run.font.color.rgb = RGBColor(0, 0, 255)
                                    elif "حجب" in part and "رفع" not in part: run.font.color.rgb = RGBColor(255, 0, 0)
                                    elif "رفع" in part or "مضافة" in part: run.font.color.rgb = RGBColor(0, 128, 0)
                                    elif "محذوفة" in part: run.font.color.rgb = RGBColor(255, 0, 0)
                                    if p_idx < len(parts) - 1: p_merge.add_run(" | ")
                            else:
                                p_merge.add_run(text_to_keep)

    # أولاً: إنشاء الورقة الأولى (الجدول الرئيسي الشامل لكل الفروقات)
    append_table_to_doc(doc, doc_df, title)
    
    # ثانياً: تصفية المخرجات واستخراج "كل حالة في ورقة" منفصلة تلقائياً بناءً على طلبك
    cases_to_extract = [
        ("حالات تبدل اسم رب الأسرة", "تبدل اسم رب الأسرة"),
        ("حالات إضافة طفل", "إضافة طفل"),
        ("حالات حجب نفر", "حجب"), 
        ("حالات رفع الحجب عن نفر", "رفع الحجب"),
        ("العوائل المضافة", "عائلة مضافة"),
        ("العوائل المحذوفة", "عائلة محذوفة")
    ]
    
    for case_title, keyword in cases_to_extract:
        # بناء شرط التصفية بدقة
        if keyword == "حجب":
            matched_mask = doc_df['الإحالة'].str.contains("حجب", na=False) & ~doc_df['الإحالة'].str.contains("رفع", na=False) & ~doc_df['الإحالة'].str.contains("محذوفة", na=False)
        else:
            matched_mask = doc_df['الإحالة'].str.contains(keyword, na=False)
            
        # إذا كان النظام من (النوع الثاني)، نحتاج لجلب السطرين (السابق والحديث) معاً للبطاقة المطابقة للشرط
        if mode == "النوع الثاني":
            matched_cards = doc_df[matched_mask]['meta_card'].dropna().unique()
            case_df = doc_df[doc_df['meta_card'].isin(matched_cards)]
        else:
            case_df = doc_df[matched_mask]
            
        # إذا كانت هناك عوائل تنطبق عليها هذه الحالة، ننشئ لها ورقة مستقلة فوراً
        if not case_df.empty:
            doc.add_page_break() # الانتقال لورقة مستقلة جديدة بالكامل داخل المستند
            append_table_to_doc(doc, case_df, f"كشف مستقل: {case_title}")
                
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

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
        ("تغيرت المحجوبين:", counters['withheld_fam']), ("عوائل مضافة:", counters['added_fam']), ("عوائل محذوفة:", counters['deleted_fam'])
    ]
    for text, val in stats_families: doc.add_paragraph().add_run(f"{val} : {text}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# 6. الواجهة الرئيسية واستدعاء المحركات
# -----------------------------------------------------------------------------
st.markdown("<h3 style='text-align: right;'>📂 منطقة الرفع والمطابقة</h3>", unsafe_allow_html=True)
uploaded_files = st.file_uploader("ارفع ملفي الشهر السابق والحالي معاً", type=['docx'], accept_multiple_files=True)

col_opts1, col_opts2, col_opts3 = st.columns(3)
with col_opts1:
    comparison_mode = st.radio("🎯 نوع المقارنة:", ["النوع الأول", "النوع الثاني", "النوع الثالث"], horizontal=True)
with col_opts2:
    card_choice_ui = st.radio("💳 البطاقة المعتمدة:", ["رقم البطاقة القديم", "رقم البطاقة الحديث"], horizontal=True)
with col_opts3:
    matching_engine = st.radio("⚙️ محرك المطابقة المستهدف:", ["المحرك القياسي", "محرك تخطي التسلسل (بطاقة فقط)"], horizontal=True)

card_type_param = "old" if card_choice_ui == "رقم البطاقة القديم" else "new"
card_col_name = card_choice_ui
swap_files = st.checkbox("🔄 **عكس الملفين يدوياً (القديم يصبح حديثاً والحديث قديماً)**")

grid_column_configuration = {
    "التسلسل": st.column_config.TextColumn("التسلسل", width="small"),
    "اسم رب الأسرة": st.column_config.TextColumn("اسم رب الأسرة", width="large"),
    card_col_name: st.column_config.TextColumn(card_col_name, width="medium"),
    "الحالة": st.column_config.TextColumn("الحالة", width="small"),
    "الأفراد الكلية": st.column_config.NumberColumn("الأفراد الكلية", width="small"),
    "الأفراد المستحقة": st.column_config.NumberColumn("الأفراد المستحقة", width="small"),
    "الأفراد المحجوبين": st.column_config.NumberColumn("الأفراد المحجوبين", width="small"),
    "الإحالة": st.column_config.TextColumn("الإحالة", width="large")
}

if st.button("بدء المقارنة الذكية واستخراج المتغيرات والأوراق"):
    if len(uploaded_files) == 2:
        with st.spinner('جاري التحليل وعزل الحالات تلقائياً...'):
            doc1, doc2 = Document(uploaded_files[0]), Document(uploaded_files[1])
            date1, date2 = extract_document_date(doc1), extract_document_date(doc2)
            
            file_a_is_older = (date1 < date2) if (date1 and date2) else True
            if swap_files: file_a_is_older = not file_a_is_older
                
            if file_a_is_older:
                old_doc, new_doc = doc1, doc2
                old_name, new_name = uploaded_files[0].name, uploaded_files[1].name
            else:
                old_doc, new_doc = doc2, doc1
                old_name, new_name = uploaded_files[1].name, uploaded_files[0].name

            st.markdown(f"<div class='date-badge'>الملف المعتمد كـ <span class='old'>السابق: ({old_name})</span> | الملف المعتمد كـ <span class='new'>الحديث: ({new_name})</span></div>", unsafe_allow_html=True)
            
            old_data = extract_clean_records(old_doc, card_type=card_type_param)
            new_data = extract_clean_records(new_doc, card_type=card_type_param)
            
            results, results_ref, counters = process_comparison(old_data, new_data, comparison_mode, card_col_name, matching_engine)
            
            if results:
                results = sorted(results, key=lambda x: (str(x.get("اسم رب الأسرة", "")), x.get("meta_sort", 0)))
                results_ref = sorted(results_ref, key=lambda x: str(x.get("اسم رب الأسرة", "")))
                
                df_results = pd.DataFrame(results)
                
                # الاحتفاظ بالنسخة الكاملة تماماً لملف الوورد لكي تعمل التصفية الفرعية بدقة هندسية
                df_results_full = df_results.copy()
                
                # تفريغ السطور للعرض المتناسق على الويب فقط منعا للازدواج البصري (للنوع الثاني)
                df_display = df_results.copy()
                if comparison_mode == "النوع الثاني":
                    for idx, row in df_display.iterrows():
                        if row.get("meta_status") == "type2_new":
                            df_display.at[idx, "التسلسل"], df_display.at[idx, "اسم رب الأسرة"], df_display.at[idx, card_col_name], df_display.at[idx, "الإحالة"] = "", "", "", ""
                
                st.markdown(f"<h3 style='text-align: right;'>📋 المخرجات ({comparison_mode}) - المعتمد: {matching_engine}</h3>", unsafe_allow_html=True)
                
                styled_df = df_display.style.apply(lambda d: style_all_types(d, old_data, new_data, card_col_name, comparison_mode), axis=None)
                
                if comparison_mode == "النوع الثاني": cols_order = ["التسلسل", "اسم رب الأسرة", card_col_name, "الحالة", "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين", "الإحالة"]
                else: cols_order = ["التسلسل", "اسم رب الأسرة", card_col_name, "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين", "الإحالة"]
                
                st.dataframe(styled_df, use_container_width=True, hide_index=True, column_order=cols_order, column_config=grid_column_configuration)
                
                base_name = new_name.rsplit('.', 1)[0]
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    # نمرر هنا الجدول الكامل غير المفرغ، ليقوم منشئ التقارير بتصفية كل حالة وعزلها في ورقة فرعية مستقلة بنجاح
                    word_report = create_word_table_report(df_results_full, f"تقرير - {comparison_mode}", comparison_mode, card_col_name, old_data, new_data, new_name)
                    st.download_button(label="📥 تحميل المخرجات Word (يحتوي الأوراق الفرعية)", data=word_report, file_name=f"تقرير_{base_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                with col_dl2:
                    word_stats = create_word_stats_report(counters, base_name)
                    st.download_button(label="📊 تحميل تقرير الإحصاء Word", data=word_stats, file_name=f"احصائيات_{base_name}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                
                st.markdown("---")
                st.markdown(f"<h3 style='text-align: right; color: #7F8C8D;'>📌 الجدول المرجعي الثابت</h3>", unsafe_allow_html=True)
                df_ref = pd.DataFrame(results_ref)
                cols_order_ref = ["التسلسل", "اسم رب الأسرة", card_col_name, "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين", "الإحالة"]
                st.dataframe(df_ref, use_container_width=True, hide_index=True, column_order=cols_order_ref, column_config=grid_column_configuration)
            else:
                st.success("🎉 تطابق تام! لا توجد فروقات بين الملفين.")
    else:
        st.warning("⚠️ يرجى رفع ملفين اثنين بالضبط للتمكن من بدء المقارنة.")
