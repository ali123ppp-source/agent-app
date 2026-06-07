import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# إعدادات الواجهة
st.set_page_config(page_title="نظام المقارنة المطور للوكلاء", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2C3E50; color: white; width: 100%; font-weight: bold; border-radius: 8px;}
    .report-box { background-color: #ECF0F1; padding: 15px; border-radius: 8px; border-right: 5px solid #2C3E50; text-align: right; margin-bottom: 10px;}
    .net-diff { font-size: 16px; font-weight: bold; margin-top: 5px; color: #2C3E50; border-top: 1px solid #BDC3C7; padding-top: 5px;}
    .stat-inc { font-size: 14px; color: #27AE60; font-weight: bold; }
    .stat-dec { font-size: 14px; color: #C0392B; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام المقارنة الشامل والذكي (نسخة كشف الفروقات المقارنة) 📄🔎</h1>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# محرك الاستخراج الدقيق
# -----------------------------------------------------------------------------
def extract_clean_records(file_obj):
    doc = Document(file_obj)
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
# محرك المقارنة الثنائي (يكشف القيم السابقة والحالية)
# -----------------------------------------------------------------------------
def compare_records(old_data, new_data):
    results = []
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
            
            # تم الإبقاء على مقارنة الأرقام فقط، ولكن سنعرض المقارنة الكاملة في الجدول لكشف اللبس
            if diff_total or diff_elig or diff_with:
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
                
                results.append({
                    "التسلسل": new_v["seq"],
                    "رقم البطاقة": card,
                    "الاسم (سابقاً)": old_v["name"],
                    "الاسم (حالياً)": new_v["name"],
                    "الكلية (سابقاً)": old_v["total"],
                    "الكلية (حالياً)": new_v["total"],
                    "المستحقة (سابقاً)": old_v["eligible"],
                    "المستحقة (حالياً)": new_v["eligible"],
                    "المحجوبين (سابقاً)": old_v["withheld"],
                    "المحجوبين (حالياً)": new_v["withheld"]
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
            
            results.append({
                "التسلسل": old_v["seq"],
                "رقم البطاقة": card,
                "الاسم (سابقاً)": old_v["name"],
                "الاسم (حالياً)": "❌ (محذوف / منقول)",
                "الكلية (سابقاً)": old_v["total"], "الكلية (حالياً)": 0,
                "المستحقة (سابقاً)": old_v["eligible"], "المستحقة (حالياً)": 0,
                "المحجوبين (سابقاً)": old_v["withheld"], "المحجوبين (حالياً)": 0
            })
            
        elif card not in old_data and card in new_data:
            new_v = new_data[card]
            counters["added_fam"] += 1
            counters["inc_total"] += new_v["total"]
            counters["net_total"] += new_v["total"]
            counters["inc_eligible"] += new_v["eligible"]
            counters["net_eligible"] += new_v["eligible"]
            counters["inc_withheld"] += new_v["withheld"]
            counters["net_withheld"] += new_v["withheld"]
            
            results.append({
                "التسلسل": new_v["seq"],
                "رقم البطاقة": card,
                "الاسم (سابقاً)": "✨ (مضاف حديثاً)",
                "الاسم (حالياً)": new_v["name"],
                "الكلية (سابقاً)": 0, "الكلية (حالياً)": new_v["total"],
                "المستحقة (سابقاً)": 0, "المستحقة (حالياً)": new_v["eligible"],
                "المحجوبين (سابقاً)": 0, "المحجوبين (حالياً)": new_v["withheld"]
            })
            
    return results, counters

# -----------------------------------------------------------------------------
# دوال إنشاء ملفات الـ Word
# -----------------------------------------------------------------------------
def create_word_table_report(df, title):
    doc = Document()
    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cols = list(df.columns)[::-1]
    
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(cols):
        hdr_cells[i].text = str(col)
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col in enumerate(cols):
            row_cells[i].text = str(row[col])
            row_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def create_word_stats_report(counters, filename_base):
    doc = Document()
    heading = doc.add_heading(f"تقرير الإحصاء لشهر - {filename_base}", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph().add_run().add_break()
    
    p_title1 = doc.add_paragraph()
    p_title1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_title1.add_run("أولاً: إحصاء حركة الأفراد").bold = True
    
    stats_individuals = [
        ("إجمالي زيادة الأفراد الكلية (المضافين):", f"+{counters['inc_total']}"),
        ("إجمالي نقصان الأفراد الكلية (المحذوفين):", f"-{counters['dec_total']}"),
        ("صافي التغيير في الأفراد الكلية:", f"{counters['net_total']:+d}"),
        ("---", ""),
        ("إجمالي زيادة الأفراد المستحقة:", f"+{counters['inc_eligible']}"),
        ("إجمالي نقصان الأفراد المستحقة:", f"-{counters['dec_eligible']}"),
        ("صافي التغيير في الأفراد المستحقة:", f"{counters['net_eligible']:+d}")
    ]
    for text, val in stats_individuals:
        if text == "---":
            doc.add_paragraph()
            continue
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(f" {val} ").bold = True
        p.add_run(f" : {text}")
        
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# واجهة الاستخدام
# -----------------------------------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    st.markdown("<h3 style='text-align: right;'>📂 ملف الشهر الجديد (الحالي)</h3>", unsafe_allow_html=True)
    new_file = st.file_uploader("ارفع الملف الجديد", type=['docx'], key="new", label_visibility="collapsed")

with col2:
    st.markdown("<h3 style='text-align: right;'>📂 ملف الشهر القديم (السابق)</h3>", unsafe_allow_html=True)
    old_file = st.file_uploader("ارفع الملف القديم", type=['docx'], key="old", label_visibility="collapsed")

if st.button("بدء المقارنة الدقيقة واستخراج المتغيرات"):
    if old_file and new_file:
        with st.spinner('جاري جلب ومطابقة القيود الرقمية...'):
            try:
                old_data = extract_clean_records(old_file)
                new_data = extract_clean_records(new_file)
                
                results, counters = compare_records(old_data, new_data)
                
                if results:
                    # ترتيب النتائج
                    results = sorted(results, key=lambda x: str(x.get("الاسم (حالياً)", "")))
                    df_results = pd.DataFrame(results)[["التسلسل", "رقم البطاقة", "الاسم (سابقاً)", "الاسم (حالياً)", "الكلية (سابقاً)", "الكلية (حالياً)", "المستحقة (سابقاً)", "المستحقة (حالياً)", "المحجوبين (سابقاً)", "المحجوبين (حالياً)"]]
                    
                    st.markdown("<h3 style='text-align: right; color: #2C3E50;'>📋 جدول المتغيرات المقارن</h3>", unsafe_allow_html=True)
                    st.dataframe(df_results, use_container_width=True, hide_index=True)
                    
                    # صناديق الإحصائيات
                    st.markdown("<h3 style='text-align: right; margin-top: 20px;'>📊 إحصائية الفروقات</h3>", unsafe_allow_html=True)
                    c1, c2, c3 = st.columns(3)
                    with c1: 
                        st.markdown(f"<div class='report-box'>حركة الكلية<br><h2>{counters['total_fam']} عائلة</h2>الصافي النهائي: {counters['net_total']:+d}</div>", unsafe_allow_html=True)
                    with c2: 
                        st.markdown(f"<div class='report-box'>حركة المستحقة<br><h2>{counters['eligible_fam']} عائلة</h2>الصافي النهائي: {counters['net_eligible']:+d}</div>", unsafe_allow_html=True)
                    with c3: 
                        st.markdown(f"<div class='report-box'>عوائل مضافة/محذوفة<br><h2>{counters['added_fam'] + counters['deleted_fam']} عائلة</h2></div>", unsafe_allow_html=True)
                    
                    base_name = new_file.name.rsplit('.', 1)[0]
                    word_report = create_word_table_report(df_results, f"متغيرات - {base_name}")
                    st.download_button(
                        label="📥 تحميل جدول المتغيرات المطور (Word)",
                        data=word_report,
                        file_name=f"متغيرات_{base_name}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                else:
                    st.success("🎉 الملفان متطابقان تماماً رقمياً!")
            except Exception as e:
                st.error(f"خطأ: {e}")
