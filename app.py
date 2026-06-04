import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# إعدادات الواجهة
st.set_page_config(page_title="نظام المقارنة الشامل للوكلاء", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2C3E50; color: white; width: 100%; font-weight: bold; border-radius: 8px;}
    .report-box { background-color: #ECF0F1; padding: 15px; border-radius: 8px; border-right: 5px solid #2C3E50; text-align: right; margin-bottom: 10px;}
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام المقارنة الشامل والذكي 📄🔎</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>يقوم بمقارنة ملف الشهر الجديد بناءً على ملف الشهر القديم لاستخراج المتغيرات الدقيقة.</p>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# محرك الاستخراج الدقيق
# -----------------------------------------------------------------------------
def extract_clean_records(file_obj):
    doc = Document(file_obj)
    records = {}
    
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            
            # تخطي أسطر العناوين
            if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "اسم رب" in "".join(cells):
                continue
            
            # تحديد حقل الاسم
            name_idx = -1
            max_len = 0
            for i, c in enumerate(cells):
                if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
                    if len(c) > max_len:
                        max_len = len(c)
                        name_idx = i
            
            if name_idx == -1: continue
            
            # تحديد أرقام البطاقات (أرقام تتجاوز 5 خانات)
            card_indices = [i for i, c in enumerate(cells) if c.isdigit() and len(c) >= 5]
            if not card_indices: continue
            
            card_num = cells[card_indices[1]] if len(card_indices) >= 2 else cells[card_indices[0]]
                
            # التقاط التسلسل
            seq = "-"
            for i in range(len(cells)-1, card_indices[-1], -1):
                if cells[i].isdigit():
                    seq = cells[i]
                    break
            
            # الحقول الحسابية
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
# محرك المقارنة (القديم مقابل الجديد)
# -----------------------------------------------------------------------------
def compare_records(old_data, new_data):
    results = []
    counters = {"name": 0, "total": 0, "eligible": 0, "withheld": 0, "added": 0, "deleted": 0}
    
    all_cards = set(old_data.keys()).union(set(new_data.keys()))
    
    for card in all_cards:
        # موجود في كلا الملفين (تعديل بيانات)
        if card in old_data and card in new_data:
            old_v, new_v = old_data[card], new_data[card]
            
            diff_name = old_v["name"] != new_v["name"]
            diff_total = old_v["total"] != new_v["total"]
            diff_elig = old_v["eligible"] != new_v["eligible"]
            diff_with = old_v["withheld"] != new_v["withheld"]
            
            if diff_name or diff_total or diff_elig or diff_with:
                if diff_name: counters["name"] += 1
                if diff_total: counters["total"] += 1
                if diff_elig: counters["eligible"] += 1
                if diff_with: counters["withheld"] += 1
                
                results.append({
                    "التسلسل": new_v["seq"],
                    "رقم البطاقة": card,
                    "اسم رب الأسرة": new_v["name"],
                    "الأفراد الكلية": new_v["total"],
                    "الأفراد المستحقة": new_v["eligible"],
                    "الأفراد المحجوبين": new_v["withheld"]
                })
                
        # موجود في القديم ومفقود في الجديد (محذوف / منقول)
        elif card in old_data and card not in new_data:
            old_v = old_data[card]
            counters["deleted"] += 1
            results.append({
                "التسلسل": old_v["seq"],
                "رقم البطاقة": card,
                "اسم رب الأسرة": old_v["name"] + " (محذوف / منقول)",
                "الأفراد الكلية": old_v["total"],
                "الأفراد المستحقة": old_v["eligible"],
                "الأفراد المحجوبين": old_v["withheld"]
            })
            
        # غير موجود في القديم وموجود في الجديد (مضاف حديثاً)
        elif card not in old_data and card in new_data:
            new_v = new_data[card]
            counters["added"] += 1
            results.append({
                "التسلسل": new_v["seq"],
                "رقم البطاقة": card,
                "اسم رب الأسرة": new_v["name"] + " (مضاف حديثاً)",
                "الأفراد الكلية": new_v["total"],
                "الأفراد المستحقة": new_v["eligible"],
                "الأفراد المحجوبين": new_v["withheld"]
            })
            
    return results, counters

# -----------------------------------------------------------------------------
# دوال إنشاء ملفات الـ Word
# -----------------------------------------------------------------------------
def create_word_table_report(df, title):
    doc = Document()
    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    cols = list(df.columns)[::-1] # عكس الترتيب ليتوافق مع اليمين لليسار في وورد
    
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # إضافة العناوين
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(cols):
        hdr_cells[i].text = str(col)
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    # إضافة البيانات
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
    heading = doc.add_heading(f"تقرير إحصاء المتغيرات لشهر - {filename_base}", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph().add_run().add_break()
    
    stats_data = [
        ("عدد التغييرات في الأفراد الكلية:", counters['total']),
        ("عدد التغييرات في الأفراد المستحقة:", counters['eligible']),
        ("عدد التغييرات في الأفراد المحجوبين:", counters['withheld']),
        ("عدد التغييرات في أسماء أرباب الأسر:", counters['name']),
        ("عوائل تم إضافتها حديثاً في هذا الشهر:", counters['added']),
        ("عوائل تم نقلها أو حذفها هذا الشهر:", counters['deleted']),
    ]
    
    for text, val in stats_data:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run(f"{val}").bold = True
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

st.markdown("<br>", unsafe_allow_html=True)

if st.button("بدء المقارنة الدقيقة واستخراج المتغيرات"):
    if old_file and new_file:
        with st.spinner('جاري قراءة الملفات ومطابقة القيود...'):
            try:
                old_data = extract_clean_records(old_file)
                new_data = extract_clean_records(new_file)
                
                results, counters = compare_records(old_data, new_data)
                
                if results:
                    results = sorted(results, key=lambda x: str(x.get("اسم رب الأسرة", "")))
                    df_results = pd.DataFrame(results)[["التسلسل", "رقم البطاقة", "اسم رب الأسرة", "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين"]]
                    
                    st.markdown("<h3 style='text-align: right; color: #2C3E50;'>📋 جدول المتغيرات</h3>", unsafe_allow_html=True)
                    st.dataframe(df_results, use_container_width=True, hide_index=True)
                    
                    st.markdown("<h3 style='text-align: right; margin-top: 20px;'>📊 إحصائية الفروقات</h3>", unsafe_allow_html=True)
                    c1, c2, c3, c4 = st.columns(4)
                    with c1: st.markdown(f"<div class='report-box'>محجوبين<br><h2>{counters['withheld']}</h2></div>", unsafe_allow_html=True)
                    with c2: st.markdown(f"<div class='report-box'>مستحقة<br><h2>{counters['eligible']}</h2></div>", unsafe_allow_html=True)
                    with c3: st.markdown(f"<div class='report-box'>الكلية<br><h2>{counters['total']}</h2></div>", unsafe_allow_html=True)
                    with c4: st.markdown(f"<div class='report-box'>إضافات/حذوفات<br><h2>{counters['added'] + counters['deleted']}</h2></div>", unsafe_allow_html=True)
                    
                    # استخراج اسم الملف الجديد ليكون اسم تقرير المتغيرات
                    base_name = new_file.name.rsplit('.', 1)[0]
                    
                    # توليد أزرار التحميل
                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        word_report = create_word_table_report(df_results, f"متغيرات شهر - {base_name}")
                        st.download_button(
                            label="📥 تحميل جدول المتغيرات (Word)",
                            data=word_report,
                            file_name=f"متغيرات_{base_name}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                    with col_dl2:
                        word_stats = create_word_stats_report(counters, base_name)
                        st.download_button(
                            label="📊 تحميل تقرير الإحصاء (Word)",
                            data=word_stats,
                            file_name=f"احصاء_{base_name}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                else:
                    st.success("🎉 الملفان متطابقان تماماً! لا توجد أي إضافات، حذوفات، أو تغييرات في الحقول هذا الشهر.")
            except Exception as e:
                st.error(f"خطأ غير متوقع أثناء المعالجة: {e}")
    else:
        st.warning("يرجى رفع كلا الملفين (القديم والجديد) أولاً.")
