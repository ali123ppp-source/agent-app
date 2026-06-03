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

st.markdown("<h1 style='text-align: right;'>نظام المقارنة الشامل والذكي 📄🔁</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: right;'>يقوم بمقارنة ملفين بغض النظر عن الأقدمية بالاعتماد الكلي على أرقام البطاقات.</p>", unsafe_allow_html=True)

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
# محرك المقارنة المتكافئ
# -----------------------------------------------------------------------------
def compare_symmetric(data1, data2):
    results = []
    counters = {"name": 0, "total": 0, "eligible": 0, "withheld": 0, "only_in_1": 0, "only_in_2": 0}
    
    all_cards = set(data1.keys()).union(set(data2.keys()))
    
    for card in all_cards:
        # موجود في كلا الملفين
        if card in data1 and card in data2:
            v1, v2 = data1[card], data2[card]
            
            diff_name = v1["name"] != v2["name"]
            diff_total = v1["total"] != v2["total"]
            diff_elig = v1["eligible"] != v2["eligible"]
            diff_with = v1["withheld"] != v2["withheld"]
            
            if diff_name or diff_total or diff_elig or diff_with:
                if diff_name: counters["name"] += 1
                if diff_total: counters["total"] += 1
                if diff_elig: counters["eligible"] += 1
                if diff_with: counters["withheld"] += 1
                
                results.append({
                    "التسلسل": v2["seq"],
                    "رقم البطاقة": card,
                    "اسم رب الأسرة": v2["name"],
                    "الأفراد الكلية": v2["total"],
                    "الأفراد المستحقة": v2["eligible"],
                    "الأفراد المحجوبين": v2["withheld"]
                })
                
        # موجود في الملف الأول فقط (إما تم حذفه في الثاني، أو مضاف في الأول)
        elif card in data1 and card not in data2:
            v1 = data1[card]
            counters["only_in_1"] += 1
            results.append({
                "التسلسل": v1["seq"],
                "رقم البطاقة": card,
                "اسم رب الأسرة": v1["name"] + " (موجود في الملف 1 فقط)",
                "الأفراد الكلية": v1["total"],
                "الأفراد المستحقة": v1["eligible"],
                "الأفراد المحجوبين": v1["withheld"]
            })
            
        # موجود في الملف الثاني فقط
        elif card not in data1 and card in data2:
            v2 = data2[card]
            counters["only_in_2"] += 1
            results.append({
                "التسلسل": v2["seq"],
                "رقم البطاقة": card,
                "اسم رب الأسرة": v2["name"] + " (موجود في الملف 2 فقط)",
                "الأفراد الكلية": v2["total"],
                "الأفراد المستحقة": v2["eligible"],
                "الأفراد المحجوبين": v2["withheld"]
            })
            
    return results, counters

# -----------------------------------------------------------------------------
# دوال إنشاء ملفات الـ Word
# -----------------------------------------------------------------------------
def create_word_table_report(df, title):
    doc = Document()
    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # عكس ترتيب الأعمدة لتظهر من اليمين لليسار في برنامج وورد
    cols = list(df.columns)[::-1]
    
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
    heading = doc.add_heading(f"تقرير إحصاء المتغيرات - {filename_base}", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph().add_run().add_break()
    
    stats_data = [
        ("عدد التغييرات في الأفراد الكلية:", counters['total']),
        ("عدد التغييرات في الأفراد المستحقة:", counters['eligible']),
        ("عدد التغييرات في الأفراد المحجوبين:", counters['withheld']),
        ("عدد التغييرات في أسماء أرباب الأسر:", counters['name']),
        ("عوائل موجودة في الملف الأول فقط:", counters['only_in_1']),
        ("عوائل موجودة في الملف الثاني فقط:", counters['only_in_2']),
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
    st.markdown("<h3 style='text-align: right;'>📂 الملف الأول</h3>", unsafe_allow_html=True)
    file1 = st.file_uploader("ارفع الملف الأول", type=['docx'], key="f1", label_visibility="collapsed")

with col2:
    st.markdown("<h3 style='text-align: right;'>📂 الملف الثاني</h3>", unsafe_allow_html=True)
    file2 = st.file_uploader("ارفع الملف الثاني", type=['docx'], key="f2", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("بدء المقارنة الشاملة"):
    if file1 and file2:
        with st.spinner('جاري إجراء المقارنة الشاملة وتجهيز ملفات الوورد...'):
            try:
                data1 = extract_clean_records(file1)
                data2 = extract_clean_records(file2)
                
                results, counters = compare_symmetric(data1, data2)
                
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
                    with c4: st.markdown(f"<div class='report-box'>إضافات/حذوفات<br><h2>{counters['only_in_1'] + counters['only_in_2']}</h2></div>", unsafe_allow_html=True)
                    
                    # استخراج اسم الملف الأساسي بدون الامتداد
                    base_name = file1.name.rsplit('.', 1)[0]
                    
                    # توليد أزرار التحميل
                    col_dl1, col_dl2 = st.columns(2)
                    with col_dl1:
                        word_report = create_word_table_report(df_results, f"المتغيرات - {base_name}")
                        st.download_button(
                            label="📥 تحميل جدول المتغيرات (Word)",
                            data=word_report,
                            file_name=f"{base_name}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                    with col_dl2:
                        word_stats = create_word_stats_report(counters, base_name)
                        st.download_button(
                            label="📊 تحميل تقرير الإحصاء (Word)",
                            data=word_stats,
                            file_name=f"{base_name} - احصاء.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        )
                else:
                    st.success("🎉 الملفان متطابقان تماماً! لا توجد أي إضافات، حذوفات، أو تغييرات في الحقول.")
            except Exception as e:
                st.error(f"خطأ غير متوقع أثناء المعالجة: {e}")
    else:
        st.warning("يرجى رفع كلا الملفين أولاً.")
