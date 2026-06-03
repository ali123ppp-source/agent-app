import streamlit as st
import pandas as pd
from io import BytesIO
import os

# إعدادات الواجهة
st.set_page_config(page_title="نظام مقارنة بيانات الوكلاء (نسخة الوورد)", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; }
    div.stButton > button { background-color: #1A365D; color: white; width: 100%; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام مقارنة ملفات الوكلاء الذكي (Word) 📄🔎</h1>", unsafe_allow_html=True)

# التحقق من توفر المكتبات
try:
    from docx import Document
    import arabic_reshaper
    from bidi.algorithm import get_display
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.pagesizes import a4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    LIBS_READY = True
except ImportError:
    LIBS_READY = False

def fix_arabic_text(text):
    if not text: return ""
    try:
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except Exception:
        return str(text)

# -----------------------------------------------------------------------------
# محرك استخراج البيانات من ملف الوورد (يحافظ على هيكل الجدول الأصلي)
# -----------------------------------------------------------------------------
def extract_data_from_word(file_obj):
    doc = Document(file_obj)
    records = {}
    headers = []
    
    id_col_name = ""   # اسم عمود رقم البطاقة
    name_col_name = "" # اسم عمود الاسم

    for table in doc.tables:
        for i, row in enumerate(table.rows):
            # استخراج النصوص من الخلايا وتنظيفها
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            
            # 1. البحث عن صف العناوين (الهيدر)
            if not headers and any("بطاقة" in c for c in cells):
                headers = cells
                # تحديد أسماء الأعمدة الهامة للمقارنة والترتيب
                for h in headers:
                    if "بطاقة" in h: id_col_name = h
                    if "اسم" in h or "الاسم" in h: name_col_name = h
                continue

            # 2. استخراج البيانات وربطها بالعناوين
            if headers and len(cells) == len(headers):
                # تخطي الصفوف الفارغة أو صفوف العناوين المكررة
                if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells):
                    continue
                
                # إنشاء قاموس يمثل الصف، وربط كل قيمة باسم العمود الخاص بها
                row_data = {headers[j]: cells[j] for j in range(len(headers))}
                
                # استخدام رقم البطاقة كمفتاح أساسي وفريد
                if id_col_name and row_data.get(id_col_name) and row_data[id_col_name].isdigit():
                    card_num = row_data[id_col_name]
                    records[card_num] = row_data
                    
    return records, headers, name_col_name

# -----------------------------------------------------------------------------
# محرك المقارنة 
# -----------------------------------------------------------------------------
def compare_word_records(old_data, new_data, headers):
    results = []
    
    # 1. البحث عن المحذوفين والمعدلين
    for card_num, old_row in old_data.items():
        if card_num in new_data:
            new_row = new_data[card_num]
            changes = []
            
            # مقارنة كل حقل بحقله (باستثناء التسلسل لأنه يتغير طبيعياً)
            for h in headers:
                if "تسلسل" not in h and "ت" != h.strip():
                    val_old = old_row.get(h, "")
                    val_new = new_row.get(h, "")
                    if val_old != val_new:
                        changes.append(f"({h}): من [{val_old}] إلى [{val_new}]")
            
            if changes:
                result_row = new_row.copy()
                result_row["نوع التغيير"] = "تعديل في البيانات"
                result_row["تفاصيل المتغيرات"] = " | ".join(changes)
                results.append(result_row)
            
            del new_data[card_num] # إزالة المطابق والمعدل من القائمة الجديدة
        else:
            result_row = old_row.copy()
            result_row["نوع التغيير"] = "محذوف / منقول"
            result_row["تفاصيل المتغيرات"] = "تم رفع العائلة من قائمة الوكيل في الشهر الجديد"
            results.append(result_row)

    # 2. الباقي في القائمة الجديدة هم المضافون حديثاً
    for card_num, new_row in new_data.items():
        result_row = new_row.copy()
        result_row["نوع التغيير"] = "مضاف حديثا"
        result_row["تفاصيل المتغيرات"] = "عائلة جديدة تم إنزالها لدى الوكيل"
        results.append(result_row)
        
    return results

# -----------------------------------------------------------------------------
# دالة توليد ملف الـ PDF 
# -----------------------------------------------------------------------------
def generate_pdf_report(df_results, title_text):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(a4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []
    
    font_path = "C:\\Windows\\Fonts\\arial.ttf"
    if not os.path.exists(font_path): font_path = "arial.ttf"
        
    try: pdfmetrics.registerFont(TTFont('ArabicArial', font_path))
    except Exception: pass
        
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontName='ArabicArial', fontSize=16,
        alignment=2, textColor=colors.HexColor('#1A365D'), spaceAfter=20
    )
    cell_text_style = ParagraphStyle('CellTextStyle', fontName='ArabicArial', fontSize=9, alignment=2, textColor=colors.black, leading=12)
    header_text_style = ParagraphStyle('HeaderStyle', fontName='ArabicArial', fontSize=10, alignment=2, textColor=colors.white, fontStyle='Bold')
    
    original_cols = list(df_results.columns)
    reversed_cols = original_cols[::-1]
    
    table_data = [[Paragraph(fix_arabic_text(col), header_text_style) for col in reversed_cols]]
    
    for _, row in df_results.iterrows():
        row_cells = []
        for col in reversed_cols:
            val = str(row[col])
            if col == "نوع التغيير":
                if "تعديل" in val: val = "تعديل في البيانات"
                elif "محذوف" in val: val = "محذوف / منقول"
                elif "مضاف" in val: val = "مضاف حديثاً"
            row_cells.append(Paragraph(fix_arabic_text(val), cell_text_style))
        table_data.append(row_cells)
        
    # حساب العرض التقريبي للأعمدة بناءً على عددها
    col_width = 800 / len(reversed_cols) if reversed_cols else 100
    col_widths = [col_width] * len(reversed_cols)
    
    # تكبير عمود التفاصيل والاسم إذا كانا موجودين
    for i, col in enumerate(reversed_cols):
        if "تفاصيل" in col: col_widths[i] = 200
        elif "اسم" in col or "الاسم" in col: col_widths[i] = 150
    
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A365D')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8F9F9')])
    ]))
    
    story.append(Paragraph(fix_arabic_text(title_text), title_style))
    story.append(Spacer(1, 10))
    story.append(t)
    doc.build(story)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# واجهة المستخدم 
# -----------------------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    st.markdown("<h3 style='text-align: right;'>📂 ملف الشهر الجديد (الحديث)</h3>", unsafe_allow_html=True)
    new_file = st.file_uploader("ارفع ملف Word الأحدث", type=['docx'], key="new", label_visibility="collapsed")

with col2:
    st.markdown("<h3 style='text-align: right;'>📂 ملف الشهر القديم (السابق)</h3>", unsafe_allow_html=True)
    old_file = st.file_uploader("ارفع ملف Word القديم", type=['docx'], key="old", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("شغل المحرك وابدأ المقارنة الدقيقة للوورد الآن"):
    if old_file and new_file:
        with st.spinner('جاري قراءة ملفات الوورد واستخراج الجداول والفروقات...'):
            try:
                # الاستخراج
                old_extracted, old_headers, name_col = extract_data_from_word(old_file)
                new_extracted, new_headers, _ = extract_data_from_word(new_file)
                
                # استخدام عناوين الملف الجديد كأساس، وإضافة عمودي التغييرات
                final_headers = new_headers + ["نوع التغيير", "تفاصيل المتغيرات"]
                
                # المقارنة
                results = compare_word_records(old_extracted, new_extracted, new_headers)
                
                if results:
                    # الترتيب الأبجدي بناءً على عمود الاسم
                    if name_col:
                        results = sorted(results, key=lambda x: str(x.get(name_col, "")))
                    
                    # ترتيب الأعمدة لتتطابق مع هيكل الملف الأصلي
                    df_results = pd.DataFrame(results)[final_headers]
                    
                    total_mod = len(df_results[df_results["نوع التغيير"] == "تعديل في البيانات"])
                    total_del = len(df_results[df_results["نوع التغيير"] == "محذوف / منقول"])
                    total_new = len(df_results[df_results["نوع التغيير"] == "مضاف حديثا"])
                    
                    st.markdown("<h3 style='text-align: right;'>📊 ملخص الفروقات المكتشفة</h3>", unsafe_allow_html=True)
                    c_new, c_mod, c_del = st.columns(3)
                    c_new.metric("عوائل مضافة جديدة", total_new)
                    c_mod.metric("عوائل تم تعديل أفرادها", total_mod)
                    c_del.metric("عوائل تم نقلها/حذفها", total_del)
                    
                    dynamic_title = f"جدول الفروقات التفصيلي بين ({old_file.name}) و ({new_file.name})"
                    st.markdown(f"<h3 style='text-align: right; color: #1A365D; border-bottom: 2px solid #1A365D; padding-bottom: 8px;'>📋 {dynamic_title} (مرتب أبجدياً)</h3>", unsafe_allow_html=True)
                    
                    df_display = df_results.copy()
                    df_display["نوع التغيير"] = df_display["نوع التغيير"].map({
                        "تعديل في البيانات": "🟡 تعديل في البيانات",
                        "محذوف / منقول": "🔴 محذوف / منقول",
                        "مضاف حديثا": "🟢 مضاف حديثاً"
                    })
                    st.dataframe(df_display, use_container_width=True, hide_index=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    if LIBS_READY:
                        pdf_data = generate_pdf_report(df_results, f"تقرير الفروقات النهائي بين: {old_file.name} و {new_file.name}")
                        st.download_button(
                            label="📥 تحميل تقرير الفروقات النهائي كملف PDF",
                            data=pdf_data, file_name=f"فروقات_{new_file.name}.pdf", mime="application/pdf",
                        )
                    else:
                        st.error("لا يمكن تحميل الـ PDF لعدم تثبيت المكتبات الداعمة.")
                else:
                    st.success("🎉 تطابق تام! لم يتم العثور على أي تغيير أو تعديل بين القائمتين.")
            except Exception as e:
                st.error(f"حدث خطأ أثناء قراءة ملفات الوورد، يرجى التأكد من أن الملفات تحتوي على جداول نظامية. تفاصيل الخطأ: {e}")
    else:
        st.error("الرجاء التأكد من رفع كلا الملفين (القديم والجديد) بصيغة Word لتتمكن من المقارنة.")
