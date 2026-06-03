import streamlit as st
import pandas as pd
from io import BytesIO
import os

# إعدادات الواجهة ودعم الاتجاه العربي
st.set_page_config(page_title="نظام مقارنة بيانات الوكلاء الدقيق", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #1A365D; color: white; width: 100%; font-weight: bold; }
    .report-box { background-color: #F0F4F8; padding: 15px; border-radius: 8px; border-right: 5px solid #1A365D; text-align: right; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام مقارنة ملفات الوكلاء الذكي والمطوّر 📄🔎</h1>", unsafe_allow_html=True)

# التحقق من توفر مكتبات تصدير الـ PDF ودعم اللغة العربية
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
# محرك الاستخراج الذكي جداً (مقاوم لاختلاف دمج الخلايا والأعمدة)
# -----------------------------------------------------------------------------
def extract_clean_records(file_obj):
    doc = Document(file_obj)
    records = {}
    
    for table in doc.tables:
        for row in table.rows:
            # استخراج النصوص وتنظيفها من الفراغات والأصناف الزائدة
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            
            # تخطي أسطر العناوين الرئيسية للمركز أو اسم الوكيل
            if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "اسم رب" in "".join(cells):
                continue
            
            # 1. التقاط اسم رب الأسرة (أطول نص عربي في السطر لا يحتوي على أرقام)
            name_idx = -1
            max_len = 0
            for i, c in enumerate(cells):
                if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
                    if len(c) > max_len:
                        max_len = len(c)
                        name_idx = i
            
            if name_idx == -1:
                continue # سطر غير صالح أو فارغ من الأسماء
            
            # 2. التقاط أرقام البطاقات (الخلايا الرقمية التي طولها 5 أرقام فما فوق)
            card_indices = [i for i, c in enumerate(cells) if c.isdigit() and len(c) >= 5]
            if not card_indices:
                continue
            
            # رقم البطاقة الرئيسي يكون دائماً الرقم الثاني أو الأخير قبل التسلسل
            if len(card_indices) >= 2:
                card_num = cells[card_indices[1]]
            else:
                card_num = cells[card_indices[0]]
                
            # 3. التقاط التسلسل (ت) يكون رقماً يقع بعد عمود البطاقة في نهاية السطر
            seq = "-"
            for i in range(len(cells)-1, card_indices[-1], -1):
                if cells[i].isdigit():
                    seq = cells[i]
                    break
            
            # 4. التقاط الحقول الرقمية الحسابية (المحجوبين، المستحقة، الكلية) قبل الاسم
            # في الملفات الأصلية تظهر بالترتيب: المحجوبين ثم المستحقة ثم الكلية
            digit_cells = [int(cells[i]) for i in range(name_idx) if cells[i].isdigit()]
            
            if len(digit_cells) >= 3:
                withheld = digit_cells[0]
                eligible = digit_cells[1]
                total = digit_cells[2]
            elif len(digit_cells) == 2:
                withheld = 0
                eligible = digit_cells[0]
                total = digit_cells[1]
            else:
                continue # قيد ناقص الأرقام الحسابية
                
            records[card_num] = {
                "seq": seq,
                "name": cells[name_idx],
                "total": total,
                "eligible": eligible,
                "withheld": withheld
            }
    return records

# -----------------------------------------------------------------------------
# محرك المقارنة وحساب التغيرات الدقيقة لكل حقل
# -----------------------------------------------------------------------------
def compare_records(old_data, new_data):
    results = []
    counters = {"name": 0, "total": 0, "eligible": 0, "withheld": 0}
    
    all_cards = set(old_data.keys()).union(set(new_data.keys()))
    
    for card_num in all_cards:
        # حالة 1: القيد موجود في الملفين (فحص التعديلات الداخلية)
        if card_num in old_data and card_num in new_data:
            old_val = old_data[card_num]
            new_val = new_data[card_num]
            
            name_chg = old_val["name"] != new_val["name"]
            total_chg = old_val["total"] != new_val["total"]
            elig_chg = old_val["eligible"] != new_val["eligible"]
            with_chg = old_val["withheld"] != new_val["withheld"]
            
            if name_chg or total_chg or elig_chg or with_chg:
                if name_chg: counters["name"] += 1
                if total_chg: counters["total"] += 1
                if elig_chg: counters["eligible"] += 1
                if with_chg: counters["withheld"] += 1
                
                results.append({
                    "التسلسل": new_val["seq"],
                    "رقم البطاقة": card_num,
                    "اسم رب الأسرة": new_val["name"],
                    "الأفراد الكلية": new_val["total"],
                    "الأفراد المستحقة": new_val["eligible"],
                    "الأفراد المحجوبين": new_val["withheld"]
                })
                
        # حالة 2: قيد محذوف أو منقول (موجود في القديم ومرفوع من الجديد)
        elif card_num in old_data:
            old_val = old_data[card_num]
            counters["name"] += 1
            counters["total"] += 1
            counters["eligible"] += 1
            counters["withheld"] += 1
            
            results.append({
                "التسلسل": old_val["seq"],
                "رقم البطاقة": card_num,
                "اسم رب الأسرة": old_val["name"] + " (تم حذفه/نقله)",
                "الأفراد الكلية": 0,
                "الأفراد المستحقة": 0,
                "الأفراد المحجوبين": 0
            })
            
        # حالة 3: قيد مضاف حديثاً (موجود في الجديد فقط)
        elif card_num in new_data:
            new_val = new_data[card_num]
            counters["name"] += 1
            counters["total"] += 1
            counters["eligible"] += 1
            counters["withheld"] += 1
            
            results.append({
                "التسلسل": new_val["seq"],
                "رقم البطاقة": card_num,
                "اسم رب الأسرة": new_val["name"] + " (مضاف حديثاً)",
                "الأفراد الكلية": new_val["total"],
                "الأفراد المستحقة": new_val["eligible"],
                "الأفراد المحجوبين": new_val["withheld"]
            })
            
    return results, counters

# -----------------------------------------------------------------------------
# دالة توليد تقرير الـ PDF المطابق للجدول
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
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='ArabicArial', fontSize=16, alignment=2, textColor=colors.HexColor('#1A365D'), spaceAfter=20)
    cell_text_style = ParagraphStyle('CellTextStyle', fontName='ArabicArial', fontSize=10, alignment=2, textColor=colors.black, leading=12)
    header_text_style = ParagraphStyle('HeaderStyle', fontName='ArabicArial', fontSize=11, alignment=2, textColor=colors.white)
    
    original_cols = list(df_results.columns)
    reversed_cols = original_cols[::-1] # عكس الاتجاه للـ ReportLab
    
    table_data = [[Paragraph(fix_arabic_text(col), header_text_style) for col in reversed_cols]]
    
    for _, row in df_results.iterrows():
        row_cells = []
        for col in reversed_cols:
            row_cells.append(Paragraph(fix_arabic_text(str(row[col])), cell_text_style))
        table_data.append(row_cells)
        
    col_widths = [100, 100, 100, 260, 140, 80] # الأبعاد المتناسقة للأعمدة الستة
    
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
# واجهة المستخدم ونقاط الرفع والتنفيذ
# -----------------------------------------------------------------------------
col1, col2 = st.columns(2)

with col1:
    st.markdown("<h3 style='text-align: right;'>📂 ملف الشهر الجديد (الحديث)</h3>", unsafe_allow_html=True)
    new_file = st.file_uploader("ارفع ملف Word الأحدث", type=['docx'], key="new", label_visibility="collapsed")

with col2:
    st.markdown("<h3 style='text-align: right;'>📂 ملف الشهر القديم (السابق)</h3>", unsafe_allow_html=True)
    old_file = st.file_uploader("ارفع ملف Word القديم", type=['docx'], key="old", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("شغل المحرك وابدأ المقارنة المطلقة وحساب التغيرات الآن"):
    if old_file and new_file:
        with st.spinner('جاري قراءة الملفات وتحليل الحقول الحسابية والأسماء بدقة متناهية...'):
            try:
                old_data = extract_clean_records(old_file)
                new_data = extract_clean_records(new_file)
                
                results, report_counters = compare_records(old_data, new_data)
                
                if results:
                    # الترتيب الأبجدي الإجباري بناءً على اسم رب الأسرة
                    results = sorted(results, key=lambda x: str(x.get("اسم رب الأسرة", "")))
                    
                    # تحويل المخرجات إلى DataFrame بالترتيب الصحيح للأعمدة
                    headers_order = ["التسلسل", "رقم البطاقة", "اسم رب الأسرة", "الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين"]
                    df_results = pd.DataFrame(results)[headers_order]
                    
                    # عرض عنوان الجدول
                    dynamic_title = f"جدول الفروقات والمستجدات بين قائمة ({old_file.name}) و قائمة ({new_file.name})"
                    st.markdown(f"<h3 style='text-align: right; color: #1A365D; border-bottom: 2px solid #1A365D; padding-bottom: 8px;'>📋 {dynamic_title} (مرتب أبجدياً)</h3>", unsafe_allow_html=True)
                    
                    # عرض الجدول النظيف الخالي من حقول نوع التغيير
                    st.dataframe(df_results, use_container_width=True, hide_index=True)
                    
                    # ---------------------------------------------------------
                    # تقرير أسفل كل حقل بعدد التغييرات التي حدثت ضمنه
                    # ---------------------------------------------------------
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown("<h3 style='text-align: right;'>📊 تقرير إحصاء التغيرات المكتشفة لكل حقل</h3>", unsafe_allow_html=True)
                    
                    rep_col1, rep_col2, rep_col3, rep_col4 = st.columns(4)
                    
                    with rep_col1:
                        st.markdown(f"""<div class='report-box'>
                            <h5>حقل الأفراد المحجوبين</h5>
                            <h2 style='color:#E67E22;'>{report_counters['withheld']}</h2>
                            <p style='font-size:12px;color:#7F8C8D;'>تغيير في أعداد المحجوبين</p>
                        </div>""", unsafe_allow_html=True)
                        
                    with rep_col2:
                        st.markdown(f"""<div class='report-box'>
                            <h5>حقل الأفراد المستحقة</h5>
                            <h2 style='color:#2980B9;'>{report_counters['eligible']}</h2>
                            <p style='font-size:12px;color:#7F8C8D;'>تغيير في أعداد المستحقين</p>
                        </div>""", unsafe_allow_html=True)
                        
                    with rep_col3:
                        st.markdown(f"""<div class='report-box'>
                            <h5>حقل الأفراد الكلية</h5>
                            <h2 style='color:#27AE60;'>{report_counters['total']}</h2>
                            <p style='font-size:12px;color:#7F8C8D;'>تغيير في المجموع الكلي</p>
                        </div>""", unsafe_allow_html=True)
                        
                    with rep_col4:
                        st.markdown(f"""<div class='report-box'>
                            <h5>حقل اسم رب الأسرة</h5>
                            <h2 style='color:#8E44AD;'>{report_counters['name']}</h2>
                            <p style='font-size:12px;color:#7F8C8D;'>تعديل إملائي أو نقل/إضافة اسم</p>
                        </div>""", unsafe_allow_html=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    # زر تحميل الـ PDF
                    if LIBS_READY:
                        pdf_data = generate_pdf_report(df_results, f"تقرير الفروقات النهائي المعتمد لـ: {new_file.name}")
                        st.download_button(
                            label="📥 تحميل التقرير النهائي كملف PDF جاهز للطباعة والتوقيع",
                            data=pdf_data, file_name=f"تقرير_فروقات_{new_file.name}.pdf", mime="application/pdf",
                        )
                    else:
                        st.error("لا يمكن تحميل ملف الـ PDF لعدم توفر المكتبات الداعمة على الخادم.")
                else:
                    st.success("🎉 تطابق تام ومطلق! لم يسجل النظام أي تغيير في الأسماء أو الحقول الحسابية للأفراد.")
            except Exception as e:
                st.error(f"حدث خطأ أثناء معالجة الجداول البرمجية، يرجى مراجعة صياغة الملف. تفاصيل الإشكال التقني: {e}")
    else:
        st.error("الرجاء رفع كلا الملفين بصيغة Word (.docx) لتتمكن من تشغيل محرك المقارنة المطور.")
