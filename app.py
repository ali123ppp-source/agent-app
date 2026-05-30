import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import os
import re

# إعدادات الواجهة ودعم الاتجاه من اليمين إلى اليسار
st.set_page_config(page_title="نظام مقارنة بيانات الوكلاء المطور", layout="wide")
st.markdown("""
    <style>
    th, td { text-align: right !important; }
    div.stButton > button { background-color: #1A365D; color: white; width: 100%; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام مقارنة ملفات الوكلاء الذكي 🔎</h1>", unsafe_allow_html=True)

# التحقق من توفر مكتبات تصدير الـ PDF ودعم اللغة العربية
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.pagesizes import a4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    LIBS_READY = True
except ImportError:
    LIBS_READY = False

# -----------------------------------------------------------------------------
# محرّك الشفاء الإملائي الذكي وتطهير الأسماء (تم التحديث بناءً على عينة الملف المرفوع)
# -----------------------------------------------------------------------------
def fix_arabic_swaps(text):
    if not text: return ""
    
    # قاموس موسع جداً لإصلاح تشوهات الحروف المقلوبة والناقصة والمدمجة في لستات الوكلاء
    corrections = {
        # الأسماء الشائعة وحالات قلب الحروف
        "حنس": "حسن", "رايض": "رياض", "اثري": "اثير", "حسني": "حسين", "الحسني": "الحسين",
        "عيل": "علي", "عالء": "علاء", "امني": "أمين", "ازناد": "زناد", "مرتىض": "مرتضى",
        "رسهيد": "رشيد", "رسيح": "رسام", "شعالن": "شعلان", "منتضر": "منتظر", "برشى": "بشرى",
        "رسحان": "رسلان", "اكفائي": "الخفاجي", "العجييل": "العجيل", "الحمريي": "الحميري",
        "جارهللا": "جار الله", "عبدالحسني": "عبد الحسين", "حسنيه": "حسنية", "خالده": "خالدة",
        "كريمه": "كريمة", "مظلومه": "مظلومة", "حمزه": "حمزة", "زهره": "زهرة",
        
        # الكلمات المدمجة مع لفظ الجلالة أو الأسماء المركبة
        "هللا": "الله", "عبدالحسن": "عبد الحسين", "عبدالرحمن": "عبد الرحمن", 
        "عبدالرزاق": "عبد الرزاق", "عبدالامير": "عبد الأمير", "عبدالله": "عبد الله",
        
        # العشائر والألقاب الشائعة الممسوخة بسبب الـ PDF
        "الشبالوي": "الشبلاوي", "الشبلاوي": "الشبلاوي", "السلطاني": "السلطاني",
        "الدليمي": "الدليمي", "الجرواني": "الجرواني", "البعيجي": "البعيجي", 
        "المعموري": "المعموري", "اليساري": "اليساري", "المشايخي": "المشايخي"
    }
    
    # تنظيف مسبق للحروف الغريبة والرموز التي تظهر أحياناً وسط الأسماء
    text = re.sub(re.compile(r'[^\w\s\s]'), '', text)
    
    words = text.split()
    fixed_words = [corrections.get(w, w) for w in words]
    
    # معالجة المسافات الزائدة وإعادة دمج النص بشكل نظيف
    cleaned_text = " ".join(fixed_words)
    
    # إصلاح الأخطاء المركبة الملتصقة التي لم تعالج في فصل الكلمات
    cleaned_text = cleaned_text.replace("جارهللا", "جار الله")
    cleaned_text = cleaned_text.replace("عبدالحسني", "عبد الحسين")
    
    return cleaned_text

def fix_arabic_text(text):
    if not text: return ""
    try:
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except Exception:
        return str(text)

# -----------------------------------------------------------------------------
# محرك الاستخراج الدقيق مع المستشعر الذكي للتسلسل
# -----------------------------------------------------------------------------
def extract_data_from_pdf(file_obj, fix_reversed_arabic=True):
    records = {}
    
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    cells = []
                    for cell in row:
                        if cell:
                            val = str(cell).strip().replace('\n', ' ')
                            if fix_reversed_arabic and any('\u0600' <= char <= '\u06FF' for char in val):
                                # عكس النص إذا كان مقلوباً وتمريره لمعالج الفرز والشفاء الإملائي
                                val = val[::-1]
                                val = fix_arabic_swaps(val)
                            cells.append(val)
                        else:
                            cells.append("")
                    
                    if len(cells) < 6 or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "رقم البطاقة" in "".join(cells):
                        continue
                    
                    try:
                        # 1. استخراج التسلسل (ت) بشكل عكسي ذكي
                        seq = "-"
                        for i in range(len(cells)-1, -1, -1):
                            if cells[i].isdigit() and 0 < len(cells[i]) <= 4:
                                seq = cells[i]
                                break
                                
                        # 2. استخراج رقم البطاقة
                        card_num = ""
                        if len(cells) > 6 and cells[6].isdigit() and len(cells[6]) >= 5:
                            card_num = cells[6]
                        else:
                            digits = [c for c in cells if c.isdigit() and len(c) >= 5]
                            if digits: card_num = digits[0]
                            
                        if not card_num:
                            continue
                            
                        # 3. استخراج الاسم وتمريره لفلتر التطهير النهائي
                        name = ""
                        if len(cells) > 4 and not cells[4].isdigit() and len(cells[4]) > 4:
                            name = fix_arabic_swaps(cells[4])
                        else:
                            non_digits = [c for c in cells if not c.isdigit() and len(c) > 3]
                            if non_digits: name = fix_arabic_swaps(non_digits[0])

                        # 4. الحقول الرقمية
                        eligible = int(cells[1]) if len(cells) > 1 and cells[1].isdigit() else 0
                        withheld = int(cells[0]) if len(cells) > 0 and cells[0].isdigit() else 0
                        total = int(cells[2]) if len(cells) > 2 and cells[2].isdigit() else 0
                        
                        records[card_num] = {
                            "seq": seq,
                            "name": name,
                            "total": total,
                            "eligible": eligible,
                            "withheld": withheld
                        }
                    except Exception:
                        continue
    return records

# -----------------------------------------------------------------------------
# محرك المقارنة 
# -----------------------------------------------------------------------------
def compare_records(old_data, new_data):
    rows_output = []
    
    for card_num, old_val in old_data.items():
        if card_num in new_data:
            new_val = new_data[card_num]
            changes = []
            
            if old_val['eligible'] != new_val['eligible']:
                changes.append(f"المستحقة: من ({old_val['eligible']}) إلى ({new_val['eligible']})")
            if old_val['withheld'] != new_val['withheld']:
                changes.append(f"المحجوبة: من ({old_val['withheld']}) إلى ({new_val['withheld']})")
            if old_val['name'] != new_val['name']:
                changes.append(f"تعديل اسم: من [{old_val['name']}] إلى [{new_val['name']}]")
                
            if changes:
                rows_output.append({
                    "التسلسل الأصلي": new_val['seq'],
                    "رقم البطاقة": card_num,
                    "اسم رب الأسرة": new_val['name'],
                    "نوع التغيير": "تعديل في البيانات",
                    "تفاصيل المتغيرات": " | ".join(changes)
                })
            del new_data[card_num]
        else:
            rows_output.append({
                "التسلسل الأصلي": old_val['seq'],
                "رقم البطاقة": card_num,
                "اسم رب الأسرة": old_val['name'],
                "نوع التغيير": "محذوف / منقول",
                "تفاصيل المتغيرات": "تم رفع العائلة من قائمة الوكيل في الشهر الجديد"
            })

    for card_num, new_val in new_data.items():
        rows_output.append({
            "التسلسل الأصلي": new_val['seq'],
            "رقم البطاقة": card_num,
            "اسم رب الأسرة": new_val['name'],
            "نوع التغيير": "مضاف حديثا",
            "تفاصيل المتغيرات": "عائلة جديدة تم إنزالها لدى الوكيل"
        })
        
    return rows_output

# -----------------------------------------------------------------------------
# دالة توليد ملف الـ PDF بالعنوان الديناميكي
# -----------------------------------------------------------------------------
def generate_pdf_report(df_results, title_text):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=a4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    story = []
    
    font_path = "C:\\Windows\\Fonts\\arial.ttf"
    if not os.path.exists(font_path): font_path = "arial.ttf"
        
    try: pdfmetrics.registerFont(TTFont('ArabicArial', font_path))
    except Exception: pass
        
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'TitleStyle', parent=styles['Heading1'], fontName='ArabicArial', fontSize=14,
        alignment=2, textColor=colors.HexColor('#1A365D'), spaceAfter=25
    )
    
    original_cols = list(df_results.columns)
    reversed_cols = original_cols[::-1]
    
    table_data = [[fix_arabic_text(col) for col in reversed_cols]]
    
    for _, row in df_results.iterrows():
        row_cells = []
        for col in reversed_cols:
            val = str(row[col])
            if col == "نوع التغيير":
                if "تعديل" in val: val = "تعديل في البيانات"
                elif "محذوف" in val: val = "محذوف / منقول"
                elif "مضاف" in val: val = "مضاف حديثاً"
            row_cells.append(fix_arabic_text(val))
        table_data.append(row_cells)
        
    col_widths = [200, 95, 135, 75, 40]
    
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A365D')),
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'ArabicArial'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
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
st.markdown("<h4 style='text-align: right;'>⚙️ خيارات المعالجة والشفاء الإملائي المتقدم للأسماء</h4>", unsafe_allow_html=True)
fix_reversed = st.checkbox("🔄 تفعيل مصحح ومطهر الأسماء الذكي التلقائي (يوصى به لملفات الـ PDF المخربطة)", value=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("<h3 style='text-align: right;'>📂 ملف الشهر الجديد (الحديث)</h3>", unsafe_allow_html=True)
    new_file = st.file_uploader("ارفع ملف PDF الأحدث", type=['pdf'], key="new", label_visibility="collapsed")

with col2:
    st.markdown("<h3 style='text-align: right;'>📂 ملف الشهر القديم (السابق)</h3>", unsafe_allow_html=True)
    old_file = st.file_uploader("ارفع ملف PDF القديم", type=['pdf'], key="old", label_visibility="collapsed")

st.markdown("<br>", unsafe_allow_html=True)

if st.button("شغل المحرك وابدأ المقارنة وتطهير الأسماء الممسوخة الآن"):
    if old_file and new_file:
        with st.spinner('جاري قراءة الجداول، تصحيح تشوهات الأسماء وتثبيت التسلسل الأصلي...'):
            old_extracted = extract_data_from_pdf(old_file, fix_reversed_arabic=fix_reversed)
            new_extracted = extract_data_from_pdf(new_file, fix_reversed_arabic=fix_reversed)
            
            results = compare_records(old_extracted, new_extracted)
            
            if results:
                df_results = pd.DataFrame(results)
                
                total_mod = len(df_results[df_results["نوع التغيير"] == "تعديل في البيانات"])
                total_del = len(df_results[df_results["نوع التغيير"] == "محذوف / منقول"])
                total_new = len(df_results[df_results["نوع التغيير"] == "مضاف حديثا"])
                
                st.markdown("<h3 style='text-align: right;'>📊 ملخص الفروقات المكتشفة</h3>", unsafe_allow_html=True)
                c_new, c_mod, c_del = st.columns(3)
                c_new.metric("عوائل مضافة جديدة", total_new)
                c_mod.metric("عوائل تم تعديل أفرادها", total_mod)
                c_del.metric("عوائل تم نقلها/حذفها", total_del)
                
                dynamic_title = f"جدول الفروقات التفصيلي بين لستة ({old_file.name}) و لستة ({new_file.name})"
                st.markdown(f"<h3 style='text-align: right; color: #1A365D; border-bottom: 2px solid #1A365D; padding-bottom: 8px;'>📋 {dynamic_title}</h3>", unsafe_allow_html=True)
                
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
                        label="📥 تحميل تقرير الفروقات النهائي كملف PDF للطباعة",
                        data=pdf_data,
                        file_name=f"فروقات_{new_file.name}.pdf",
                        mime="application/pdf",
                    )
                else:
                    st.error("لا يمكن تحميل الـ PDF لعدم تثبيت المكتبات الداعمة.")
            else:
                st.success("🎉 تطابق تام! لم يتم العثور على أي تغيير أو تعديل بين القائمتين.")
    else:
        st.error("الرجاء التأكد من رفع كلا الملفين (القديم والجديد) لتتمكن من المقارنة.")
