import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
import os
import re

# إعدادات الواجهة
st.set_page_config(page_title="نظام مقارنة بيانات الوكلاء المطور", layout="wide")

st.markdown("""
    <style>
    th, td { text-align: right !important; }
    div.stButton > button { background-color: #1A365D; color: white; width: 100%; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام مقارنة ملفات الوكلاء الذكي 🔎</h1>", unsafe_allow_html=True)

try:
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

# -----------------------------------------------------------------------------
# محرك الشفاء الجذري والهيكلي للأسماء (يدعم القاموس الديناميكي والمستقبلي)
# -----------------------------------------------------------------------------
def fix_arabic_swaps(text, custom_dict=None):
    if not text: return ""
    
    # 1. تنظيف أولي للمسافات الزائدة
    text = " ".join(text.split())
    
    # 2. إصلاحات هيكلية للحروف الملتصقة والمقلوبة بسبب إحداثيات الـ PDF
    text = text.replace("هللا", "الله")
    text = text.replace("الء", "لاء")   # عالء -> علاء
    text = text.replace("عالن", "علان") # شعالن -> شعلان
    text = text.replace("الوي", "لاوي") # شبالوي -> شبلاوي، الفتالوي -> الفتلاوي
    text = text.replace("ىض", "ضى")     # مرتىض -> مرتضى
    text = text.replace("ىف", "فى")     # مصطىف -> مصطفى
    text = text.replace("ىد", "دى")     # هىد -> هدى
    text = text.replace("رشى", "شرى")   # برشى -> بشرى
    text = text.replace("سحني", "حسين") # إصلاح اسم حسين
    text = text.replace("شالل", "شلال") 
    text = text.replace("تحسني", "تحسين") 
    text = text.replace("امرية", "اميرة") 
    text = text.replace("امري", "امير") 
    text = text.replace("حنس", "حسن")   # إصلاح اسم حسن
    
    # 3. قاموس الكلمات القياسي المثبت مسبقاً للحالات الشائعة
    exact_matches = {
        "حسني": "حسين", "الحسني": "الحسين",
        "عبدالحسني": "عبد الحسين", "عبدالحسن": "عبد الحسين",
        "عيل": "علي", "العيل": "العلي", "عالء": "علاء",
        "امني": "أمين", "الامني": "الأمين",
        "اثري": "اثير", "الاثري": "الاثير",
        "raiyadh": "رياض", "رايض": "رياض", "الرايض": "الرياض",
        "الفتالوي": "الفتلاوي", "الشبالوي": "الشبلاوي", "الشبلاوي": "الشبلاوي",
        "برشى": "بشرى", "رسحان": "رسلان", "اكفائي": "الخفاجي",
        "حسنيه": "حسنية", "خالده": "خالدة", "كريمه": "كريمة", "مظلومه": "مظلومة", "حمزه": "حمزة", "زهره": "زهرة", "هديه": "هدية",
        "ازناد": "زناد", "رسهيد": "رشيد", "رسيح": "رسام",
        "منتضر": "منتظر", "العجييل": "العجيل", "الحمريي": "الحميري",
        "كاضم": "كاظم", "ابرهيم": "ابراهيم",
        "جارهللا": "جار الله", "جارالله": "جار الله",
        "عبدالرحمن": "عبد الرحمن", "عبدالرزاق": "عبد الرزاق",
        "عبدالامير": "عبد الأمير", "عبدالله": "عبد الله"
    }
    
    # دمج القاموس المخصص الذكي الذي يدخله المستخدم من الواجهة مستقبلاً
    if custom_dict:
        exact_matches.update(custom_dict)
    
    # 4. تطبيق نظام التحليل الدقيق للكلمات والأسماء المركبة
    words = text.split()
    fixed_words = []
    for w in words:
        if w in exact_matches:
            fixed_words.append(exact_matches[w])
        else:
            if w.startswith("عبد") and w[3:] in exact_matches:
                fixed_words.append("عبد " + exact_matches[w[3:]])
            elif w.startswith("ال") and w[2:] in exact_matches:
                fixed_words.append("ال" + exact_matches[w[2:]])
            else:
                fixed_words.append(w)
                
    return " ".join(fixed_words)

def fix_arabic_text(text):
    if not text: return ""
    try:
        reshaped = arabic_reshaper.reshape(str(text))
        return get_display(reshaped)
    except Exception:
        return str(text)

def extract_data_from_pdf(file_obj, fix_reversed_arabic=True, custom_dict=None):
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
                                val = val[::-1]
                                val = fix_arabic_swaps(val, custom_dict=custom_dict)
                            cells.append(val)
                        else:
                            cells.append("")
                    if len(cells) < 6 or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "رقم البطاقة" in "".join(cells):
                        continue
                    try:
                        seq = "-"
                        for i in range(len(cells)-1, -1, -1):
                            if cells[i].isdigit() and 0 < len(cells[i]) <= 4:
                                seq = cells[i]
                                break
                        card_num = ""
                        if len(cells) > 6 and cells[6].isdigit() and len(cells[6]) >= 5:
                            card_num = cells[6]
                        else:
                            digits = [c for c in cells if c.isdigit() and len(c) >= 5]
                            if digits: card_num = digits[0]
                        if not card_num: continue
                        name = ""
                        if len(cells) > 4 and not cells[4].isdigit() and len(cells[4]) > 4:
                            name = fix_arabic_swaps(cells[4], custom_dict=custom_dict)
                        else:
                            non_digits = [c for c in cells if not c.isdigit() and len(c) > 3]
                            if non_digits: name = fix_arabic_swaps(non_digits[0], custom_dict=custom_dict)
                        eligible = int(cells[1]) if len(cells) > 1 and cells[1].isdigit() else 0
                        withheld = int(cells[0]) if len(cells) > 0 and cells[0].isdigit() else 0
                        total = int(cells[2]) if len(cells) > 2 and cells[2].isdigit() else 0
                        records[card_num] = {"seq": seq, "name": name, "total": total, "eligible": eligible, "withheld": withheld}
                    except Exception:
                        continue
    return records

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
                    "التسلسل الأصلي": new_val['seq'], "رقم البطاقة": card_num, "اسم رب الأسرة": new_val['name'],
                    "نوع التغيير": "تعديل في البيانات", "تفاصيل المتغيرات": " | ".join(changes)
                })
            del new_data[card_num]
        else:
            rows_output.append({
                "التسلسل الأصلي": old_val['seq'], "رقم البطاقة": card_num, "اسم رب الأسرة": old_val['name'],
                "نوع التغيير": "محذوف / منقول", "تفاصيل المتغيرات": "تم رفع العائلة من قائمة الوكيل في الشهر الجديد"
            })
    for card_num, new_val in new_data.items():
        rows_output.append({
            "التسلسل الأصلي": new_val['seq'], "رقم البطاقة": card_num, "اسم رب الأسرة": new_val['name'],
            "نوع التغيير": "مضاف حديثا", "تفاصيل المتغيرات": "عائلة جديدة تم إنزالها لدى الوكيل"
        })
    return rows_output

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
            fixed_val = fix_arabic_text(val)
            row_cells.append(Paragraph(fixed_val, cell_text_style))
        table_data.append(row_cells)
        
    col_widths = [300, 110, 185, 100, 50]
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

# --- واجهة خيارات معالجة القاموس المطور ---
st.markdown("<h4 style='text-align: right;'>⚙️ خيارات المعالجة والشفاء الإملائي المتقدم للأسماء</h4>", unsafe_allow_html=True)
fix_reversed = st.checkbox("🔄 تفعيل مصحح ومطهر الأسماء الذكي التلقائي (النسخة الجذرية)", value=True)

# معالجة القاموس المخصص من قبل المستخدم تلقائياً
custom_dict = {}
with st.expander("🛠️ لوحة تحكم القاموس الذكي (لتصحيح أي أسماء مشوهة جديدة مستقبلاً)"):
    st.markdown("""
    <p style='text-align: right;'>إذا ظهر لك اسم مشوه جديد في ملف الـ PDF مستقبلاً، اكتب الكلمة المشوهة ثم علامة <b>=</b> ثم الكلمة الصحيحة (كل كلمة في سطر مستقل):</p>
    <p style='text-align: right;'><i>مثال: <br>جاسن = جاسم<br>بشيب = شبيب</i></p>
    """, unsafe_allow_html=True)
    custom_input = st.text_area("أدخل القواعد المخصصة هنا (اختياري):", height=100, placeholder="الكلمة_المشوهة = الكلمة_الصحيحة", label_visibility="collapsed")
    if custom_input:
        for line in custom_input.split('\n'):
            if '=' in line:
                parts = line.split('=')
                wrong = parts[0].strip()
                right = parts[1].strip()
                if wrong and right:
                    custom_dict[wrong] = right

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
            old_extracted = extract_data_from_pdf(old_file, fix_reversed_arabic=fix_reversed, custom_dict=custom_dict)
            new_extracted = extract_data_from_pdf(new_file, fix_reversed_arabic=fix_reversed, custom_dict=custom_dict)
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
                        data=pdf_data, file_name=f"فروقات_{new_file.name}.pdf", mime="application/pdf",
                    )
                else:
                    st.error("لا يمكن تحميل الـ PDF لعدم تثبيت المكتبات الداعمة.")
            else:
                st.success("🎉 تطابق تام! لم يتم العثور على أي تغيير أو تعديل بين القائمتين.")
    else:
        st.error("الرجاء التأكد من رفع كلا الملفين (القديم والجديد) لتتمكن من المقارنة.")
