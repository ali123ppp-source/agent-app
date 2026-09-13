"""اختبارات ارتداد (regression) شاملة على خط الأنابيب الكامل: استخراج ←
تحقق/تنظيف ← مقارنة ← تقارير PDF. القيم المتوقعة هنا مأخوذة من نتائج
مُتحقق منها يدوياً سابقاً على نفس الملفات الحقيقية — أي تغيير مستقبلي
يكسر هذي الأرقام لازم يُراجع بوعي، مو يمر بالصدفة."""
import pandas as pd

import app


def _run_pipeline(old_file, new_file, mode="النوع الأول"):
    old_data, new_data, card_col_name, _, _ = app.extract_matched_by_either_card(
        app.extract_records_smart, old_file, new_file
    )
    is_safe, old_data, new_data, errors = app.validate_and_clean_pair(old_data, new_data, "old", "new")
    assert is_safe, f"pipeline should be safe on known-good fixtures, got errors: {errors[:3]}"
    results, results_ref, counters = app.process_comparison(old_data, new_data, mode, card_col_name, "المحرك القياسي")
    return results, counters, card_col_name


def test_pipeline_954_matches_known_baseline(agent954_files):
    old_file, new_file = agent954_files
    _, counters, _ = _run_pipeline(old_file, new_file)
    assert counters["added_fam"] == 148
    assert counters["deleted_fam"] == 0
    assert counters["total_fam"] == 26  # عوائل تغيّرت كليتها


def test_pipeline_921_matches_known_baseline(agent921_files):
    old_file, new_file = agent921_files
    _, counters, _ = _run_pipeline(old_file, new_file)
    assert counters["added_fam"] == 26
    assert counters["deleted_fam"] == 10
    assert counters["total_fam"] == 50


def test_pipeline_921_no_row_has_impossible_values(agent921_files):
    old_file, new_file = agent921_files
    results, _, _ = _run_pipeline(old_file, new_file)
    for row in results:
        for field in ("الأفراد الكلية", "الأفراد المستحقة", "الأفراد المحجوبين"):
            val = row.get(field)
            if val in (None, "-"):
                continue
            assert int(val) <= app.MAX_REASONABLE_FAMILY_SIZE, f"impossible value leaked into report row: {row}"


def test_category_pdf_reports_generate_without_error(agent921_files):
    old_file, new_file = agent921_files
    results, counters, card_col_name = _run_pipeline(old_file, new_file)
    df = pd.DataFrame(results)
    reports, agent_label = app.create_category_pdf_reports(df, card_col_name, new_file.name)
    assert agent_label == "921"
    assert len(reports) > 0
    for rep in reports:
        pdf_bytes = rep["pdf"].getvalue()
        assert pdf_bytes[:4] == b"%PDF", "generated file is not a valid PDF"
        assert len(pdf_bytes) > 1000


def test_combined_pdf_report_generates_without_error(agent921_files):
    old_file, new_file = agent921_files
    results, counters, card_col_name = _run_pipeline(old_file, new_file)
    df = pd.DataFrame(results)
    combined_pdf, agent_label = app.create_combined_pdf_report(df, card_col_name, new_file.name)
    assert combined_pdf is not None
    pdf_bytes = combined_pdf.getvalue()
    assert pdf_bytes[:4] == b"%PDF"


def test_canva_template_also_generates_valid_pdf(agent921_files):
    old_file, new_file = agent921_files
    results, counters, card_col_name = _run_pipeline(old_file, new_file)
    df = pd.DataFrame(results)
    reports, _ = app.create_category_pdf_reports(df, card_col_name, new_file.name, template="canva")
    assert len(reports) > 0
    assert reports[0]["pdf"].getvalue()[:4] == b"%PDF"


def test_classic_report_pdf_generates_single_table_without_error(agent921_files):
    """النموذج الأصلي: ملف PDF واحد بثلاثة أقسام (معدّلة، مضافة، منقولة)."""
    old_file, new_file = agent921_files
    results, counters, card_col_name = _run_pipeline(old_file, new_file)
    df = pd.DataFrame(results)
    pdf_buf, agent_name = app.create_classic_report_pdf(df, card_col_name, "921-FOOD.docx")
    pdf_bytes = pdf_buf.getvalue()
    assert pdf_bytes[:4] == b"%PDF"
    assert len(pdf_bytes) > 1000
    assert agent_name == "921"


def test_classic_report_pdf_separates_added_and_transferred_into_own_sections():
    """اختبار وحدة مركّز: كل حالة (معدّلة/مضافة/منقولة) بجدولها الخاص فقط
    داخل نفس الملف — ولا حالة تختلط بجدول حالة ثانية."""
    import pandas as pd
    df = pd.DataFrame([
        {"التسلسل": "1", "اسم رب الأسرة": "احمد علي احمد", "رقم البطاقة": "1111", "الأفراد الكلية": 3, "الأفراد المستحقة": 3, "الأفراد المحجوبين": 0, "الإحالة": "عائلة مضافة", "meta_status": "added"},
        {"التسلسل": "2", "اسم رب الأسرة": "محمد كريم محمد", "رقم البطاقة": "2222", "الأفراد الكلية": 4, "الأفراد المستحقة": 4, "الأفراد المحجوبين": 0, "الإحالة": "عائلة منقولة", "meta_status": "deleted"},
        {"التسلسل": "3", "اسم رب الأسرة": "علي حسين علي", "رقم البطاقة": "3333", "الأفراد الكلية": 5, "الأفراد المستحقة": 4, "الأفراد المحجوبين": 1, "الإحالة": "نقصان 1 نفر", "meta_status": "modified"},
    ])
    pdf_buf, agent_name = app.create_classic_report_pdf(df, "رقم البطاقة", "test.docx")

    # نستخدم أرقام البطاقات للتحقق (مو الأسماء العربية) لأن استخراج النص
    # من PDF يعيد تشكيل/عكس حروف العربي أحياناً، بينما الأرقام تبقى كما هي.
    import fitz
    doc = fitz.open(stream=pdf_buf.getvalue(), filetype="pdf")
    pages_text = [page.get_text() for page in doc]
    full_text = "".join(pages_text)

    assert "العوائل المضافة" in full_text
    assert "العوائل المنقولة" in full_text
    assert "1111" in full_text
    assert "2222" in full_text
    assert "3333" in full_text

    # الجدول الرئيسي (الصفحة الأولى) ما يحتوي المضاف ولا المنقول
    assert "1111" not in pages_text[0]
    assert "2222" not in pages_text[0]
    assert "3333" in pages_text[0]

    # كل قسم بصفحته الخاصة، بلا تسرّب صفوف حالات ثانية إليه
    added_page = next(p for p in pages_text if "1111" in p)
    assert "2222" not in added_page and "3333" not in added_page
    transferred_page = next(p for p in pages_text if "2222" in p)
    assert "1111" not in transferred_page and "3333" not in transferred_page


def test_classic_status_html_colors_match_word_report_scheme():
    assert 'color:#0000FF' in app._classic_status_html("إضافة طفل")
    assert 'color:#008000' in app._classic_status_html("عائلة مضافة")
    assert 'color:#FF0000' in app._classic_status_html("عائلة منقولة")
    assert 'color:#800000' in app._classic_status_html("حجب كلي")
    multi = app._classic_status_html("تم حجب 1 نفر | إضافة طفل")
    assert 'color:#FF0000' in multi and 'color:#0000FF' in multi


def test_category_section_html_escapes_malicious_name_field():
    """اسم عائلة فيه HTML/JS خام (مصدره ملف مستخدم، مو موثوق) ما يصير جزء
    فعلي من الصفحة — لازم يظهر كنص حرفي مهرّب، مو يكسر بنية الجدول أو
    يحقن سكربت."""
    cat = app.CATEGORY_DEFS[0]
    malicious_row = {
        "اسم رب الأسرة": "<script>alert(1)</script>",
        "رقم البطاقة": "1234",
        "الأفراد الكلية": 3,
        "الأفراد المستحقة": 2,
        "الأفراد المحجوبين": 1,
        "الإحالة": "<img src=x onerror=alert(2)>",
    }
    html_out = app._category_section_html([malicious_row], cat, "رقم البطاقة", "وكيل \"921\" <b>")
    assert "<script>alert(1)</script>" not in html_out
    assert "&lt;script&gt;" in html_out
    assert "onerror=" not in html_out or "&lt;img" in html_out
    assert "<b>" not in html_out  # agent_label نفسه المهرّب ما يفلت أيضاً


def test_process_comparison_ignores_whitespace_only_name_differences():
    """اسم نفسه بالضبط لكن بمسافات مزدوجة/غير منتظمة بملف عن الثاني
    (شائع جداً بين إكسل ووورد) ما لازم يُحتسب "تغيير اسم" مزيّف — بس فرق
    حقيقي بالكلمات لازم يُحتسب فعلاً."""
    old_data = {
        "1111": {"seq": "1", "name": "احمد  خورشيد  كاظم", "total": 4, "eligible": 4, "withheld": 0, "alt_card": ""},
        "2222": {"seq": "2", "name": "علي حسن محمد", "total": 3, "eligible": 3, "withheld": 0, "alt_card": ""},
    }
    new_data = {
        "1111": {"seq": "1", "name": "احمد خورشيد كاظم", "total": 5, "eligible": 5, "withheld": 0, "alt_card": ""},
        "2222": {"seq": "2", "name": "علي حسن كريم", "total": 3, "eligible": 3, "withheld": 0, "alt_card": ""},
    }
    results, _, _ = app.process_comparison(old_data, new_data, "النوع الأول", "رقم البطاقة", "المحرك القياسي")
    by_card = {r["رقم البطاقة"]: r for r in results}

    # 1111: نفس الاسم فعلياً (فرق مسافات بس) + زيادة أفراد — ما لازم يظهر "تغيير الاسم"
    assert "تغيير الاسم" not in by_card["1111"]["الإحالة"]
    assert "إضافة طفل" in by_card["1111"]["الإحالة"]

    # 2222: تغيير حقيقي بالكلمة الأخيرة (محمد → كريم) — لازم يظهر "تغيير الاسم"
    assert "تغيير الاسم" in by_card["2222"]["الإحالة"]


def test_is_minor_typo_difference_detects_single_character_edits():
    """أمثلة حقيقية من بيانات وكيل فعلي: فرق حرف وحيد (حذف/استبدال/تبديل
    حرفين متجاورين) يُعتبر نفس الاسم، وفرق أكبر (اسم مختلف كلياً، أو فرق
    حقيقي بكلمة كاملة) ما يُعتبر."""
    assert app._is_minor_typo_difference("شذى عبد الساده كصاد الفتلاوي", "شذى عبد السادة كصاد الفتلاوي")  # ه↔ة
    assert app._is_minor_typo_difference("عبدالله لبيد روؤف الغلاي", "عبدالله لبيد رؤوف الغلاي")  # تبديل حرفين متجاورين
    assert app._is_minor_typo_difference("صفاء وهبي امين", "صفاء وهي امين")  # حرف محذوف
    assert app._is_minor_typo_difference("صابربن ضياء صالح", "صابرين ضياء صالح")  # استبدال حرف
    assert app._is_minor_typo_difference("ميثم شبل عاصي الكعبي", "ميثم شبل عامي الكعبي")  # استبدال حرف
    assert not app._is_minor_typo_difference("معتز كاظم علي", "محمد خلف خشان الحبيب")  # اسم مختلف كلياً
    assert not app._is_minor_typo_difference("منال مجيد طاهر البو بياض", "منال مجيد طاهر ابو رياض")  # فرق حقيقي بكلمة


def test_process_comparison_ignores_single_letter_typo_and_missing_space_in_name():
    """نفس أمثلة الاختبار السابق لكن من خلال process_comparison كاملة —
    الاسم اللي يختلف بحرف وحيد بس (أو مسافة ناقصة صنعت كلمة إضافية) ما
    يُحتسب تغيير اسم، حتى لو التطبيع البسيط (بدون قص لثلاث كلمات) هو اللي
    يكتشفه صح."""
    old_data = {
        "1111": {"seq": "1", "name": "محمدعلي عبدالواحد عاتي", "total": 5, "eligible": 5, "withheld": 0, "alt_card": ""},
        "2222": {"seq": "2", "name": "معتز كاظم علي", "total": 4, "eligible": 4, "withheld": 0, "alt_card": ""},
    }
    new_data = {
        "1111": {"seq": "1", "name": "محمد علي عبدالواحد عاتي", "total": 5, "eligible": 5, "withheld": 0, "alt_card": ""},
        "2222": {"seq": "2", "name": "محمد خلف خشان الحبيب", "total": 4, "eligible": 4, "withheld": 0, "alt_card": ""},
    }
    results, _, _ = app.process_comparison(old_data, new_data, "النوع الأول", "رقم البطاقة", "المحرك القياسي")
    by_card = {r["رقم البطاقة"]: r for r in results}

    # 1111: نفس الاسم فعلياً، بس مسافة ناقصة ولّدت كلمة إضافية بملف — ما لازم "تغيير الاسم"
    assert "1111" not in by_card or "تغيير الاسم" not in by_card["1111"]["الإحالة"]
    # 2222: اسم مختلف كلياً — لازم "تغيير الاسم"
    assert "تغيير الاسم" in by_card["2222"]["الإحالة"]


def test_status_pills_html_gives_each_status_its_own_distinct_color():
    """صف فيه أكثر من حالة بنفس الوقت (حجب + إضافة طفل) لازم ياخذ فقاعتين
    منفصلتين بلونين مختلفين، مو فقاعة وحدة بلون رمادي موحّد للكل."""
    html = app._status_pills_html("تم حجب 2 نفر | إضافة طفل")
    assert html.count('class="status-pill"') == 2
    assert "تم حجب 2 نفر" in html and "إضافة طفل" in html
    # لونين مختلفين فعلياً (مو نفس الخلفية مكررة)
    block_accent, block_soft, block_dark = app._status_part_color("تم حجب 2 نفر")
    add_accent, add_soft, add_dark = app._status_part_color("إضافة طفل")
    assert block_soft != add_soft

    empty = app._status_pills_html("")
    assert empty == ""


def test_status_part_color_distinguishes_known_categories():
    """كل نوع حالة معروف (حجب كلي/رفع حجب/زيادة أفراد/تغيير اسم...) ياخذ
    لون accent_soft مختلف عن البقية، وحالة غير معروفة تاخذ اللون الرمادي
    الافتراضي."""
    colors = {
        label: app._status_part_color(text)[1]
        for label, text in [
            ("full_block", "حجب كلي"),
            ("block_up", "تم حجب 2 نفر"),
            ("block_down", "تم رفع الحجب عن 1 نفر"),
            ("members_up", "إضافة طفل"),
            ("name_change", "تغيير الاسم من فلان الى علان"),
        ]
    }
    assert len(set(colors.values())) == len(colors)  # كلها ألوان مختلفة عن بعضها
    assert app._status_part_color("نص غير معروف إطلاقاً")[1] == "#F4F6F6"  # رمادي افتراضي


def test_colgroup_seq_column_wide_enough_for_three_digit_numbers():
    """طلب صريح بعد ملاحظة أرقام تسلسل من 3 خانات تنقص بعلامة '...'
    بالجدول المدموج (اللي يجمع تسلسل أطول من أي حالة منفردة) — عمود "ت"
    لازم يكون أوسع من عرضه الأصلي (4%)، والمجموع الكلي يبقى 100%."""
    import re
    html = app._colgroup_html()
    widths = [float(w) for w in re.findall(r'width:([\d.]+)%', html)]
    assert len(widths) == 7
    assert abs(sum(widths) - 100) < 0.01
    assert widths[0] > 4


def test_matched_categories_merges_all_but_added_and_deleted_without_duplication():
    """المضافة والمنقولة تبقيان بجدولهما الخاص، وكل باقي الحالات (حجب،
    زيادة أفراد، تغيير اسم...) تندمج بجدول واحد فقط — وأي عائلة تنطبق
    عليها أكثر من حالة بنفس الوقت (هنا: حجب + زيادة أفراد معاً) تُذكر
    مرة واحدة بس بالجدول المدموج، مو مرتين."""
    df = pd.DataFrame([
        {"اسم رب الأسرة": "احمد علي", "رقم البطاقة": "1111", "الأفراد الكلية": 3, "الأفراد المستحقة": 3, "الأفراد المحجوبين": 0, "الإحالة": "عائلة مضافة", "meta_status": "added"},
        {"اسم رب الأسرة": "محمد كريم", "رقم البطاقة": "2222", "الأفراد الكلية": 4, "الأفراد المستحقة": 4, "الأفراد المحجوبين": 0, "الإحالة": "عائلة منقولة", "meta_status": "deleted"},
        {"اسم رب الأسرة": "علي حسين", "رقم البطاقة": "3333", "الأفراد الكلية": 5, "الأفراد المستحقة": 4, "الأفراد المحجوبين": 1, "الإحالة": "تم حجب 1 نفر | إضافة طفل", "meta_status": "modified"},
        {"اسم رب الأسرة": "حسين جبار", "رقم البطاقة": "4444", "الأفراد الكلية": 3, "الأفراد المستحقة": 3, "الأفراد المحجوبين": 0, "الإحالة": "تغيير الاسم من X الى Y", "meta_status": "modified"},
    ])
    matched = app._matched_categories(df)
    by_key = {cat["key"]: rows for cat, rows in matched}

    assert set(by_key.keys()) == {"added", "deleted", "other_changes"}
    assert len(by_key["added"]) == 1
    assert len(by_key["deleted"]) == 1
    # علي (حجب+زيادة أفراد معاً) وحسين (تغيير اسم) = صفّان بالضبط، لا تكرار
    assert len(by_key["other_changes"]) == 2
    other_cards = {r["رقم البطاقة"] for r in by_key["other_changes"]}
    assert other_cards == {"3333", "4444"}


def test_extract_matched_by_either_card_raises_no_exception_on_empty_files():
    """ملفان بدون أي جدول بيانات: الاستخراج يرجع قواميس فاضية بهدوء (مو
    استثناء) — طبقة main() فوقه هي اللي تقرر توقف العرض للمستخدم بدل ما
    تكمل حساب على بيانات فاضية وتعرض 'تطابق تام' مضلل."""
    import io
    from docx import Document

    def empty_docx():
        buf = io.BytesIO()
        doc = Document()
        doc.add_paragraph("لا يوجد جدول هنا إطلاقاً.")
        doc.save(buf)
        buf.seek(0)
        buf.name = "empty.docx"
        return buf

    old_data, new_data, _, old_dupes, new_dupes = app.extract_matched_by_either_card(
        app.extract_records_smart, empty_docx(), empty_docx()
    )
    assert old_data == {}
    assert new_data == {}
    assert old_dupes == []
    assert new_dupes == []
