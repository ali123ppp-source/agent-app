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
    """النموذج الأصلي: الجدول الشامل الرئيسي يحوي العوائل المعدّلة فقط
    (المضافة والمنقولة استُبعدتا بطلب صريح)، والمضافة تصير بملف PDF
    مستقل، والمنقولة بقسم/جدول منفصل داخل نفس الملف الرئيسي."""
    old_file, new_file = agent921_files
    results, counters, card_col_name = _run_pipeline(old_file, new_file)
    df = pd.DataFrame(results)
    main_pdf, added_pdf, agent_name = app.create_classic_report_pdf(df, card_col_name, "921-FOOD.docx")
    main_bytes = main_pdf.getvalue()
    assert main_bytes[:4] == b"%PDF"
    assert len(main_bytes) > 1000
    assert agent_name == "921"

    added_count = sum(1 for r in results if r.get("meta_status") == "added")
    if added_count:
        assert added_pdf is not None
        added_bytes = added_pdf.getvalue()
        assert added_bytes[:4] == b"%PDF"
    else:
        assert added_pdf is None


def test_classic_report_pdf_excludes_added_and_transferred_from_main_table():
    """اختبار وحدة مركّز: سجل 'مضاف' وسجل 'منقول' ما يظهرون بالجدول
    الرئيسي إطلاقاً — المضاف يروح لملف PDF منفصل تماماً، والمنقول يروح
    لقسم/جدول منفصل بنفس الملف (تحت عنوان 'العوائل المنقولة')."""
    import pandas as pd
    df = pd.DataFrame([
        {"التسلسل": "1", "اسم رب الأسرة": "احمد علي احمد", "رقم البطاقة": "1111", "الأفراد الكلية": 3, "الأفراد المستحقة": 3, "الأفراد المحجوبين": 0, "الإحالة": "عائلة مضافة", "meta_status": "added"},
        {"التسلسل": "2", "اسم رب الأسرة": "محمد كريم محمد", "رقم البطاقة": "2222", "الأفراد الكلية": 4, "الأفراد المستحقة": 4, "الأفراد المحجوبين": 0, "الإحالة": "عائلة منقولة", "meta_status": "deleted"},
        {"التسلسل": "3", "اسم رب الأسرة": "علي حسين علي", "رقم البطاقة": "3333", "الأفراد الكلية": 5, "الأفراد المستحقة": 4, "الأفراد المحجوبين": 1, "الإحالة": "نقصان 1 نفر", "meta_status": "modified"},
    ])
    main_pdf, added_pdf, agent_name = app.create_classic_report_pdf(df, "رقم البطاقة", "test.docx")
    assert added_pdf is not None

    # نستخدم أرقام البطاقات للتحقق (مو الأسماء العربية) لأن استخراج النص
    # من PDF يعيد تشكيل/عكس حروف العربي أحياناً، بينما الأرقام تبقى كما هي.
    import fitz
    main_doc = fitz.open(stream=main_pdf.getvalue(), filetype="pdf")
    main_text = "".join(page.get_text() for page in main_doc)
    assert "1111" not in main_text  # المضاف ما يظهر بالرئيسي
    assert "3333" in main_text      # المعدّل يظهر بالرئيسي
    assert "العوائل المنقولة" in main_text  # قسم المنقولة موجود بنفس الملف
    assert "2222" in main_text      # وسجل المنقول فعلياً موجود بقسمه

    added_doc = fitz.open(stream=added_pdf.getvalue(), filetype="pdf")
    added_text = "".join(page.get_text() for page in added_doc)
    assert "1111" in added_text
    assert "2222" not in added_text
    assert "3333" not in added_text


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
