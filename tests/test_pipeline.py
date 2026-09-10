"""اختبارات ارتداد (regression) شاملة على خط الأنابيب الكامل: استخراج ←
تحقق/تنظيف ← مقارنة ← تقارير PDF. القيم المتوقعة هنا مأخوذة من نتائج
مُتحقق منها يدوياً سابقاً على نفس الملفات الحقيقية — أي تغيير مستقبلي
يكسر هذي الأرقام لازم يُراجع بوعي، مو يمر بالصدفة."""
import pandas as pd

import app


def _run_pipeline(old_file, new_file, mode="النوع الأول"):
    old_data, new_data, card_col_name = app.extract_matched_by_either_card(
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
