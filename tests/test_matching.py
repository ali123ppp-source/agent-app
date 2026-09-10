import app


def test_dual_key_matching_921(agent921_files):
    old_file, new_file = agent921_files
    old_data, new_data, card_col_name = app.extract_matched_by_either_card(
        app.extract_records_smart, old_file, new_file
    )
    common = set(old_data.keys()) & set(new_data.keys())
    assert len(old_data) == 406
    assert len(new_data) == 422
    assert len(common) == 396
    assert card_col_name == "رقم البطاقة"


def test_dual_key_matching_954(agent954_files):
    old_file, new_file = agent954_files
    old_data, new_data, card_col_name = app.extract_matched_by_either_card(
        app.extract_records_smart, old_file, new_file
    )
    common = set(old_data.keys()) & set(new_data.keys())
    # كل عوائل الإكسل القديم (254) لهم تطابق بالوورد الحديث — صفر محذوفات
    assert len(common) == 254


def test_dual_key_matching_with_mismatched_files_finds_no_matches(agent921_mismatched_files):
    """هذا زوج ملفات مش متناظر فعلياً (لقطتان مختلفتان) — التطابق صفر هو
    السلوك الصحيح هنا، مو خطأ. الاختبار يوثّق هذا السلوك المتوقع صراحة
    بدل ما يُكتشف بالصدفة بالإنتاج زي ما صار فعلياً."""
    old_file, new_file = agent921_mismatched_files
    old_data, new_data, _ = app.extract_matched_by_either_card(
        app.extract_records_smart, old_file, new_file
    )
    common = set(old_data.keys()) & set(new_data.keys())
    assert len(common) == 0


def test_merge_records_by_either_card_matches_via_alt_when_primary_differs():
    """اختبار وحدة مركّز لدالة الدمج نفسها بمعزل عن قراءة الملفات: عائلة
    نفس البطاقة الحديثة بس بطاقة قديمة مختلفة الشكل بالملفين لازم توصف
    كمتطابقة عبر alt_card، مو تُحسب مضافة ومحذوفة بالخطأ."""
    old_data = {
        "1111": {"name": "a", "total": 3, "eligible": 3, "withheld": 0, "alt_card": "9999999"},
    }
    new_data = {
        "2222": {"name": "a", "total": 3, "eligible": 2, "withheld": 1, "alt_card": "9999999"},
    }
    unified_old, unified_new = app.merge_records_by_either_card(old_data, new_data)
    assert set(unified_old.keys()) == {"1111"}
    assert unified_new["1111"]["eligible"] == 2  # جاب سجل new الصحيح عبر alt_card


def test_merge_records_by_either_card_no_double_counting_on_ambiguous_alt():
    """لو سجلين قدامى عندهم نفس alt_card (تصادم نادر)، ما يصير أي سجل
    جديد يتحسب مرتين كمتطابق مع الاثنين."""
    old_data = {
        "1111": {"name": "a", "total": 1, "eligible": 1, "withheld": 0, "alt_card": "5555555"},
        "1112": {"name": "b", "total": 2, "eligible": 2, "withheld": 0, "alt_card": "5555555"},
    }
    new_data = {
        "2222": {"name": "a", "total": 1, "eligible": 1, "withheld": 0, "alt_card": "5555555"},
    }
    unified_old, unified_new = app.merge_records_by_either_card(old_data, new_data)
    matched_keys = [k for k in unified_old if k in unified_new]
    assert len(matched_keys) == 1  # مو الاثنين
