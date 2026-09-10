import pytest

import app


@pytest.mark.parametrize(
    "desc,card,rec,expect_error",
    [
        ("normal family (4=2+2)", "1234", {"name": "a", "total": 4, "eligible": 2, "withheld": 2}, False),
        ("zero family", "1234", {"name": "a", "total": 0, "eligible": 0, "withheld": 0}, False),
        ("boundary total = max allowed", "1234", {"name": "a", "total": 60, "eligible": 60, "withheld": 0}, False),
        ("total over max", "1234", {"name": "a", "total": 61, "eligible": 61, "withheld": 0}, True),
        ("huge corrupted total (card number leaked in)", "0002158", {"name": "a", "total": 7369723, "eligible": 2158, "withheld": 1}, True),
        ("eligible+withheld off by 1 (source rounding tolerance)", "1234", {"name": "a", "total": 5, "eligible": 3, "withheld": 3}, False),
        ("eligible+withheld off by more than tolerance", "1234", {"name": "a", "total": 5, "eligible": 6, "withheld": 4}, True),
        ("non-int value leaked into a count field", "1234", {"name": "a", "total": "-", "eligible": 2, "withheld": 2}, True),
        ("negative value", "1234", {"name": "a", "total": -1, "eligible": 2, "withheld": 2}, True),
        ("card number with letters", "12A4", {"name": "a", "total": 4, "eligible": 2, "withheld": 2}, True),
        ("card number too short", "12", {"name": "a", "total": 4, "eligible": 2, "withheld": 2}, True),
        ("card number too long", "12345678901", {"name": "a", "total": 4, "eligible": 2, "withheld": 2}, True),
    ],
)
def test_validate_record(desc, card, rec, expect_error):
    err = app.validate_record(card, rec)
    assert (err is not None) == expect_error, f"{desc}: got {err!r}"


def test_validate_and_clean_pair_is_safe_for_known_good_954(agent954_files):
    old_file, new_file = agent954_files
    old_data, new_data, _, _, _ = app.extract_matched_by_either_card(app.extract_records_smart, old_file, new_file)
    is_safe, clean_old, clean_new, errors = app.validate_and_clean_pair(old_data, new_data, "old", "new")
    assert is_safe is True
    assert errors == []
    assert len(clean_old) == len(old_data)
    assert len(clean_new) == len(new_data)


def test_validate_and_clean_pair_is_safe_for_known_good_921(agent921_files):
    old_file, new_file = agent921_files
    old_data, new_data, _, _, _ = app.extract_matched_by_either_card(app.extract_records_smart, old_file, new_file)
    is_safe, clean_old, clean_new, errors = app.validate_and_clean_pair(old_data, new_data, "old", "new")
    assert is_safe is True
    assert errors == []


def test_validate_and_clean_pair_blocks_the_actual_corruption_bug(agent921_mismatched_files):
    """هذا هو الخلل الحقيقي اللي صار بالإنتاج: قراءة جدول الوورد (فيه عمود
    فارغ) عبر المحرك القديم extract_clean_records سرّبت رقم البطاقة لعمود
    'الأفراد الكلية'. الحارس لازم يوقفها (is_safe=False)، مو يمررها."""
    old_file, new_file = agent921_mismatched_files
    old_data, new_data, _, _, _ = app.extract_matched_by_either_card(app.extract_clean_records, old_file, new_file)
    is_safe, clean_old, clean_new, errors = app.validate_and_clean_pair(old_data, new_data, "old", "new")
    assert is_safe is False
    assert len(errors) > 0
    # ولا سجل فاسد وحد ينزل بالنسخة النظيفة
    for rec in list(clean_old.values()) + list(clean_new.values()):
        assert rec["total"] <= app.MAX_REASONABLE_FAMILY_SIZE


def test_smart_extractor_avoids_the_corruption_on_the_same_file(agent921_mismatched_files):
    """نفس الملفين، لكن بالمحرك الذكي (الأساسي فعلياً بالتطبيق) ما يصير
    فيهم أي فساد من الأصل."""
    old_file, new_file = agent921_mismatched_files
    old_data, new_data, _, _, _ = app.extract_matched_by_either_card(app.extract_records_smart, old_file, new_file)
    is_safe, _, _, errors = app.validate_and_clean_pair(old_data, new_data, "old", "new")
    assert is_safe is True
    assert errors == []


def test_validate_record_skip_consistency_check_for_eligible_only_mode():
    """نموذج 'المستحق فقط' يُصفّر الكلي والمحجوب عمداً — فحص الاتساق
    بينهم وبين المستحق لازم يُستثنى صراحة بهذا النمط، وإلا عائلة سليمة
    ترفض غلط لمجرد إنها مستحقة أكثر من 'الكلي=0' المصطنع."""
    rec = {"name": "a", "total": 0, "eligible": 5, "withheld": 0}
    assert app.validate_record("1234", rec, skip_consistency_check=False) is not None
    assert app.validate_record("1234", rec, skip_consistency_check=True) is None


def test_validate_and_clean_pair_skip_consistency_check_propagates():
    old_data = {"1234": {"name": "a", "total": 0, "eligible": 5, "withheld": 0}}
    new_data = {"1234": {"name": "a", "total": 0, "eligible": 5, "withheld": 0}}
    is_safe, clean_old, clean_new, errors = app.validate_and_clean_pair(
        old_data, new_data, "old", "new", skip_consistency_check=True
    )
    assert is_safe is True
    assert errors == []
    assert len(clean_old) == 1
    assert len(clean_new) == 1
