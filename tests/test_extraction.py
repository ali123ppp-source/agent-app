import app


def test_extract_records_smart_954_counts(agent954_files):
    old_file, new_file = agent954_files
    old_data, old_dupes = app.extract_records_smart(old_file, card_type="new")
    new_data, new_dupes = app.extract_records_smart(new_file, card_type="new")
    assert len(old_data) == 254
    assert len(new_data) == 402
    assert old_dupes == []
    assert new_dupes == []


def test_extract_records_smart_921_counts(agent921_files):
    old_file, new_file = agent921_files
    old_data, _ = app.extract_records_smart(old_file, card_type="new")
    new_data, _ = app.extract_records_smart(new_file, card_type="new")
    assert len(old_data) == 406
    assert len(new_data) == 422


def test_extract_records_smart_handles_multi_sheet_without_repeated_header(agent954_files):
    """ملف 954 الإكسل مقسوم على شيتين (Table 1 له صف عناوين، Table 2
    بيانات فقط بدون تكرار العناوين) — هذا بالضبط ما يثبت إن last_role_map
    يشتغل صح عبر أوراق متعددة."""
    old_file, _ = agent954_files
    old_data, _ = app.extract_records_smart(old_file, card_type="new")
    assert len(old_data) > 200  # لو فشل بالورقة الثانية كان العدد يطلع أقل بكثير


def test_preview_columns_detects_expected_roles_xlsx(agent954_files):
    old_file, _ = agent954_files
    preview = app.preview_columns_for_file(old_file)
    assert preview is not None
    assert "اسم رب الأسرة" in preview["detected"]
    assert "الأفراد المستحقة" in preview["detected"]
    assert "الأفراد المحجوبين" in preview["detected"]
    assert len(preview["sample_records"]) == 2


def test_preview_columns_detects_expected_roles_docx(agent954_files):
    _, new_file = agent954_files
    preview = app.preview_columns_for_file(new_file)
    assert preview is not None
    assert "اسم رب الأسرة" in preview["detected"]
    assert "الأفراد الكلية" in preview["detected"]


def test_preview_columns_returns_none_for_file_with_no_clear_table():
    """ملف نصي عشوائي بدون أي جدول/عناوين معروفة — لازم يرجع None بدل ما
    يرمي استثناء أو يخمّن أعمدة عشوائية."""
    import io
    buf = io.BytesIO()
    from docx import Document
    doc = Document()
    doc.add_paragraph("هذا ملف بدون أي جدول بيانات إطلاقاً.")
    doc.save(buf)
    buf.seek(0)
    buf.name = "empty.docx"
    assert app.preview_columns_for_file(buf) is None
