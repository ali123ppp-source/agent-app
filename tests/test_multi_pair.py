"""اختبارات ميزة مقارنة عدة أزواج ملفات دفعة وحدة: تحديد أي xlsx يرتبط
بأي docx تلقائياً حسب تطابق البيانات الفعلية (رقم البطاقة)، وليس أسماء
الملفات أو ترتيب الرفع."""
import io

import app


def test_auto_pair_files_by_content_matches_correct_agents_despite_shuffled_order(agent954_files, agent921_files):
    """رفع ملفات وكيلين معاً بترتيب مقصود مبعثر (921 قبل 954 بالإكسل،
    والعكس بالوورد) — النظام لازم يربط كل وكيل بملفه الصحيح حسب تطابق
    أرقام البطاقات الفعلي، مو حسب ترتيب الرفع."""
    old_954, new_954 = agent954_files
    old_921, new_921 = agent921_files

    xlsx_files = [old_921, old_954]
    docx_files = [new_954, new_921]

    pairs, unmatched_xlsx, unmatched_docx = app.auto_pair_files_by_content(xlsx_files, docx_files)

    assert unmatched_xlsx == []
    assert unmatched_docx == []
    assert len(pairs) == 2

    pair_map = {xf.name: (df.name, overlap) for xf, df, overlap in pairs}
    assert pair_map[old_954.name][0] == new_954.name
    assert pair_map[old_921.name][0] == new_921.name
    # كل زوج صحيح لازم يكون له تطابق فعلي بعشرات البطاقات، مو صدفة رقم أو رقمين
    assert pair_map[old_954.name][1] > 50
    assert pair_map[old_921.name][1] > 50


def _make_simple_docx(rows):
    from docx import Document
    doc = Document()
    table = doc.add_table(rows=1 + len(rows), cols=4)
    header = table.rows[0].cells
    header[0].text, header[1].text, header[2].text, header[3].text = "ت", "رقم البطاقة", "اسم رب الأسرة", "الأفراد المستحقة"
    for i, (card, name) in enumerate(rows, start=1):
        cells = table.rows[i].cells
        cells[0].text, cells[1].text, cells[2].text, cells[3].text = str(i), card, name, "2"
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    buf.name = f"f{id(rows)}.docx"
    return buf


def test_auto_pair_files_by_content_leaves_zero_overlap_file_unmatched():
    """ملف xlsx (هنا بصيغة docx لتبسيط الإنشاء داخل الاختبار، الدالة لا
    تفرّق بينهما) ما عنده ولا بطاقة مشتركة مع أي ملف بالمجموعة المقابلة
    لازم يرجع ضمن unmatched، مو يترابط غلط بأقرب شي متوفر."""
    matching_a = _make_simple_docx([("1001", "احمد علي"), ("1002", "محمد كريم")])
    matching_b = _make_simple_docx([("1001", "احمد علي"), ("1002", "محمد كريم")])
    lonely = _make_simple_docx([("9999999", "شخص غريب تماما")])

    pairs, unmatched_xlsx, unmatched_docx = app.auto_pair_files_by_content([matching_a, lonely], [matching_b])

    assert len(pairs) == 1
    assert pairs[0][0].name == matching_a.name
    assert pairs[0][1].name == matching_b.name
    assert pairs[0][2] == 2
    assert unmatched_xlsx == [lonely]
    assert unmatched_docx == []


def test_file_key_set_includes_alt_card():
    records = {
        "123": {"name": "a", "alt_card": "9999"},
        "456": {"name": "b", "alt_card": ""},
    }
    keys = app._file_key_set(records)
    assert keys == {"123", "9999", "456"}
