import io
import os

import pytest

FIXTURES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")


def make_file(path, name):
    """يبني كائن BytesIO بخاصية .name (يحاكي ملف Streamlit المرفوع)، وبما
    إن BytesIO يدعم seek/read بشكل كامل، يصلح مباشرة لدوال الاستخراج
    اللي تتوقع file_obj.name و file_obj.seek(0)."""
    with open(path, "rb") as f:
        data = f.read()
    buf = io.BytesIO(data)
    buf.name = name
    return buf


@pytest.fixture
def fixture_path():
    def _path(filename):
        return os.path.join(FIXTURES_DIR, filename)
    return _path


@pytest.fixture
def agent954_files():
    old = make_file(os.path.join(FIXTURES_DIR, "agent954_old.xlsx"), "954.pdf_2.xlsx")
    new = make_file(os.path.join(FIXTURES_DIR, "agent954_new.docx"), "954.pdf_1.docx")
    return old, new


@pytest.fixture
def agent921_files():
    old = make_file(os.path.join(FIXTURES_DIR, "agent921_old.xlsx"), "921.pdf_2.xlsx")
    new = make_file(os.path.join(FIXTURES_DIR, "agent921_new.docx"), "921.pdf_1.docx")
    return old, new


@pytest.fixture
def agent921_mismatched_files():
    """زوج ملفات لا يتشاركان أي رقم بطاقة فعلياً (لقطتان مختلفتان) —
    يُستخدم للتأكد إن النظام يتعامل مع "صفر تطابق" بشكل نظيف (الكل
    مضاف/محذوف) بدون أي فساد بالأرقام، مو للتأكد من نجاح المطابقة."""
    old = make_file(os.path.join(FIXTURES_DIR, "agent921_mismatched_old.xlsx"), "921.pdf_5.xlsx")
    new = make_file(os.path.join(FIXTURES_DIR, "agent921_mismatched_new.docx"), "921.pdf_1.docx")
    return old, new


@pytest.fixture
def bugfix_shifted_files():
    """زوج ملفات مصطنع يحاكي بالضبط سبب ظهور "الكل مضاف/محجوب": ملف xlsx
    قديم بورقتين (الأولى بعناوين، والثانية بلا عناوين وبعمود فارغ إضافي يزيح
    عمودي البطاقة)، وملف docx جديد فيه عمود مرجعي قديم بأصفار بادئة
    ("0000022") بجانب عمود "البطاقة الجديدة" الحقيقي."""
    old = make_file(os.path.join(FIXTURES_DIR, "bugfix_shifted_old.xlsx"), "shifted_old.xlsx")
    new = make_file(os.path.join(FIXTURES_DIR, "bugfix_shifted_new.docx"), "shifted_new.docx")
    return old, new
