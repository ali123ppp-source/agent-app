"""يضمن وجود جذر المشروع بمسار الاستيراد حتى تشتغل `import app` من داخل
حزمة الاختبارات بدون أي حِيَل، بغض النظر عن مجلد تشغيل pytest."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
