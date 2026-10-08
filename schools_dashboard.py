import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import glob
import os
import math
import io
import re
import folium
from folium import plugins
from streamlit_folium import st_folium
import streamlit.components.v1 as components
import json as _json
import json
import requests
import zipfile
import xlsxwriter
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                      Paragraph, Spacer)
import arabic_reshaper
from bidi.algorithm import get_display

# إعداد الصفحة
st.set_page_config(page_title="لوحة تحكم المدارس الحكومية", layout="wide")

# ============ التصميم ============
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700;800;900&display=swap');

:root {
  --primary: #1e40af;
  --primary-light: #3b82f6;
  --primary-dark: #1e3a8a;
  --secondary: #f59e0b;
  --success: #10b981;
  --danger: #ef4444;
  --bg: #f0f9ff;
  --card-bg: #ffffff;
  --text: #1e293b;
  --text-muted: #64748b;
  --border: #e2e8f0;
}

* { font-family: 'Cairo', sans-serif !important; }

.stApp {
  background: linear-gradient(135deg, #e0f2fe 0%, #bae6fd 50%, #7dd3fc 100%);
  min-height: 100vh;
}

.block-container {
  padding: 0 1.5rem 3rem;
  max-width: 1800px;
}

#MainMenu, footer, header { visibility: hidden; }

/* ====== الترويسة ====== */
.hero {
  position: relative;
  overflow: hidden;
  margin: 0 -1.5rem 1.5rem;
  padding: 40px 40px 35px;
  background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 50%, #3b82f6 100%);
  border-radius: 0 0 32px 32px;
  box-shadow: 0 25px 50px -12px rgba(30,64,175,.4);
}
.hero::before {
  content: "";
  position: absolute;
  top: -50%;
  right: -10%;
  width: 600px;
  height: 600px;
  background: radial-gradient(circle, rgba(255,255,255,.1) 0%, transparent 70%);
}
.hero-in {
  position: relative;
  z-index: 2;
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 20px;
  flex-wrap: wrap;
}
.hero h1 {
  color: #fff;
  font-size: 2.5rem;
  font-weight: 900;
  margin: 0;
  text-shadow: 0 2px 10px rgba(0,0,0,.2);
}
.hero p {
  color: #bfdbfe;
  margin: .5rem 0 0;
  font-size: 1.05rem;
}
.badges { display: flex; gap: .6rem; flex-wrap: wrap; }
.badge {
  background: rgba(255,255,255,.12);
  backdrop-filter: blur(10px);
  border: 1px solid rgba(255,255,255,.2);
  color: #dbeafe;
  padding: .45rem 1rem;
  border-radius: 999px;
  font-size: .85rem;
  font-weight: 600;
}

/* ====== بطاقات KPI ====== */
.kpis {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 16px;
  margin: 0 0 24px;
}
.kpi {
  background: var(--card-bg);
  border-radius: 20px;
  padding: 24px;
  box-shadow: 0 4px 6px -1px rgba(0,0,0,.1);
  transition: all .3s cubic-bezier(.4,0,.2,1);
  position: relative;
  overflow: hidden;
  border: 1px solid var(--border);
}
.kpi::before {
  content: "";
  position: absolute;
  top: 0; left: 0; right: 0;
  height: 4px;
}
.kpi:hover {
  transform: translateY(-6px);
  box-shadow: 0 20px 25px -5px rgba(0,0,0,.1);
}
.kpi .ic { font-size: 2rem; margin-bottom: 8px; display: block; }
.kpi .v { font-size: 2rem; font-weight: 800; color: var(--text); line-height: 1.2; }
.kpi .t { font-size: .85rem; color: var(--text-muted); font-weight: 600; margin-top: 4px; }
.kpi.a::before { background: linear-gradient(90deg, #3b82f6, #60a5fa); }
.kpi.b::before { background: linear-gradient(90deg, #10b981, #34d399); }
.kpi.c::before { background: linear-gradient(90deg, #f59e0b, #fbbf24); }
.kpi.d::before { background: linear-gradient(90deg, #8b5cf6, #a78bfa); }
.kpi.e::before { background: linear-gradient(90deg, #ef4444, #f87171); }
.kpi.f::before { background: linear-gradient(90deg, #06b6d4, #22d3ee); }

/* ====== الفلاتر ====== */
.filters-anchor { display: block; width: 0; height: 0; }
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child {
  background: var(--card-bg);
  border-radius: 20px;
  padding: 24px 20px;
  box-shadow: 0 4px 6px -1px rgba(0,0,0,.1);
  border: 1px solid var(--border);
  align-self: start;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child h3 {
  color: var(--text);
  border-bottom: 2px solid var(--primary);
  padding-bottom: .5rem;
  margin-top: .2rem;
  font-weight: 700;
}

/* ====== عناصر الإدخال ====== */
.stSelectbox > div > div, .stMultiSelect > div > div {
  background: #f8fafc;
  border: 2px solid var(--border);
  border-radius: 12px;
  transition: all .2s;
}
.stSelectbox > div > div:hover, .stMultiSelect > div > div:hover {
  border-color: var(--primary);
  box-shadow: 0 0 0 3px rgba(59,130,246,.1);
}

/* ====== الأزرار ====== */
.stButton > button {
  width: 100%;
  background: linear-gradient(135deg, var(--primary), var(--primary-dark));
  color: #fff;
  border: none;
  border-radius: 12px;
  padding: .75rem 1.5rem;
  font-weight: 700;
  font-size: 1rem;
  box-shadow: 0 4px 14px 0 rgba(30,64,175,.4);
  transition: all .3s;
}
.stButton > button:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 25px 0 rgba(30,64,175,.5);
}

/* ====== التبويبات ====== */
.stTabs [data-baseweb="tab-list"] {
  gap: 12px;
  border-bottom: 2px solid var(--border);
  flex-wrap: wrap;
  overflow: visible;
  padding-bottom: 0;
}
.stTabs [data-baseweb="tab"] {
  white-space: nowrap;
  font-size: 1.15rem;
  background: transparent;
  color: var(--text-muted);
  border-radius: 12px 12px 0 0;
  padding: .85rem 1.8rem;
  font-weight: 800;
  border: 2px solid transparent;
  border-bottom: none;
  transition: all .2s;
}
.stTabs [data-baseweb="tab"]:hover {
  background: rgba(59,130,246,.05);
  color: var(--primary);
}
.stTabs [aria-selected="true"] {
  background: var(--card-bg);
  color: var(--primary) !important;
  border-color: var(--border);
  border-bottom: 2px solid var(--card-bg);
  margin-bottom: -2px;
  box-shadow: 0 -4px 10px rgba(0,0,0,.05);
}

/* ====== الجداول ====== */
[data-testid="stDataFrame"] {
  border: 1px solid var(--border);
  border-radius: 16px;
  overflow: hidden;
  box-shadow: 0 4px 6px -1px rgba(0,0,0,.05);
}

/* ====== عناوين ====== */
h1, h2, h3, h4 { color: var(--text); font-weight: 700; }
hr { border-color: var(--border); margin: 1.5rem 0; }

/* ====== بطاقات المحتوى ====== */
.content-card {
  background: var(--card-bg);
  border-radius: 20px;
  padding: 24px;
  box-shadow: 0 4px 6px -1px rgba(0,0,0,.1);
  border: 1px solid var(--border);
  margin-bottom: 20px;
}

/* ====== إشعارات ====== */
.stSuccess, .stInfo, .stWarning, .stError {
  border-radius: 12px;
  padding: 1rem 1.5rem;
  font-weight: 500;
}

/* ====== شريط التقدم ====== */
.stProgress > div > div > div {
  background: linear-gradient(90deg, var(--primary), var(--primary-light));
  border-radius: 10px;
}

/* ====== تحسين التمرير ====== */
::-webkit-scrollbar { width: 8px; height: 8px; }
::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 10px; }
::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 10px; }
::-webkit-scrollbar-thumb:hover { background: #94a3b8; }

/* ====== تأثيرات الظهور ====== */
@keyframes fadeIn {
  from { opacity: 0; transform: translateY(20px); }
  to { opacity: 1; transform: translateY(0); }
}
.kpi, .content-card { animation: fadeIn .5s ease-out forwards; }

/* ====== عمود البحث (أسود) ====== */
[data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child {
  background: linear-gradient(180deg, #1a1a1a 0%, #0d0d0d 100%);
  border-radius: 20px;
  padding: 24px 20px;
  box-shadow: 0 10px 30px rgba(0,0,0,.3);
  border: 1px solid #333;
  align-self: start;
}
[data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child h3 {
  color: #fff;
  border-bottom: 2px solid #3b82f6;
  padding-bottom: .5rem;
  margin-top: .2rem;
  font-weight: 700;
}
[data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child label {
  color: #e0e0e0;
  font-weight: 600;
}
[data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stTextInput > div > div {
  background: #2a2a2a;
  border: 2px solid #444;
  border-radius: 10px;
  color: #fff;
}
[data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stTextInput > div > div:focus {
  border-color: #3b82f6;
  box-shadow: 0 0 0 3px rgba(59,130,246,.2);
}
[data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stRadio > div > label {
  color: #e0e0e0;
}
[data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stDataFrame {
  border: 1px solid #444;
  border-radius: 12px;
  overflow: hidden;
}
[data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stAlert {
  border-radius: 10px;
}

/* ====== عمود الفلاتر (أسود) ====== */
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child {
  background: linear-gradient(180deg, #1a1a1a 0%, #0d0d0d 100%);
  border-radius: 20px;
  padding: 24px 20px;
  box-shadow: 0 10px 30px rgba(0,0,0,.3);
  border: 1px solid #333;
  align-self: start;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child h3 {
  color: #fff;
  border-bottom: 2px solid #3b82f6;
  padding-bottom: .5rem;
  margin-top: .2rem;
  font-weight: 700;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child label {
  color: #e0e0e0;
  font-weight: 600;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stSelectbox > div > div,
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stMultiSelect > div > div {
  background: #2a2a2a;
  border: 2px solid #444;
  border-radius: 10px;
  color: #fff !important;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stSelectbox > div > div:hover,
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stMultiSelect > div > div:hover {
  border-color: #3b82f6;
  box-shadow: 0 0 0 3px rgba(59,130,246,.2);
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stSelectbox > div > div * {
  color: #fff !important;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stMultiSelect > div > div * {
  color: #fff !important;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stSelectbox > div > div input,
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stMultiSelect > div > div input {
  color: #fff !important;
  caret-color: #fff;
}
[data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stTextInput > div > div * {
  color: #fff !important;
}
[data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stTextInput > div > div input {
  color: #fff !important;
  caret-color: #fff;
}
[data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stRadio > div > label * {
  color: #e0e0e0 !important;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stButton > button {
  background: linear-gradient(135deg, #3b82f6, #1e40af);
  color: #fff;
  border: none;
  border-radius: 12px;
  padding: .75rem 1.5rem;
  font-weight: 700;
  font-size: 1rem;
  box-shadow: 0 4px 14px 0 rgba(59,130,246,.4);
  transition: all .3s;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stButton > button:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 25px 0 rgba(59,130,246,.5);
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stDownloadButton > button {
  background: linear-gradient(135deg, #10b981, #047857);
  color: #fff;
  border: none;
  border-radius: 12px;
  padding: .6rem 1rem;
  font-weight: 700;
  font-size: .92rem;
  box-shadow: 0 4px 12px 0 rgba(16,185,129,.35);
  transition: all .3s;
  width: 100%;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stDownloadButton > button:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 20px 0 rgba(16,185,129,.5);
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stDownloadButton > button * {
  color: #fff !important;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child h4 {
  color: #fff;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child .stCaptionContainer p {
  color: #cbd5e1 !important;
  font-size: .8rem;
}
[data-testid="stHorizontalBlock"]:has(.filters-anchor) > [data-testid="stColumn"]:first-child hr {
  border-color: #444;
  margin: 1rem 0;
}
</style>
""", unsafe_allow_html=True)

# الترويسة
st.markdown('''
<div class="hero"><div class="hero-in">
  <div>
    <h1>🏫 لوحة تحكم المدارس الحكومية</h1>
    <p>تحليل شامل للبيانات التعليمية — آخر تحديث: 22/9/2026</p>
  </div>
  <div class="badges">
    <span class="badge">● البيانات محمّلة</span>
    <span class="badge">📊 تحليل تفاعلي</span>
    <span class="badge">📈 رسوم بيانية</span>
  </div>
</div></div>
''', unsafe_allow_html=True)

# قراءة البيانات
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_NAME = 'المدارس الحكومية محدث جديد حتى تاريخ 221-9-2026.xlsx'
DATA_FILE = os.path.join(BASE_DIR, DATA_NAME)

# لو نُقل الملف، نبحث في هذه الأماكن بالترتيب
CANDIDATES = [
    DATA_FILE,
    os.path.join(os.path.dirname(BASE_DIR), DATA_NAME),
    os.path.join(os.path.dirname(os.path.dirname(BASE_DIR)), DATA_NAME),
]

@st.cache_data
def _read_excel(path_or_bytes, is_bytes=False, sheet_pref=None):
    src = io.BytesIO(path_or_bytes) if is_bytes else path_or_bytes
    xl = pd.ExcelFile(src)
    if sheet_pref and sheet_pref in xl.sheet_names:
        sheet = sheet_pref
    elif 'Sheet1' in xl.sheet_names:
        sheet = 'Sheet1'
    else:
        sheet = xl.sheet_names[0]
    d = pd.read_excel(src, sheet_name=sheet)
    d.columns = [str(c).strip() for c in d.columns]
    return d, sheet, xl.sheet_names


def default_path():
    path = next((q for q in CANDIDATES if os.path.isfile(q)), None)
    if path is None:
        files = [f for f in glob.glob(os.path.join(BASE_DIR, '*.xlsx'))
                 if '~$' not in os.path.basename(f)]
        path = files[0] if files else None
    return path


def clean_frame(d):
    for c in ['مجموع الطلاب', 'مجموع الطلاب الذكور', 'مجموع الطلاب اناث',
              'مجموع الطلاب الاناث']:
        if c in d.columns:
            d[c] = pd.to_numeric(d[c], errors='coerce')
    return d


def load_data(upload_bytes=None, upload_name=None, sheet_pref=None):
    """الملف المرفوع إن وُجد، وإلا الملف الافتراضي."""
    if upload_bytes is not None:
        d, sh, _ = _read_excel(upload_bytes, True, sheet_pref)
        return clean_frame(d), (upload_name or 'الملف المرفوع'), sh
    path = default_path()
    if path is None:
        st.error('لم يتم العثور على %s — ارفع ملفاً من لوحة الاستيراد.' % DATA_NAME)
        st.stop()
    d, sh, _ = _read_excel(path, False, sheet_pref)
    return clean_frame(d), os.path.basename(path), sh

# ============ الأعمدة أولاً (حتى يظهر زر الاستيراد في اللوحة) ============
col1, col2, col3 = st.columns([1, 4, 1])

# ---------- لوحة استيراد الملف ----------
with col1:
    st.markdown('<span class="import-anchor"></span>', unsafe_allow_html=True)
    st.markdown("#### \U0001F4E5 استيراد ملف")
    uploaded_file = st.file_uploader(
        "ارفع ملف إكسل (.xlsx / .xlsm)", type=["xlsx", "xlsm"],
        key="uploader",
        help="عند رفع أي ملف يُطبَّق البرنامج عليه فوراًinstead من الملف "
             "الافتراضي. يجب أن يحتوي نفس الأعمدة.")
    _up_bytes = uploaded_file.getvalue() if uploaded_file is not None else None
    _up_name = uploaded_file.name if uploaded_file is not None else None
    _sheet_choice = None
    if _up_bytes is not None:
        try:
            _xl_names = pd.ExcelFile(io.BytesIO(_up_bytes)).sheet_names
            if len(_xl_names) > 1:
                _sheet_choice = st.selectbox("الورقة", _xl_names, key="up_sheet")
            else:
                _sheet_choice = _xl_names[0]
        except Exception as exc:
            st.error("تعذّرت قراءة الملف: %s" % exc)
            _up_bytes = None
            _up_name = None
        if _up_bytes is not None and st.button(
                "\U0001F5D1 العودة للملف الافتراضي", key="back_default",
                use_container_width=True):
            for _k in ('uploader', 'up_sheet'):
                st.session_state.pop(_k, None)
            st.rerun()

with st.spinner("جارٍ تحميل البيانات…"):
    df, DATA_SOURCE, SHEET_USED = load_data(_up_bytes, _up_name, _sheet_choice)

with col1:
    if _up_bytes is not None:
        st.success("\U0001F7E2 مُطبَّق على: **%s** — الورقة **%s** — %s صف"
                   % (DATA_SOURCE, SHEET_USED, f"{len(df):,}"))
    else:
        st.caption("\U0001F4C4 %s — الورقة **%s** — %s صف"
                   % (DATA_SOURCE, SHEET_USED, f"{len(df):,}"))
    st.markdown("---")

# ============ الفلاتر والمحتوى ============

# ============================================================
#  التصدير: Excel · PDF · ZIP(CSV) · JSON
# ============================================================
def ar(txt):
    """تشكيل النص العربي ليُعرض صحيحاً في PDF."""
    try:
        return get_display(arabic_reshaper.reshape(str(txt)))
    except Exception:
        return str(txt)


def _csv(dframe):
    return dframe.to_csv(index=False).encode('utf-8-sig')


def build_sheets(frame):
    """كل الجداول التي يصدّرها البرنامج."""
    out = {}
    out['البيانات كاملة'] = frame
    for col in ['المديرية', 'المحافظة', 'الجنس', 'نوع التعليم', 'الملكية',
                'أدنى صف', 'أعلى صف', 'فترة المؤسسة']:
        if col in frame.columns:
            g = (frame.groupby(col, dropna=False)
                  .agg({'اسم المدرسة': 'count',
                        'مجموع الطلاب': 'sum',
                        'مجموع الطلاب الذكور': 'sum',
                        'مجموع الطلاب الاناث': 'sum',
                        'الطاقة الاستيعابية': 'sum'})
                  .reset_index().rename(columns={
                      col: col,
                      'اسم المدرسة': 'عدد المدارس',
                      'مجموع الطلاب': 'إجمالي الطلاب',
                      'مجموع الطلاب الذكور': 'الذكور',
                      'مجموع الطلاب الاناث': 'الإناث',
                      'الطاقة الاستيعابية': 'إجمالي الطاقة'}))
            if 'عدد المدارس' in g.columns:
                g = g.sort_values('عدد المدارس', ascending=False)
            out['توزيع: ' + col] = g
    if {'مجموع الطلاب', 'اسم المدرسة'}.issubset(frame.columns):
        top = frame.sort_values('مجموع الطلاب', ascending=False)
        out['المدارس حسب عدد الطلاب'] = top[['اسم المدرسة', 'المديرية', 'المحافظة',
                                      'الجنس', 'نوع التعليم', 'الملكية',
                                      'أدنى صف', 'أعلى صف', 'مجموع الطلاب',
                                      'الطاقة الاستيعابية']].drop(
            columns=[c for c in ['الطاقة الاستيعابية']
                     if c not in frame.columns], errors='ignore')
    # مؤشرات عامة
    if 'مجموع الطلاب' in frame.columns:
        kpi = pd.DataFrame([
            ['عدد المدارس', f"{len(frame):,}"],
            ['إجمالي الطلاب', f"{int(frame['مجموع الطلاب'].sum()):,}"],
            ['الطلاب الذكور', f"{int(frame['مجموع الطلاب الذكور'].sum()):,}"
             if 'مجموع الطلاب الذكور' in frame.columns else '—'],
            ['الطالبات الإناث', f"{int(frame['مجموع الطلاب الاناث'].sum()):,}"
             if 'مجموع الطلاب الاناث' in frame.columns else '—'],
            ['متوسط الطلاب لكل مدرسة', f"{frame['مجموع الطلاب'].mean():,.0f}"],
            ['عدد المديريات', f"{frame['المديرية'].nunique()}"
             if 'المديرية' in frame.columns else '—'],
            ['عدد المحافظات', f"{frame['المحافظة'].nunique()}"
             if 'المحافظة' in frame.columns else '—'],
        ], columns=['المؤشر', 'القيمة'])
        out['المؤشرات'] = kpi
    return out


def to_excel_bytes(sheets, title="لوحة تحكم المدارس الحكومية"):
    buf = io.BytesIO()
    with xlsxwriter.Workbook(buf, {'in_memory': True}) as wb:
        head = wb.add_format({'bold': True, 'font_size': 16, 'font_color': '#1e40af'})
        sub = wb.add_format({'font_size': 10, 'font_color': '#64748b'})
        tab = wb.add_format({'bold': True, 'bg_color': '#1e40af', 'font_color': 'white',
                             'border': 1, 'align': 'center'})
        cell = wb.add_format({'border': 1, 'border_color': '#e2e8f0'})
        num = wb.add_format({'border': 1, 'border_color': '#e2e8f0', 'num_format': '#,##0'})
        pct = wb.add_format({'border': 1, 'border_color': '#e2e8f0', 'num_format': '#,##0.0'})

        wsc = wb.add_worksheet('التقرير')
        wsc.write(0, 0, title, head)
        wsc.write(1, 0, "عدد الجداول المصدَّرة: %d" % len(sheets), sub)
        wsc.set_column(0, 0, 44)

        for i, (name, d) in enumerate(sheets.items(), start=1):
            # أسماء أوراق Excel لا تقبل : ? * / \
            _sn = str(name)
            for _bad in ':?*/\\':
                _sn = _sn.replace(_bad, '-')
            _sn = _sn[:31] or "ورقة %d" % i
            ws = wb.add_worksheet(_sn)
            for c, cn in enumerate(d.columns):
                ws.write(0, c, str(cn), tab)
                ws.set_column(c, c, min(26, max(11, len(str(cn)) + 4)))
            for r, (_, row) in enumerate(d.iterrows(), start=1):
                for c, v in enumerate(row):
                    if v is None or v is pd.NaT or (isinstance(v, float)
                                                     and v != v):
                        ws.write_blank(r, c, None, cell)
                    elif isinstance(v, bool):
                        ws.write_string(r, c, 'نعم' if v else 'لا', cell)
                    elif isinstance(v, (int, float)) and not isinstance(v, bool):
                        try:
                            fv = float(v)
                        except (TypeError, ValueError):
                            ws.write(r, c, str(v), cell)
                            continue
                        if fv != fv or fv in (float('inf'), float('-inf')):
                            ws.write_blank(r, c, None, cell)
                        elif isinstance(v, int):
                            ws.write_number(r, c, fv, num)
                        else:
                            ws.write_number(r, c, fv, pct)
                    else:
                        ws.write(r, c, str(v), cell)
            ws.freeze_panes(1, 0)
            if len(d):
                ws.autofilter(0, 0, len(d), max(0, len(d.columns) - 1))
    return buf.getvalue()


def to_pdf_bytes(sheets, title="لوحة تحكم المدارس الحكومية"):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                            leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm)
    ss = getSampleStyleSheet()
    h1 = ParagraphStyle('h1', parent=ss['Title'], fontSize=20, textColor=colors.HexColor('#1e40af'))
    h2 = ParagraphStyle('h2', parent=ss['Heading2'], fontSize=13,
                        textColor=colors.HexColor('#0f172a'), spaceBefore=10)
    body = ParagraphStyle('b', parent=ss['Normal'], fontSize=8, alignment=2)  # RTL
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e40af')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#cbd5e1')),
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('FONTSIZE', (0, 1), (-1, -1), 7.5),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1),
         [colors.white, colors.HexColor('#f8fafc')]),
    ])

    story = [Paragraph(ar(title), h1)]
    story.append(Spacer(1, 4 * mm))
    # المؤشرات أولاً
    if 'المؤشرات' in sheets:
        story.append(Paragraph(ar('المؤشرات العامة'), h2))
        d = sheets['المؤشرات']
        rows = [[ar(c) for c in d.columns]] +                [[ar(v) for v in r] for r in d.values.tolist()]
        t = Table(rows, colWidths=[70 * mm, 55 * mm])
        t.setStyle(style)
        story.append(t)
    # ثم كل جدول (كل الصفوف)
    for name, d in sheets.items():
        if name in ('المؤشرات', 'البيانات كاملة'):
            continue
        story.append(Paragraph(ar(name), h2))
        d = d
        if not len(d):
            continue
        rows = [[ar(c) for c in d.columns]] +                [[ar(v) for v in r] for r in d.values.tolist()]
        ncol = len(d.columns)
        t = Table(rows, colWidths=[min(60 * mm, (270 * mm) / max(ncol, 1))] * ncol)
        t.setStyle(style)
        story.append(t)
    doc.build(story)
    return buf.getvalue()


def to_zip_bytes(sheets):
    buf = io.BytesIO()
    safe = {'المديرية': 'directorates', 'المحافظة': 'governorates',
            'الجنس': 'genders', 'نوع التعليم': 'education', 'الملكية': 'ownership',
            'أدنى صف': 'lowest_grade', 'أعلى صف': 'highest_grade',
            'فترة المؤسسة': 'period', 'المؤشرات': 'summary'}
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for i, (name, d) in enumerate(sheets.items(), start=1):
            stem = safe.get(name)
            fname = ("%02d_%s.csv" % (i, stem) if stem
                     else "%02d_%s.csv" % (i, name.replace('/', '-').replace(' ', '_')))
            z.writestr(fname, d.to_csv(index=False).encode('utf-8-sig'))
    return buf.getvalue()


def to_json_bytes(frame):
    return json.dumps({
        'البرنامج': 'لوحة تحكم المدارس الحكومية',
        'عدد المدارس': int(len(frame)),
        'إجمالي الطلاب': int(pd.to_numeric(frame['مجموع الطلاب'],
                                          errors='coerce').sum())
        if 'مجموع الطلاب' in frame.columns else None,
        'البيانات': frame.where(pd.notna(frame), None).to_dict(orient='records'),
    }, ensure_ascii=False, indent=1).encode('utf-8')


def show_df(d, title='جدول', fname='جدول', **kw):
    """عرض الجدول مع زر لتصديره بالصيغة التي تختارها — من أي تبويب."""
    kw.setdefault('width', 'stretch')
    kw.setdefault('use_container_width', True)
    st.dataframe(d, **kw)
    show_df._n = getattr(show_df, '_n', 0) + 1
    _n = show_df._n
    _fmt = st.radio('صيغة التصدير:', ['Excel', 'PDF', 'CSV'],
                    horizontal=True, key='showdf_fmt_%d' % _n)
    _safe = (fname or title).replace(' ', '_')
    if _fmt == 'CSV':
        st.download_button('⬇️ تنزيل «' + title + '» (CSV)',
                           data=d.to_csv(index=False).encode('utf-8-sig'),
                           file_name=_safe + '.csv', mime='text/csv',
                           key='showdf_dl_%d' % _n)
    elif _fmt == 'PDF':
        try:
            _pdf = to_pdf_bytes({title: d}, title=title)
        except Exception:
            _pdf = b''
        st.download_button('⬇️ تنزيل «' + title + '» (PDF)',
                           data=_pdf,
                           file_name=_safe + '.pdf', mime='application/pdf',
                           key='showdf_dl_%d' % _n)
    else:
        try:
            _buf = io.BytesIO()
            with pd.ExcelWriter(_buf, engine='xlsxwriter') as _w:
                d.to_excel(_w, index=False, sheet_name=str(title)[:31])
            _xlsx_data = _buf.getvalue()
        except Exception:
            _xlsx_data = b''
        st.download_button('⬇️ تنزيل «' + title + '» (Excel)',
                           data=_xlsx_data,
                           file_name=_safe + '.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           key='showdf_dl_%d' % _n)


def directorate_table(frame, title='جدول كل المديريات', sort_by='عدد المدارس'):
    """
    جدول كامل لجميع المديريات: عدد مدارسها وطلابها وطاقتها.
    مرتب دائماً تنازلياً حسب العمود المختار.
    """
    if 'المديرية' not in frame.columns:
        return None
    g = frame.groupby('المديرية', dropna=False).size().reset_index(name='عدد المدارس')
    for c, lab in (('مجموع الطلاب', 'إجمالي الطلاب'),
                   ('مجموع الطلاب الذكور', 'الذكور'),
                   ('مجموع الطلاب الاناث', 'الإناث'),
                   ('الطاقة الاستيعابية', 'إجمالي الطاقة')):
        if c in frame.columns:
            _s = frame.groupby('المديرية', dropna=False)[c].sum(min_count=1)
            g[lab] = g['المديرية'].map(_s)
    for lab in list(g.columns)[1:]:
        if g[lab].dtype.kind == 'f':
            g[lab] = g[lab].round(0)

    # ---- اختيار عمود الترتيب (تنازلي دائماً) ----
    _opts = [c for c in g.columns if c != 'المديرية']
    if sort_by not in _opts:
        sort_by = 'عدد المدارس'
    _sc1, _sc2 = st.columns([2, 3])
    with _sc1:
        sort_by = st.selectbox(
            "ترتيب تنازلي حسب:", _opts, index=_opts.index(sort_by),
            key="dir_sort_%s" % title[:18])

    g = g.sort_values([sort_by, 'المديرية'], ascending=[False, True]) \
            .reset_index(drop=True)
    g.insert(0, 'الترتيب', range(1, len(g) + 1))

    st.markdown(f"###### {title} ({len(g):,} مديرية — مرتبة تنازلياً حسب "
                f"«{sort_by}»)")
    show_df(g, title, 'جدول_مديريات')
    return g



with col1:
    st.markdown('<span class="filters-anchor"></span>', unsafe_allow_html=True)
    st.subheader("🔍 الفلاتر")

    directorates = ['الكل'] + sorted(df['المديرية'].dropna().unique().tolist())
    selected_directorate = st.selectbox("المديرية", directorates)

    if selected_directorate != 'الكل':
        governorates = ['الكل'] + sorted(df[df['المديرية'] == selected_directorate]['المحافظة'].dropna().unique().tolist())
    else:
        governorates = ['الكل'] + sorted(df['المحافظة'].dropna().unique().tolist())
    selected_governorate = st.selectbox("المحافظة", governorates)

    genders = ['الكل'] + sorted(df['الجنس'].dropna().unique().tolist())
    selected_gender = st.selectbox("الجنس", genders)

    edu_types = ['الكل'] + sorted(df['نوع التعليم'].dropna().unique().tolist())
    selected_edu_type = st.selectbox("نوع التعليم", edu_types)

    ownerships = ['الكل'] + sorted(df['الملكية'].dropna().unique().tolist())
    selected_ownership = st.selectbox("الملكية", ownerships)

    st.markdown("---")
    if st.button("🔄 تطبيق الفلاتر", type="primary"):
        st.rerun()

# تطبيق الفلاتر
filtered_df = df.copy()
if selected_directorate != 'الكل':
    filtered_df = filtered_df[filtered_df['المديرية'] == selected_directorate]
if selected_governorate != 'الكل':
    filtered_df = filtered_df[filtered_df['المحافظة'] == selected_governorate]
if selected_gender != 'الكل':
    filtered_df = filtered_df[filtered_df['الجنس'] == selected_gender]
if selected_edu_type != 'الكل':
    filtered_df = filtered_df[filtered_df['نوع التعليم'] == selected_edu_type]
if selected_ownership != 'الكل':
    filtered_df = filtered_df[filtered_df['الملكية'] == selected_ownership]

# ====== محرك البحث (عمود مستقل على اليمين) ======
# ====== لوحة التصدير (داخل عمود الفلاتر) ======
with col1:
    st.markdown("---")
    st.markdown("#### \U0001F4E6 تصدير")
    st.caption("يصدّر ما تعرضه الفلاتر أعلاه — %s مدرسة"
               % f"{len(filtered_df):,}")

    _sheets = build_sheets(filtered_df)
    _stamp = "__".join([str(selected_directorate)[:6], str(selected_gender)[:6],
                        str(selected_ownership)[:6]])

    st.download_button(
        "\U0001F4CA Excel \u2014 كل الجداول (xlsx)",
        data=to_excel_bytes(_sheets),
        file_name="لوحة_تحكم_المدارس_%s.xlsx" % _stamp,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.download_button(
        "\U0001F4C4 PDF \u2014 تقرير كامل",
        data=to_pdf_bytes(_sheets),
        file_name="تقرير_المدارس_%s.pdf" % _stamp,
        mime="application/pdf",
        use_container_width=True,
    )
    st.download_button(
        "\U0001F5DC ZIP \u2014 ملف CSV لكل جدول",
        data=to_zip_bytes(_sheets),
        file_name="جداول_المدارس_%s.zip" % _stamp,
        mime="application/zip",
        use_container_width=True,
    )
    st.download_button(
        "\U0001F4C3 JSON \u2014 البيانات الخام",
        data=to_json_bytes(filtered_df),
        file_name="بيانات_المدارس_%s.json" % _stamp,
        mime="application/json",
        use_container_width=True,
    )

    with st.expander("تصدير جدول واحد (Excel)"):
        _pick = st.selectbox("الجدول", list(_sheets.keys()), key="exp_pick")
        _b1 = io.BytesIO()
        with pd.ExcelWriter(_b1, engine='xlsxwriter') as _w:
            _sheets[_pick].to_excel(_w, index=False, sheet_name=str(_pick)[:31])
        st.download_button(
            "تنزيل %s.xlsx" % _pick,
            data=_b1.getvalue(),
            file_name=("%s.xlsx" % _pick.replace(' ', '_').replace(':', '')),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.caption("الأوراق المصدَّرة: %d — %s"
               % (len(_sheets), " · ".join(list(_sheets)[:6])))


with col3:
    st.markdown('<span class="search-anchor"></span>', unsafe_allow_html=True)
    st.subheader("🔎 محرك البحث")

    st.markdown("**نوع البحث:**")
    search_by_directorate = st.checkbox("المديرية", value=True, key="cb_directorate")
    search_by_name = st.checkbox("اسم المدرسة", key="cb_name")
    search_by_id = st.checkbox("الرقم الوطني", key="cb_id")
    search_by_coords = st.checkbox("الإحداثيات", key="cb_coords")

    search_types = []
    if search_by_name:
        search_types.append("اسم المدرسة")
    if search_by_directorate:
        search_types.append("المديرية")
    if search_by_id:
        search_types.append("الرقم الوطني")
    if search_by_coords:
        search_types.append("الإحداثيات")

    search_results = pd.DataFrame()

    if "المديرية" in search_types:
        directorate_options = sorted(filtered_df['المديرية'].dropna().unique().tolist())
        selected_directorates = st.multiselect("اختر المديرية:", directorate_options, key="directorate_select")
        if selected_directorates:
            filtered_df = filtered_df[filtered_df['المديرية'].isin(selected_directorates)]

    if "اسم المدرسة" in search_types:
        school_options = filtered_df['اسم المدرسة'].dropna().unique().tolist()
        selected_school = st.selectbox("ابحث واختر المدرسة:", school_options, key="school_select")
        if selected_school:
            search_results = filtered_df[filtered_df['اسم المدرسة'] == selected_school]

    if "الرقم الوطني" in search_types:
        search_query = st.text_input("أدخل الرقم الوطني:", key="search_id")
        if search_query:
            id_results = filtered_df[filtered_df['الرقم الوطني'].astype(str).str.contains(search_query, na=False)]
            search_results = pd.concat([search_results, id_results]).drop_duplicates()

    if "الإحداثيات" in search_types:
        search_lat = st.text_input("خط العرض (y):", key="search_lat")
        search_lon = st.text_input("خط الطول (x):", key="search_lon")
        if search_lat and search_lon:
            try:
                lat_val = float(search_lat)
                lon_val = float(search_lon)
                coord_results = filtered_df[
                    (filtered_df['y'].astype(float, errors='ignore') - lat_val).abs() < 0.01 &
                    (filtered_df['x'].astype(float, errors='ignore') - lon_val).abs() < 0.01
                ]
                search_results = pd.concat([search_results, coord_results]).drop_duplicates()
            except:
                st.error("إحداثيات غير صحيحة")

    # عرض نتائج البحث
    if len(search_results) > 0:
        st.success(f"تم العثور على {len(search_results)} مدرسة")
        for idx, row in search_results.iterrows():
            st.markdown(f"**{row['اسم المدرسة']}** — {row['المديرية']}")
        filtered_df = search_results.copy()
    elif 'search_query' in locals() and search_query:
        st.warning("لم يتم العثور على نتائج")

# ====== قسم الدراسات (عمود مستقل تحت محرك البحث) ======
with col3:
    st.markdown("---")
    st.markdown("""
    <style>
    [data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stMarkdown p,
    [data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stMarkdown li,
    [data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stMarkdown h4 {
        color: #ffffff !important;
        font-weight: 700;
        font-size: 1rem;
    }
    [data-testid="stHorizontalBlock"]:has(.search-anchor) > [data-testid="stColumn"]:last-child .stMarkdown strong {
        color: #ffffff !important;
        font-weight: 800;
    }
    </style>
    """, unsafe_allow_html=True)
# ============ حساب وعرض بطاقات المؤشرات بعد البحث ============
total_schools = len(filtered_df)
total_students = filtered_df['مجموع الطلاب'].sum()
total_male = filtered_df['مجموع الطلاب الذكور'].sum()
total_female = filtered_df['مجموع الطلاب الاناث'].sum()
avg_students = filtered_df['مجموع الطلاب'].mean()
total_directorates = filtered_df['المديرية'].nunique()

st.markdown(f'''
<div class="kpis">
  <div class="kpi a"><span class="ic">🏫</span><div class="v">{total_schools:,}</div><div class="t">إجمالي المدارس</div></div>
  <div class="kpi b"><span class="ic">👥</span><div class="v">{int(total_students):,}</div><div class="t">إجمالي الطلاب</div></div>
  <div class="kpi c"><span class="ic">👦</span><div class="v">{int(total_male):,}</div><div class="t">الطلاب الذكور</div></div>
  <div class="kpi d"><span class="ic">👧</span><div class="v">{int(total_female):,}</div><div class="t">الطالبات الإناث</div></div>
  <div class="kpi e"><span class="ic">📊</span><div class="v">{int(avg_students)}</div><div class="t">متوسط الطلاب لكل مدرسة</div></div>
  <div class="kpi f"><span class="ic">🏢</span><div class="v">{total_directorates}</div><div class="t">عدد المديريات</div></div>
</div>
''', unsafe_allow_html=True)


# ============================================================
#  بناء خريطة المدارس — مخزَّنة مؤقتاً (st.cache_data)
# 站立performance: نقاط على canvas + أسماء محدودة فقط
# ============================================================
import math as _math
import json as _json

MAP_NAME_FONT_PX = 10      # ← حجم خط اسم المدرسة على الخريطة
MAP_NAME_WEIGHT = 600      # ← سماكة الخط
MAP_LABEL_BASE = 250       # ← عدد الأسماء الظاهرة عند تكبير 8 (على مستوى الدولة)
MAP_LABEL_PER_ZOOM = 500   # ← أسماء إضافية لكل درجة تكبير
MAP_MAX_NAMES = 400        # (قديم — لم يعد مستخدماً)
MAP_MAX_HEAT = 1200        # ← أقصى عدد نقاط الخريطة الحرارية
MAP_MAX_DOTS = 6000        # ← أقصى عدد نقاط مدارس

GENDER_COLORS = [('مختلط', '#1a9850'), ('اناث', '#e75480'), ('إناث', '#e75480'),
                 ('انث', '#e75480'), ('ذكور', '#2166ac')]


def _gender_color(v):
    s = str(v).strip().lower()
    for key, col in GENDER_COLORS:
        if key in s:
            return col
    return '#666666'


@st.cache_data(show_spinner=False, max_entries=6)
def build_school_map(points, gender, students, names, show_names, show_heat,
                     name_font, name_weight):
    """
    points   : ((lat, lon), ...)
    gender   : (str, ...)  — لتلوين النقاط
    students : (float, ...) — لحجم النقطة وترتيب الأسماء
    names    : (str, ...)
    تُرسم النقاط كطبقة GeoJson واحدة (سريعة جداً) بدل آلاف العلامات.
    """
    m = folium.Map(location=[31.5, 36.0], zoom_start=8,
                   tiles='OpenStreetMap', prefer_canvas=True)

    # 1) الخريطة الحرارية (مخفّفة بعدّ النقاط)
    if show_heat and points:
        heat = list(points)[:MAP_MAX_HEAT]
        plugins.HeatMap(
            heat,
            min_opacity=0.30, max_zoom=10, radius=15, blur=10,
            gradient={0.4: 'blue', 0.65: 'lime', 0.8: 'yellow', 1: 'red'}
        ).add_to(m)

    # 2) كل المدارس كطبقة GeoJson واحدة لكل جنس (بدون آلاف عناصر DOM)
    buckets = {}
    for (lat, lon), g, s, nm in zip(points, gender, students, names):
        if lat is None or lon is None:
            continue
        key = _gender_color(g)
        try:
            rad = 2.5 + min(float(s or 0), 1200) / 400.0
        except (TypeError, ValueError):
            rad = 3.0
        buckets.setdefault(key, []).append({
            "type": "Feature",
            "geometry": {"type": "Point", "coordinates": [lon, lat]},
            "properties": {"name": nm, "radius": rad},
        })

    for color, feats in buckets.items():
        folium.GeoJson(
            {"type": "FeatureCollection", "features": feats},
            name=f"{_gender_label(color)} ({len(feats):,})",
            marker=folium.CircleMarker(
                radius=3.2, color=color, weight=1, opacity=0.85,
                fill=True, fill_color=color, fill_opacity=0.75),
            tooltip=folium.GeoJsonTooltip(
                fields=["name"], aliases=["المدرسة: "], sticky=False),
            show=True,
        ).add_to(m)

    # 3) الأسماء — طبقة حيّة: تُرسم لما هو داخل الشاشة فقط، وتُحدَّث مع كل تحريك.
    #    هذا يجعل «كل» الأسماء ظاهرة دون إنشاء 4,000 عنصر في الصفحة دفعة واحدة.
    if show_names and names:
        _order = sorted(range(len(students)), key=lambda i: -(students[i] or 0))
        _payload = _json.dumps(
            [[round(points[i][0], 5), round(points[i][1], 5), str(names[i])]
             for i in _order],
            ensure_ascii=False, separators=(',', ':'))
        _halo = ("1px 1px 2px white, -1px -1px 2px white, "
                 "1px -1px 2px white, -1px 1px 2px white")
        # خطّاف Leaflet الرسمي: يُستدعى أثناء إنشاء الخريطة، فنأجّل
        # التهيئة إلى ما بعد انتهاء الإنشاء (حتى لا تُقاس حا��ة الخريطة = 0).
        _js = """
        (function () {
          var SCHOOLS = __PAYLOAD__;
          var BASE = __BASE__, PER = __PER__;
          if (!SCHOOLS.length || !window.L || !L.Map || !L.Map.addInitHook) { return; }

          // ملاحظة: Leaflet ينادي الخطّاف بـ fn.call(this) — بلا وسائط،
          // فالخريطة هي 'this' وليس معاملاً.
          L.Map.addInitHook(function () {
            var map = this;
            window.setTimeout(function () {
              try {
                var layer = L.layerGroup();
                map.addLayer(layer);

                function draw() {
                  layer.clearLayers();
                  var b = map.getBounds();
                  var z = map.getZoom();
                  var cap = Math.min(SCHOOLS.length, BASE + Math.max(0, z - 8) * PER);
                  var n = 0;
                  for (var i = 0; i < SCHOOLS.length && n < cap; i++) {
                    var s = SCHOOLS[i];
                    if (!b.contains([s[0], s[1]])) { continue; }
                    n++;
                    layer.addLayer(L.marker([s[0], s[1]], {
                      interactive: false,
                      keyboard: false,
                      icon: L.divIcon({
                        className: 'school-label',
                        html: '<div style="font-size:__FONT__px;font-weight:__WEIGHT__;' +
                              'color:#111;text-shadow:__HALO__;white-space:nowrap;' +
                              'line-height:1.1;pointer-events:none;">' + s[2] + '</div>',
                        iconSize: [200, 18],
                        iconAnchor: [100, 9]
                      })
                    }));
                  }
                  window.__labelCount = n;
                }

                map.on('zoomend', draw);
                map.on('moveend', draw);
                draw();
                window.setTimeout(draw, 500);
              } catch (err) {
                window.__labelError = String(err);
              }
            }, 0);
          });
        })();
        """
        _js = (_js.replace('__PAYLOAD__', _payload)
                 .replace('__FONT__', str(name_font))
                 .replace('__WEIGHT__', str(name_weight))
                 .replace('__HALO__', _halo)
                 .replace('__BASE__', str(MAP_LABEL_BASE))
                 .replace('__PER__', str(MAP_LABEL_PER_ZOOM)))
        _js = _js.strip()
        assert '<script' not in _js.lower(), 'لا ضع وسم script يدوياً'
        m.get_root().script.add_child(folium.Element(_js))

    # 4) طبقة الستلايت
    folium.TileLayer(
        tiles='https://server.arcgisonline.com/ArcGIS/rest/services/'
              'World_Imagery/MapServer/tile/{z}/{y}/{x}',
        attr='Esri', name='Satellite', overlay=True, control=True
    ).add_to(m)

    folium.LayerControl().add_to(m)
    return m.get_root().render()


_GENDER_LABELS = {'#2166ac': 'ذكور', '#e75480': 'إناث',
                  '#1a9850': 'مختلطة', '#666666': 'غير محدد'}


def _gender_label(color):
    return _GENDER_LABELS.get(color, 'أخرى')



# ============================================================
#  المسافة بين مدارس — Haversine + مسار الشارع عبر OSRM
# ============================================================
ROUTE_CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'routes_cache')
RADIUS_KM = 6371.0


def straight_distance_km(lat1, lon1, lat2, lon2):
    """المسافة المستقيمة بالكيلومتر."""
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = p2 - p1
    dl = math.radians(lon2) - math.radians(lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return RADIUS_KM * 2 * math.asin(math.sqrt(a))


@st.cache_data(show_spinner=False, max_entries=4000)
def road_route(lat1, lon1, lat2, lon2):
    """مسار الشارع. يعيد [المسافة م, الزمن ث, نقاط [[lat,lon],...]] أو [None,None,خطأ]."""
    os.makedirs(ROUTE_CACHE_DIR, exist_ok=True)
    key = "%.5f_%.5f_%.5f_%.5f" % (lat1, lon1, lat2, lon2)
    cpath = os.path.join(ROUTE_CACHE_DIR, key + ".json")
    if os.path.isfile(cpath):
        try:
            with io.open(cpath, encoding="utf-8") as fh:
                return _json.load(fh)
        except Exception:
            pass
    url = ("https://router.project-osrm.org/route/v1/driving/"
           "%.6f,%.6f;%.6f,%.6f?overview=full&geometries=geojson"
           % (lon1, lat1, lon2, lat2))
    try:
        r = requests.get(url, timeout=30,
                         headers={'User-Agent': 'schools-dashboard/1.0'})
        r.raise_for_status()
        rt = r.json()['routes'][0]
        coords = rt['geometry']['coordinates']        # [(lon, lat), ...]
        out = [float(rt['distance']), float(rt['duration']),
               [[c[1], c[0]] for c in coords]]
    except Exception as exc:
        out = [None, None, str(exc)]
    try:
        with io.open(cpath, "w", encoding="utf-8") as fh:
            _json.dump(out, fh)
    except Exception:
        pass
    return out



# ============================================================
#  إمكانية نقل المدارس المستأجرة إلى مدارس الملك القريبة
# ============================================================
RENTED_TO_OWNED_RADIUS_KM = 10.0


@st.cache_data(show_spinner=False)
def rented_to_owned_pairs(df_in, radius_km=RENTED_TO_OWNED_RADIUS_KM,
                          grade_lo=None, grade_hi=None, match_mode='range',
                          require_cover=True, require_grades=True,
                          allowed_genders=None,
                          exclude_second_period=True, require_capacity=True,
                          king_exclude_codes=(-1, -2), require_gender_match=True):
    """
    أزواج (مستأجر → ملك) ضمن نصف القطر مع توفر الطاقة.

    grade_lo / grade_hi : رمز أدنى صف وأعلى صف المطلوب للمدرسة المستأجرة
                          (None = بلا قيد)
    match_mode : 'range' → أدنى صف المستأجرة ≥ grade_lo وأعلى صفها ≤ grade_hi
                'exact' → أدنى صفها = grade_lo وأعلى صفها = grade_hi بالضبط
    require_grades : تفعيل قيد الصفوف من عدمه
    allowed_genders : مجموعة أجنس مقبولة للمدارس المستأجرة والملك معاً
                      (None أو فارغة = بلا قيد). مثال {'ذكور'} أو {'إناث','مختلطة'}
    exclude_second_period : استبعاد مدارس «فترة ثانية» من مدارس الملك
    require_capacity : يجب أن تكون الطاقة الاستيعابية للملك
                       ≥ عدد طلاب المدرسة المستأجرة
    king_exclude_codes : أكواد الصفوف المستثناة من مدارس الملك
                         (-1 روضة أولى ، -2 روضة ثانية). فارغ = بلا استثناء.
    require_cover  : لازم مدرسة الملك تغطي مدى صفوف المستأجرة
                     (ملك أدنى ≤ مستأجرة أدنى  و  ملك أعلى ≥ مستأجرة أعلى)
    """
    need = ['الملكية', 'الطاقة الاستيعابية', 'مجموع الطلاب', 'x', 'y']
    miss = [c for c in need if c not in df_in.columns]
    for c in ('أدنى صف', 'أعلى صف', 'الجنس', 'فترة المؤسسة'):
        if c not in need:
            need.append(c)
    miss = [c for c in need if c not in df_in.columns]
    if miss:
        return None, miss

    # ملاحظة: «ملك ومستأجر» لا يصحّ أن تُحسب في الطرفين معاً
    _own_flag = df_in['الملكية'].astype(str)
    ren = df_in[_own_flag.str.contains('مستأجر', na=False)
                & ~_own_flag.str.contains('ملك', na=False)].copy()
    own = df_in[_own_flag.str.contains('ملك', na=False)
                & ~_own_flag.str.contains('مستأجر', na=False)].copy()

    # استثناء مدارس «فترة ثانية» من مدارس الملك (تُقارن الفترات أعلاه)
    if exclude_second_period and ('فترة المؤسسة' in own.columns):
        _per = own['فترة المؤسسة'].astype(str).str.replace(r'\s+', '', regex=True)
        own = own[~_per.str.contains('ثانية', na=False)].copy()

    # استثناء مدارس الملك التي «أدنى صف» فيها ضمن الأكواد المستثناة
    if king_exclude_codes and ('أدنى صف' in own.columns):
        _king_lo = own['أدنى صف'].map(grade_number)
        own = own[~_king_lo.isin(set(king_exclude_codes))].copy()
    if len(ren) == 0 or len(own) == 0:
        return None, []

    for d in (ren, own):
        d['مجموع الطلاب'] = pd.to_numeric(d['مجموع الطلاب'], errors='coerce')
        d['الطاقة الاستيعابية'] = pd.to_numeric(d['الطاقة الاستيعابية'], errors='coerce')
        d['x'] = pd.to_numeric(d['x'], errors='coerce')
        d['y'] = pd.to_numeric(d['y'], errors='coerce')
        d = d.dropna(subset=['x', 'y'])

    if len(ren) == 0 or len(own) == 0:
        return None, []

    r_lat = ren['y'].tolist(); r_lon = ren['x'].tolist()
    o_lat = own['y'].tolist(); o_lon = own['x'].tolist()
    r_stu = ren['مجموع الطلاب'].tolist()
    o_cap = own['الطاقة الاستيعابية'].tolist()
    r_lo = ren['أدنى صف'].map(grade_number).tolist()
    r_hi = ren['أعلى صف'].map(grade_number).tolist()
    o_lo = own['أدنى صف'].map(grade_number).tolist()
    o_hi = own['أعلى صف'].map(grade_number).tolist()
    r_gen = ren['الجنس'].astype(str).str.strip().tolist()
    o_gen = own['الجنس'].astype(str).str.strip().tolist()
    r_per = ren['فترة المؤسسة'].astype(str).str.strip().tolist() \
        if 'فترة المؤسسة' in ren.columns else [''] * len(ren)
    o_per = own['فترة المؤسسة'].astype(str).str.strip().tolist() \
        if 'فترة المؤسسة' in own.columns else [''] * len(own)
    r_name = ren.get('اسم المدرسة', pd.Series([''] * len(ren))).astype(str).tolist()
    o_name = own.get('اسم المدرسة', pd.Series([''] * len(own))).astype(str).tolist()
    r_gov = ren.get('المديرية', pd.Series([''] * len(ren))).astype(str).tolist()

    rows = []
    for i in range(len(r_lat)):
        lat1, lon1 = math.radians(r_lat[i]), math.radians(r_lon[i])
        for j in range(len(o_lat)):
            lat2, lon2 = math.radians(o_lat[j]), math.radians(o_lon[j])
            dp, dl = lat2 - lat1, lon2 - lon1
            a = math.sin(dp / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * \
                math.sin(dl / 2) ** 2
            d = RADIUS_KM * 2 * math.asin(math.sqrt(a))
            # لا زوج بين مدرسة لنفسها (نفس الاسم أو نفس الموقع)
            if (o_name[j] == r_name[i]) or (r_lat[i] == o_lat[j] and r_lon[i] == o_lon[j]):
                continue
            if not (d <= radius_km and d > 0.01):
                continue
            # ---- شرط توفر السعة في مدرسة الملك ----
            if require_capacity and (o_cap[j] or 0) < (r_stu[i] or 0):
                continue

            # ---- شرط جنس المدرسة ----
            if allowed_genders:
                if r_gen[i] not in allowed_genders or o_gen[j] not in allowed_genders:
                    continue
            # ---- مراعاة الجنس بين المستأجرة والملك ----
            if require_gender_match:
                _rg = str(r_gen[i]).strip()
                _og = str(o_gen[j]).strip()
                _is_boys = 'ذكور' in _rg
                _is_girls = ('انث' in _rg) or ('إناث' in _rg)
                _is_mixed = 'مختلطة' in _rg or 'مختلط' in _rg
                if _is_boys and not _is_mixed:
                    # بنين ← بنين فقط
                    if ('ذكور' not in _og) or ('مختلط' in _og) or ('انث' in _og) or ('إناث' in _og):
                        continue
                else:
                    # مختلطة أو إناث ← مختلطة أو إناث فقط
                    _og_ok = ('مختلط' in _og) or ('انث' in _og) or ('إناث' in _og)
                    if not _og_ok:
                        continue

            # ---- شروط الصفوف (بالأكواد) ----
            if require_grades and grade_lo is not None and grade_hi is not None:
                if r_lo[i] is None or r_hi[i] is None:
                    continue
                if match_mode == 'exact':
                    if r_lo[i] != grade_lo or r_hi[i] != grade_hi:
                        continue
                else:
                    if r_lo[i] < grade_lo or r_hi[i] > grade_hi:
                        continue
            if require_cover:
                if o_lo[j] is None or o_hi[j] is None:
                    continue
                # المدرسة الملك يجب أن تبدأ عند أدنى صف أو قبله،
                # وتنتهي عند أعلى صف أو بعده
                if o_lo[j] > r_lo[i] or o_hi[j] < r_hi[i]:
                    continue

            rows.append({
                'المدرسة المستأجرة': r_name[i],
                'المديرية': r_gov[i],
                'المدرسة الملك': o_name[j],
                'جنس المستأجرة': r_gen[i],
                'جنس الملك': o_gen[j],
                'فترة مؤسسة المستأجرة': r_per[i],
                'فترة مؤسسة الملك': o_per[j],
                'أدنى صف المستأجرة': r_lo[i],
                'أعلى صف المستأجرة': r_hi[i],
                'أدنى صف الملك': o_lo[j],
                'أعلى صف الملك': o_hi[j],
                'عدد طلاب المستأجرة': r_stu[i],
                'الطاقة الاستيعابية للملك': o_cap[j],
                'المسافة (كم)': round(d, 2),
                'فائض الطاقة (Places)': round((o_cap[j] or 0) - (r_stu[i] or 0), 0),
                '_ry': r_lat[i], '_rx': r_lon[i],
                '_oy': o_lat[j], '_ox': o_lon[j],
            })
    if not rows:
        return pd.DataFrame(columns=['المدرسة المستأجرة', 'المديرية', 'المدرسة الملك',
                                    'جنس المستأجرة', 'جنس الملك',
                                    'فترة مؤسسة المستأجرة', 'فترة مؤسسة الملك',
                                    'أدنى صف المستأجرة', 'أعلى صف المستأجرة', 'أدنى صف الملك', 'أعلى صف الملك',
                                    'عدد طلاب المستأجرة', 'الطاقة الاستيعابية للملك',
                                    'المسافة (كم)', 'فائض الطاقة (Places)']), []
    return pd.DataFrame(rows), []


def render_rented_to_owned(df_pairs, radius_km):
    """يرسم نتائج نقل المستأجر إلى الملك."""
    st.markdown("#### إمكانية نقل المدارس المستأجرة إلى مدارس الملك")
    st.caption("نبحث عن مدرسة ملك قريبة تستوعب طلاب المدرسة المستأجرة "
               f"ضمن نصف قطر {radius_km:g} كم، بطاقة استيعابية كافية.")

    # ====== فلتر المديرية (خاص بهذه الدراسة) ======
    _dirs = ['الكل'] + sorted({str(v) for v in filtered_df.get('المديرية', pd.Series(dtype=str)).dropna().unique()})
    _sel_dirs = st.multiselect("◼ المديرية (فلتر خاص بالدراسة)", _dirs,
                               default=['الكل'], key="r2o_dirs")
    if _sel_dirs and 'الكل' not in _sel_dirs:
        _rdf = filtered_df[filtered_df['المديرية'].isin(_sel_dirs)].copy()
    else:
        _rdf = filtered_df.copy()

    for _lab, _col in (('القضاء', 'القضاء'), ('اللواء', 'اللواء'), ('البلدية', 'البلدية')):
        if _col in _rdf.columns:
            _opts = ['الكل'] + sorted({str(v) for v in _rdf[_col].dropna().unique()})
            _sel = st.multiselect("◼ %s (فلتر خاص بالدراسة)" % _lab, _opts,
                                  default=['الكل'], key="r2o_%s" % _col)
            if _sel and 'الكل' not in _sel:
                _rdf = _rdf[_rdf[_col].isin(_sel)].copy()

    # ====== شروط النقل: كل شرط في سطر مستقل (سهل القراءة) ======
    st.markdown("**⚙︎ شروط النقل**")

    radius_km = st.number_input("◼ نصف القطر (كم)", min_value=1.0, max_value=50.0,
                                value=float(RENTED_TO_OWNED_RADIUS_KM), step=0.5,
                                key="r2o_radius")

    require_grades = st.checkbox(
        "◼ اشتراط صفوف المستأجرة", value=True, key="r2o_req_grades",
        help="عند التفعيل: يجب أن يكون «أدنى صف» و«أعلى صف» للمستأجرة "
             "ضمن الحدود المحددة أدناه.")
    st.caption("يحدد أي صفوف تُسمح بنقلها للمدرسة المستأجرة.")

    require_cover = st.checkbox(
        "◼ المدرسة الملك يجب أن تغطي صفوف المستأجرة", value=True, key="r2o_req_cover",
        help="ملك أدنى ≤ مستأجرة أدنى  و  ملك أعلى ≥ مستأجرة أعلى.")
    st.caption("يمنع اقتراح مدرسة ملك لا تغطي كامل مدى صفوف المستأجرة.")

    require_gender = st.checkbox(
        "◼ اشتراط جنس المدرسة", value=True, key="r2o_req_gender",
        help="عند التفعيل: لا يُقبل نقل مدرسة إلى مدرسة ملك خارج "
             "الجنس المحدد بالشرطين أدناه.")

    exclude_second_period = st.checkbox(
        "◼ استثناء مدارس الفترة الثانية من مدارس الملك", value=True,
        key="r2o_no_second",
        help="لا يُقترح النقل إلى مدرسة ملك نظامها «فترة ثانية».")
    st.caption("مدارس «فترة ثانية» تُستبعد من قائمة مدارس الملك.")

    require_capacity = st.checkbox(
        "◼ اشتراط توفر السعة في مدرسة الملك", value=True, key="r2o_req_cap",
        help="يُقبل الاقتراح فقط إذا كان عدد طلاب المدرسة المستأجرة "
             "أقل من الطاقة الاستيعابية لمدرسة الملك.")
    if require_capacity:
        st.caption("القاعدة: طلاب المستأجرة ≤ الطاقة الاستيعابية للملك.")

    king_exclude = st.checkbox(
        "◼ استثناء مدارس الملك ذات الصفوف ‎-1 و ‎-2", value=True,
        key="r2o_no_kg",
        help="لا تُقترح مدارس الملك التي يبدأ أدنى صف فيها من "
             "روضة ثانية (-2) أو روضة أولى (-1).")
    king_exclude_codes = (-1, -2) if king_exclude else ()
    if king_exclude:
        st.caption("المستثنى من مدارس الملك: **روضة ثانية (-2)** "
                   "و**روضة أولى (-1)**.")

    # ---------- شرط الجنس: شرطان في سطرين ----------
    allowed_genders = None
    if require_gender:
        st.markdown("**◼ الشرط الأول والثاني (الجنس)**")
        cond_male = st.checkbox("◼ الشرط الأول: مدارس الذكور", value=True,
                                key="r2o_g_male")
        cond_female = st.checkbox("◼ الشرط الثاني: مدارس الإناث أو المختلطة",
                                  value=True, key="r2o_g_female")
        allowed_genders = set()
        if cond_male:
            allowed_genders.add('ذكور')
        if cond_female:
            allowed_genders |= {'إناث', 'مختلطة'}
        if not allowed_genders:
            st.info("لم تُحدَّد أي شرط جنس — سيُطبَّق شرط الجنس على كل الأصناف.")
        else:
            st.caption("الأصناف المقبولة (للمستأجرة والملك معاً): **"
                       + "** و **".join(sorted(allowed_genders)) + "**")

    # ---- جدول الأكواد ----
    with st.expander("جدول أكواد الصفوف", expanded=False):
        show_df(pd.DataFrame(GRADE_TABLE, columns=['الصف', 'الرمز']),
                'جدول أكواد الصفوف', 'جدول_أكواد_الصفوف')

    grade_lo = grade_hi = None
    match_mode = 'range'
    if require_grades:
        # ---------- تحديد أدنى صف وأعلى صف بنفسي ----------
        st.markdown("**حدّد مستوى الصف المطلوب**")
        g1, g2, g3 = st.columns([2, 2, 2])
        with g1:
            pick_lo = st.selectbox("أدنى صف (للمدرسة المستأجرة)", GRADE_SORT,
                                   index=0, key="r2o_lo",
                                   help="أدنى مستوى صف يُسمح بنقله.")
        with g2:
            pick_hi = st.selectbox("أعلى صف (للمدرسة المستأجرة)", GRADE_SORT,
                                   index=len(GRADE_SORT) - 1, key="r2o_hi",
                                   help="أعلى مستوى صف يُسمح بنقله.")
        with g3:
            match_mode = st.radio(
                "طريقة المقارنة:",
                ['range', 'exact'],
                format_func=lambda x: ("المدى: أدنى ≥ المحدد وأعلى ≤ المحدد"
                                       if x == 'range'
                                       else "تطابق تام: أدنى = المحدد وأعلى = المحدد"),
                key="r2o_mode")

        n_lo, n_hi = GRADE_NUM[pick_lo], GRADE_NUM[pick_hi]
        if n_lo > n_hi:
            st.error("«أدنى صف» أكبر من «أعلى صف» — صحّح الاختيار.")
            grade_lo = grade_hi = None
        else:
            grade_lo, grade_hi = n_lo, n_hi
            if match_mode == 'range':
                st.caption("المدى المقبول: من **%s (%d)** إلى **%s (%d)**"
                           % (pick_lo, n_lo, pick_hi, n_hi))
            else:
                st.caption("المطابقة التامة: أدنى صف = **%s** وأعلى صف = **%s**"
                           % (pick_lo, pick_hi))

        # أزرار سريعة تضبط القائمتين معاً
        qq = st.columns(4)
        presets = [("الكل", GRADE_SORT[0], GRADE_SORT[-1]),
                   ("الابتدائية", GRADE_SORT[0], "السادس"),
                   ("الثانوية", "السابع", GRADE_SORT[-1]),
                   ("الروضة", "روضة ثانية", "الثاني")]
        for _i, (_lbl, _a, _b) in enumerate(presets):
            with qq[_i]:
                def _set_preset(_a=_a, _b=_b):
                    st.session_state['r2o_lo'] = _a
                    st.session_state['r2o_hi'] = _b
                st.button(_lbl, key="r2o_p%d" % _i, use_container_width=True,
                          on_click=_set_preset)

    with st.spinner("جارٍ البحث عن المدارس القابلة للنقل..."):
        pairs, miss = rented_to_owned_pairs(_rdf, radius_km,
                                             grade_lo=grade_lo, grade_hi=grade_hi,
                                             match_mode=match_mode,
                                             require_cover=require_cover,
                                             require_grades=require_grades,
                                             allowed_genders=allowed_genders,
                                             exclude_second_period=exclude_second_period,
                                             require_capacity=require_capacity,
                                             king_exclude_codes=king_exclude_codes)

    if miss:
        st.warning("أعمدة ناقصة: " + "، ".join(miss))
        return
    if pairs is None:
        st.info("لا توجد مدارس مستأجرة أو مدارس ملك في البيانات المفلترة.")
        return
    if len(pairs) == 0:
        _why = []
        if king_exclude:
            _why.append("استبعاد ملك بأدنى صف -1 أو -2")
        if require_capacity:
            _why.append("سعة الملك ≥ طلاب المستأجرة")
        if exclude_second_period:
            _why.append("استبعاد مدارس «فترة ثانية»")
        if require_gender and allowed_genders:
            _why.append("الجنس: " + " / ".join(sorted(allowed_genders)))
        if require_cover:
            _why.append("الملك يجب أن يغطي صفوف المستأجرة")
        if require_grades and grade_lo is not None and grade_hi is not None:
            _why.append("صفوف المستأجرة من %s إلى %s"
                        % (grade_label(grade_lo), grade_label(grade_hi)))
        st.info(f"لا توجد مدارس يمكن نقلها إلى مدارس ملك قريبة ضمن {radius_km:g} كم.")
        if _why:
            st.caption("الشروط المفعّلة: " + " · ".join(_why))
        return

    ren_n = pairs['المدرسة المستأجرة'].nunique()
    own_n = pairs['المدرسة الملك'].nunique()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("عدد المدارس المستأجرة القابلة للنقل", f"{ren_n:,}")
    k2.metric("عدد مدارس الملك المستقبِلة", f"{own_n:,}")
    k3.metric("متوسط المسافة", f"{pairs['المسافة (كم)'].mean():.2f} كم")
    k4.metric("إجمالي الطلاب القابلون للنقل",
              f"{int(pairs.groupby('المدرسة المستأجرة')['عدد طلاب المستأجرة'].first().sum()):,}")

    # ---- أثر شرط السعة ----
    if 'الطاقة الاستيعابية للملك' in pairs.columns:
        _tight = int((pairs['الطاقة الاستيعابية للملك'] <=
                      pairs['عدد طلاب المستأجرة']).sum())
        _tight10 = int((pairs['الطاقة الاستيعابية للملك'] <
                        pairs['عدد طلاب المستأجرة'] + 10).sum())
        st.caption("🔎 أزواج ضيّقة السعة: **%d** زوجاً الطاقة فيها = عدد الطلاب بالضبط، "
                   "و **%d** زوجاً الفرق أقل من 10 طلاب."
                   % (_tight, _tight10))
    if king_exclude:
        _kg = int((pairs['أدنى صف الملك'].isin([-1, -2])).sum()) \
            if 'أدنى صف الملك' in pairs.columns else 0
        st.caption("🎓 مدارس الملك بأدنى صف ‎-1/‎-2 ضمن النتائج: **%d** "
                   "(المفترض صفر بعد الاستثناء)." % _kg)
    if not require_capacity:
        st.warning("شرط السعة مُعطَّل — يظهر اقتراحات لمدرسة ملك طاقتها "
                   "أقل من عدد طلاب المستأجرة.")

    st.markdown("###### كل الأزواج (الأقرب والأكبر)")
    best = (pairs.sort_values(['المسافة (كم)', 'عدد طلاب المستأجرة'],
                              ascending=[True, False])
            .drop_duplicates(subset=['المدرسة المستأجرة']))
    # ---- التوزيع حسب الجنس ----
    st.markdown("###### التوزيع حسب جنس المدرسة")
    _gd = (pairs.groupby(['جنس المستأجرة', 'جنس الملك'])
           .size().reset_index(name='عدد الأزواج'))
    if len(_gd):
        _fg = px.bar(_gd, x='جنس المستأجرة', y='عدد الأزواج', color='جنس الملك',
                     barmode='stack', title='أزواج النقل حسب جنس المستأجرة والملك',
                     labels={'عدد الأزواج': 'عدد الأزواج', 'جنس الملك': 'جنس المدرسة الملك'},
                     color_discrete_map={'ذكور': '#2166ac', 'إناث': '#e75480',
                                         'مختلطة': '#1a9850'})
        _fg.update_layout(font_family='Cairo', legend_title_text='')
        st.plotly_chart(_fg, width='stretch', use_container_width=True)
        show_df(_gd, 'توزيع الأزواج حسب الجنس', 'توزيع_الأزواج_الجنس')
    else:
        st.caption("لا توجد أزواج لعرض التوزيع حسب الجنس.")

    _show = best.copy()
    for _c in ('أدنى صف المستأجرة', 'أعلى صف المستأجرة',
               'أدنى صف الملك', 'أعلى صف الملك'):
        if _c in _show.columns:
            _show[_c + ' (الاسم)'] = _show[_c].map(grade_label)
    show_df(_show, 'أفضل أزواج النقل (مستأجر → ملك)', 'أفضل_الأزواج')

    with st.expander("عرض كل الأزواج"):
        show_df(pairs.sort_values('المسافة (كم)'), 'كل أزواج النقل', 'كل_الأزواج')

    # ===== خريطة نتائج النقل =====
    try:
        _need_xy = [c for c in ('اسم المدرسة', 'x', 'y') if c in _rdf.columns]
        if len(_need_xy) < 3:
            st.caption("لا توجد أعمدة الإحداثيات (x، y) لرسم الخريطة.")
        else:
            # أفضل اقتراح واحد لكل مدرسة مستأجرة (أقرب + أوسع فائض)
            _best = (pairs.sort_values(['المسافة (كم)', 'فائض الطاقة (Places)'],
                                       ascending=[True, False])
                     .drop_duplicates(subset=['المدرسة المستأجرة'])
                     .copy())
            if not {'_rx', '_ry', '_ox', '_oy'}.issubset(_best.columns):
                _X = _rdf.copy()
                _X['_x'] = pd.to_numeric(_X['x'], errors='coerce')
                _X['_y'] = pd.to_numeric(_X['y'], errors='coerce')
                _best = _best.merge(
                    _X[['اسم المدرسة', '_x', '_y']].rename(
                        columns={'_x': '_rx', '_y': '_ry'}),
                    left_on='المدرسة المستأجرة', right_on='اسم المدرسة', how='left')
                _best = _best.merge(
                    _X[['اسم المدرسة', '_x', '_y']].rename(
                        columns={'_x': '_ox', '_y': '_oy'}),
                    left_on='المدرسة الملك', right_on='اسم المدرسة',
                    how='left', suffixes=('', '_o'))
            _best = _best.dropna(subset=['_rx', '_ry', '_ox', '_oy'])

            _mc1, _mc2, _mc3 = st.columns(3)
            with _mc1:
                _max_lines = st.slider("أقصى عدد خطوط على الخريطة",
                                       50, 1200, min(400, max(50, len(_best))),
                                       step=50, key="r2o_map_n")
            with _mc2:
                _as_road = st.checkbox("خطوط الشارع (OSRM) للاقتراحات",
                                       value=True, key="r2o_map_road",
                                       help="أبطأ — يجلب مسار الشارع لكل زوج.")
            with _mc3:
                st.metric("عدد الاقتراحات المرسومة", f"{min(len(_best), _max_lines):,}")

            _show = _best.head(_max_lines)

            _mm = folium.Map(location=[31.5, 36.0], zoom_start=8, prefer_canvas=True)
            folium.TileLayer(
                tiles='https://server.arcgisonline.com/ArcGIS/rest/services/'
                      'World_Imagery/MapServer/tile/{z}/{y}/{x}',
                attr='Esri', name='ستلايت', overlay=True, control=True, show=False
            ).add_to(_mm)

            # ---- خطوط ----
            if _as_road:
                for _, _r in _show.iterrows():
                    _res = road_route(float(_r['_ry']), float(_r['_rx']),
                                      float(_r['_oy']), float(_r['_ox']))
                    _pts = (list(_res[2])
                            if _res and isinstance(_res[2], list)
                            and _res[2] and isinstance(_res[2][0], list)
                            else [[_r['_ry'], _r['_rx']], [_r['_oy'], _r['_ox']]])
                    folium.PolyLine(
                        _pts,
                        color='#c0392b', weight=3, opacity=0.75,
                        tooltip=(f"<b>من:</b> {_r['المدرسة المستأجرة']}<br>"
                                 f"<b>إلى:</b> {_r['المدرسة الملك']}<br>"
                                 f"<b>المسافة:</b> {_r['المسافة (كم)']} كم")
                    ).add_to(_mm)
            else:
                for _, _r in _show.iterrows():
                    folium.PolyLine(
                        [[_r['_ry'], _r['_rx']], [_r['_oy'], _r['_ox']]],
                        color='#c0392b', weight=3, opacity=0.75,
                        tooltip=(f"<b>من:</b> {_r['المدرسة المستأجرة']}<br>"
                                 f"<b>إلى:</b> {_r['المدرسة الملك']}<br>"
                                 f"<b>المسافة:</b> {_r['المسافة (كم)']} كم<br>"
                                 f"<b>الطلاب:</b> {_r['عدد طلاب المستأجرة']}<br>"
                                 f"<b>الطاقة:</b> {_r['الطاقة الاستيعابية للملك']}")
                    ).add_to(_mm)

            # ---- نقاط النهاية ----
            for _, _r in _show.iterrows():
                folium.CircleMarker(
                    [_r['_ry'], _r['_rx']], radius=7, color='#c0392b', weight=2,
                    fill=True, fill_color='#ffffff', fill_opacity=1.0,
                    popup=(f"<b>مستأجر:</b> {_r['المدرسة المستأجرة']}<br>"
                           f"طلاب: {_r['عدد طلاب المستأجرة']}<br>"
                           f"جنس: {_r['جنس المستأجرة']}")).add_to(_mm)
                folium.CircleMarker(
                    [_r['_oy'], _r['_ox']], radius=7, color='#1a9850', weight=2,
                    fill=True, fill_color='#1a9850', fill_opacity=1.0,
                    popup=(f"<b>ملك:</b> {_r['المدرسة الملك']}<br>"
                           f"الطاقة: {_r['الطاقة الاستيعابية للملك']}<br>"
                           f"جنس: {_r['جنس الملك']}<br>"
                           f"فائض: {_r['فائض الطاقة (Places)']}")).add_to(_mm)

            try:
                _la = list(_show['_ry']) + list(_show['_oy'])
                _lo = list(_show['_rx']) + list(_show['_ox'])
                if _la:
                    _mm.fit_bounds([[min(_la), min(_lo)], [max(_la), max(_lo)]],
                                   padding=(40, 40))
            except Exception:
                pass

            folium.LayerControl().add_to(_mm)
            components.html(_mm.get_root().render(), height=620, scrolling=False)
            st.caption("🔴 نقطة المستأجرة (دائرة بيضاء) — 🟢 نقطة الملك — "
                       "الخط الأحمر يربط كل مستأجرة بمدرسة الملك المقترحة. "
                       "متعدد الاقتراحات لنفس المدرسة تُرسم جميعها.")
            _b3 = io.BytesIO()
            with pd.ExcelWriter(_b3, engine='xlsxwriter') as _w:
                _show.to_excel(_w, index=False, sheet_name='اقتراحات')
            _fmt_p = st.radio('صيغة تصدير الاقتراحات:', ['Excel', 'PDF', 'CSV'],
                              horizontal=True, key='prop_fmt')
            if _fmt_p == 'CSV':
                st.download_button(
                    "⬇️ تنزيل الاقتراحات (CSV)",
                    data=_show.to_csv(index=False).encode('utf-8-sig'),
                    file_name='اقتراحات_نقل_المدارس.csv',
                    mime='text/csv')
            elif _fmt_p == 'PDF':
                st.download_button(
                    "⬇️ تنزيل الاقتراحات (PDF)",
                    data=to_pdf_bytes({'الاقتراحات': _show},
                                      title='اقتراحات نقل المدارس'),
                    file_name='اقتراحات_نقل_المدارس.pdf',
                    mime='application/pdf')
            else:
                st.download_button(
                    "⬇️ تنزيل الاقتراحات (Excel)",
                    data=_b3.getvalue(),
                    file_name='اقتراحات_نقل_المدارس.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as exc:
        st.caption("تعذّر رسم الخريطة: %s" % exc)



# ============================================================
#  ترقيم الصفوف: روضة ثانية = -2 … الثاني عشر = 12
# ============================================================
GRADE_TABLE = [
    ('روضة ثانية', -2),
    ('روضة أولى', -1),
    ('الأول', 1),
    ('الثاني', 2),
    ('الثالث', 3),
    ('الرابع', 4),
    ('الخامس', 5),
    ('السادس', 6),
    ('السابع', 7),
    ('الثامن', 8),
    ('التاسع', 9),
    ('العاشر', 10),
    ('الحادي عشر', 11),
    ('الثاني عشر', 12),
]

GRADE_NUM2NAME = {n: nm for nm, n in GRADE_TABLE}
GRADE_SORT = [nm for nm, _ in GRADE_TABLE]
GRADE_NUMS = {n for _, n in GRADE_TABLE}


def _norm_grade(v):
    """توحيد الكتابة: مسافات، ألف،ياء."""
    s = str(v).strip()
    if s.lower() == 'nan' or s == '':
        return None
    s = re.sub(r'\s+', ' ', s)
    s = s.replace('\u0623', '\u0627').replace('\u0625', '\u0627').replace('\u0622', '\u0627')
    s = s.replace('\u0649', '\u064a')
    s = s.replace('الاخ ', 'الحادي ')
    return s.strip()


GRADE_NUM = {}
for _nm, _n in GRADE_TABLE:
    GRADE_NUM[_nm] = _n
    _alt = _norm_grade(_nm)
    if _alt:
        GRADE_NUM[_alt] = _n
# متغيّرات شائعة
GRADE_NUM.update({
    'روضة ثانيه': -2, 'روضة 1': -2, 'تمهيدي': -2,
    'روضة اولي': -1, 'تمهيدي 1': -1,
    'الحادي عشر': 11, 'حادي عشر': 11,
    'ثاني عشر': 12, ' الثاني عشر': 12,
})


def grade_number(v):
    """رقم الصف أو None إن كان غير معروف."""
    s = _norm_grade(v)
    if s is None:
        return None
    if s in GRADE_NUM:
        return GRADE_NUM[s]
    for k, n in GRADE_NUM.items():
        if k and (k in s or s in k):
            return n
    return None


def grade_label(n):
    return GRADE_NUM2NAME.get(n, "غير محدد")


with col2:
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 نظرة عامة",
        "🏫 تحليل المدارس",
        "👥 تحليل الطلاب",
        "📋 البيانات التفصيلية",
        "🗺️ خريطة الواقع التربوي للمدارس",
        "📚 دراسات متنوعة"
    ])

    # ====== تبويب 1: نظرة عامة ======
    with tab1:
        st.subheader("📊 نظرة عامة على البيانات")

        col_chart1, col_chart2 = st.columns(2)

        with col_chart1:
            st.markdown("#### توزيع المدارس حسب المديرية")
            dir_counts = filtered_df['المديرية'].value_counts()
            fig1 = px.bar(
                x=dir_counts.values,
                y=dir_counts.index,
                orientation='h',
                color=dir_counts.values,
                color_continuous_scale='Blues',
                text=dir_counts.values
            )
            fig1.update_layout(
                height=400,
                showlegend=False,
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family='Cairo', size=12),
                xaxis_title="عدد المدارس",
                yaxis_title=""
            )
            st.plotly_chart(fig1, width='stretch')

        with col_chart2:
            st.markdown("#### توزيع المدارس حسب الجنس")
            gender_counts = filtered_df['الجنس'].value_counts()
            fig2 = px.pie(
                values=gender_counts.values,
                names=gender_counts.index,
                color=gender_counts.index,
                color_discrete_map={'ذكور': '#3b82f6', 'إناث': '#ec4899', 'مختلطة': '#10b981'}
            )
            fig2.update_layout(
                height=400,
                showlegend=True,
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family='Cairo', size=12)
            )
            st.plotly_chart(fig2, width='stretch')

        col_chart3, col_chart4 = st.columns(2)

        with col_chart3:
            st.markdown("#### توزيع المدارس حسب نوع التعليم")
            edu_counts = filtered_df['نوع التعليم'].value_counts()
            fig3 = px.bar(
                x=edu_counts.index,
                y=edu_counts.values,
                color=edu_counts.values,
                color_continuous_scale='Viridis',
                text=edu_counts.values
            )
            fig3.update_layout(
                height=350,
                showlegend=False,
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family='Cairo', size=12),
                xaxis_title="",
                yaxis_title="عدد المدارس"
            )
            st.plotly_chart(fig3, width='stretch')

        with col_chart4:
            st.markdown("#### توزيع المدارس حسب الملكية")
            own_counts = filtered_df['الملكية'].value_counts()
            fig4 = px.pie(
                values=own_counts.values,
                names=own_counts.index,
                color=own_counts.index,
                color_discrete_map={'ملك': '#10b981', 'مستأجر': '#f59e0b'}
            )
            fig4.update_layout(
                height=350,
                showlegend=True,
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family='Cairo', size=12)
            )
            st.plotly_chart(fig4, width='stretch')

    # ====== تبويب 2: تحليل المدارس ======
    with tab2:
        st.subheader("🏫 تحليل المدارس")

        col_s1, col_s2 = st.columns(2)

        with col_s1:
            st.markdown("#### أعلى 10 مدارس من حيث عدد الطلاب")
            top_schools = filtered_df.nlargest(10, 'مجموع الطلاب')[['اسم المدرسة', 'مجموع الطلاب', 'المديرية']]
            fig5 = px.bar(
                x=top_schools['اسم المدرسة'],
                y=top_schools['مجموع الطلاب'],
                color=top_schools['مجموع الطلاب'],
                color_continuous_scale='Blues',
                text=top_schools['مجموع الطلاب']
            )
            fig5.update_layout(
                height=400,
                showlegend=False,
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family='Cairo', size=11),
                xaxis_title="",
                yaxis_title="عدد الطلاب"
            )
            st.plotly_chart(fig5, width='stretch')

        with col_s2:
            st.markdown("#### توزيع المدارس حسب الإقليم")
            region_counts = filtered_df['الإقليم'].value_counts()
            fig6 = px.pie(
                values=region_counts.values,
                names=region_counts.index,
                color=region_counts.index,
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig6.update_layout(
                height=400,
                showlegend=True,
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family='Cairo', size=12)
            )
            st.plotly_chart(fig6, width='stretch')

        st.markdown("---")
        st.markdown("#### توزيع المدارس حسب المحافظة")
        gov_counts = filtered_df['المحافظة'].value_counts()
        fig7 = px.bar(
            x=gov_counts.index,
            y=gov_counts.values,
            color=gov_counts.values,
            color_continuous_scale='Teal',
            text=gov_counts.values
        )
        fig7.update_layout(
            height=400,
            showlegend=False,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family='Cairo', size=12),
            xaxis_title="",
            yaxis_title="عدد المدارس"
        )
        st.plotly_chart(fig7, width='stretch')

    # ====== تبويب 3: تحليل الطلاب ======
    with tab3:
        st.subheader("👥 تحليل الطلاب")

        col_st1, col_st2 = st.columns(2)

        with col_st1:
            st.markdown("#### توزيع الطلاب حسب المديرية")
            dir_students = filtered_df.groupby('المديرية')['مجموع الطلاب'].sum().sort_values(ascending=False)
            fig8 = px.bar(
                x=dir_students.values,
                y=dir_students.index,
                orientation='h',
                color=dir_students.values,
                color_continuous_scale='Greens',
                text=dir_students.values
            )
            fig8.update_layout(
                height=400,
                showlegend=False,
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family='Cairo', size=12),
                xaxis_title="عدد الطلاب",
                yaxis_title=""
            )
            st.plotly_chart(fig8, width='stretch')

        with col_st2:
            st.markdown("#### نسبة الذكور إلى الإناث")
            gender_data = filtered_df.groupby('الجنس')[['مجموع الطلاب الذكور', 'مجموع الطلاب الاناث']].sum()
            fig9 = go.Figure()
            fig9.add_trace(go.Bar(
                name='الذكور',
                x=gender_data.index,
                y=gender_data['مجموع الطلاب الذكور'],
                marker_color='#3b82f6'
            ))
            fig9.add_trace(go.Bar(
                name='الإناث',
                x=gender_data.index,
                y=gender_data['مجموع الطلاب الاناث'],
                marker_color='#ec4899'
            ))
            fig9.update_layout(
                barmode='group',
                height=400,
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family='Cairo', size=12),
                xaxis_title="",
                yaxis_title="عدد الطلاب"
            )
            st.plotly_chart(fig9, width='stretch')

        st.markdown("---")
        st.markdown("#### توزيع الطلاب حسب المحافظة")
        gov_students = filtered_df.groupby('المحافظة')['مجموع الطلاب'].sum().sort_values(ascending=False)
        fig10 = px.bar(
            x=gov_students.index,
            y=gov_students.values,
            color=gov_students.values,
            color_continuous_scale='Oranges',
            text=gov_students.values
        )
        fig10.update_layout(
            height=400,
            showlegend=False,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family='Cairo', size=12),
            xaxis_title="",
            yaxis_title="عدد الطلاب"
        )
        st.plotly_chart(fig10, width='stretch')

    # ====== تبويب 4: البيانات التفصيلية ======
    with tab4:
        st.subheader("📋 البيانات التفصيلية")

        col_stat1, col_stat2, col_stat3, col_stat4 = st.columns(4)
        with col_stat1:
            st.metric("عدد المدارس", f"{len(filtered_df):,}")
        with col_stat2:
            st.metric("إجمالي الطلاب", f"{int(filtered_df['مجموع الطلاب'].sum()):,}")
        with col_stat3:
            st.metric("متوسط الطلاب", f"{int(filtered_df['مجموع الطلاب'].mean())}")
        with col_stat4:
            st.metric("أكبر مدرسة", f"{int(filtered_df['مجموع الطلاب'].max())}")

        st.markdown("---")

        display_cols = ['الرقم الوطني', 'المديرية', 'اسم المدرسة', 'الجنس', 'أدنى صف', 'أعلى صف',
                       'المحافظة', 'اللواء', 'نوع التعليم', 'الملكية', 'مجموع الطلاب الذكور',
                       'مجموع الطلاب الاناث', 'مجموع الطلاب']
        display_df = filtered_df[display_cols].copy()
        display_df.columns = ['الرقم الوطني', 'المديرية', 'اسم المدرسة', 'الجنس', 'أدنى صف', 'أعلى صف',
                             'المحافظة', 'اللواء', 'نوع التعليم', 'الملكية', 'الذكور', 'الإناث', 'المجموع']

        show_df(display_df, 'البيانات التفصيلية للمدارس', 'المدارس_الحكومية', height=500)

    # ====== تبويب 5: خريطة الواقع التربوي للمدارس ======
    with tab5:
        st.subheader("🗺️ خريطة الواقع التربوي للمدارس")

        # التحقق من وجود أعمدة الإحداثيات
        x_col = None
        y_col = None
        for c in filtered_df.columns:
            c_str = str(c).strip().lower()
            if c_str in ['x', 'خط الطول', 'longitude', 'lng']:
                x_col = c
            if c_str in ['y', 'خط العرض', 'latitude', 'lat']:
                y_col = c

        if x_col and y_col:
            # تحويل الإحداثيات إلى أرقام
            filtered_df[x_col] = pd.to_numeric(filtered_df[x_col], errors='coerce')
            filtered_df[y_col] = pd.to_numeric(filtered_df[y_col], errors='coerce')

            # تصفية الصفوف ذات الإحداثيات الصحيحة
            map_df = filtered_df[(filtered_df[x_col].notna()) & (filtered_df[y_col].notna())].copy()

            # تصفية الإحداثيات داخل حدود الأردن فقط
            # حدود الأردن التقريبية: خط العرض 29.0-33.5، خط الطول 34.5-39.5
            map_df = map_df[
                (map_df[y_col] >= 29.0) & (map_df[y_col] <= 33.5) &
                (map_df[x_col] >= 34.5) & (map_df[x_col] <= 39.5)
            ].copy()

            if len(map_df) > 0:
                # ====== خيارات العرض والأداء ======
                c1, c2, c3 = st.columns(3)
                with c1:
                    show_names = st.checkbox(
                        "عرض أسماء المدارس", value=False,
                        help="تُرسم أسماء كل المدارس داخل الشاشة فقط، "
                             "وتتحدّث تلقائياً عند التكبير — "
                             " zooming in يعرض أسماء أكثر.")
                with c2:
                    show_heat = st.checkbox("طبقة الكثافة الحرارية", value=True)
                with c3:
                    st.metric("عدد المدارس على الخريطة", f"{len(map_df):,}")

                _gcol = 'الجنس' if 'الجنس' in map_df.columns else None
                _scol = 'مجموع الطلاب' if 'مجموع الطلاب' in map_df.columns else None

                _pts = [[float(a), float(b)] for a, b in
                        zip(map_df[y_col].tolist(), map_df[x_col].tolist())]
                _gen = (map_df[_gcol].astype(str).tolist() if _gcol
                        else [''] * len(_pts))
                _stu = (pd.to_numeric(map_df[_scol], errors='coerce').fillna(0).tolist()
                        if _scol else [0] * len(_pts))
                # الأسماء للأكبر حجماً فقط (حتى لا تُرسم أسماء فارغة)
                _nam = [str(x) for x in map_df['اسم المدرسة'].tolist()] \
                    if 'اسم المدرسة' in map_df.columns else [''] * len(_pts)

                with st.spinner("جارٍ بناء الخريطة..."):
                    _html = build_school_map(
                        tuple(_pts), tuple(_gen), tuple(_stu), tuple(_nam),
                        bool(show_names), bool(show_heat),
                        MAP_NAME_FONT_PX, MAP_NAME_WEIGHT)
                components.html(_html, height=600, scrolling=False)

                if show_names:
                    st.caption("الأسماء الظاهرة تتبع التكبير: كل ما داخل الشاشة يُعرض، "
                               "ويبدأ من الأكبر حجماً عند مستويات التكبير المنخفضة.")

                # إضافة دليل الألوان
                st.markdown("""
                <div style="display: flex; gap: 20px; margin-top: 10px;">
                    <div><span style="color: blue; font-size: 20px;">●</span> مدارس ذكور</div>
                    <div><span style="color: pink; font-size: 20px;">●</span> مدارس إناث</div>
                    <div><span style="color: green; font-size: 20px;">●</span> مدارس مختلطة</div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.warning("لا توجد إحداثيات صالحة في البيانات المفلترة")
        else:
            st.warning("لم يتم العثور على أعمدة الإحداثيات (x, y) في البيانات")

    # ====== تبويب 6: دراسات متنوعة ======
    with tab6:
        st.subheader("📚 دراسات متنوعة")
        st.caption("اختر دراسة لعرض نتائجها، ثم يمكنك إلغائها أو مسح كل الدراسات.")


        NO_STUDY = "— بدون دراسة —"
        study_options = [
            NO_STUDY,
            "توزيع المدارس حسب المديرية",
            "توزيع المدارس حسب المحافظة",
            "توزيع المدارس حسب الجنس",
            "توزيع المدارس حسب نوع التعليم",
            "توزيع المدارس حسب الملكية",
            "تحليل الطاقة الاستيعابية",
            "تحليل أعداد الطلاب",
            "العلاقة بين المديرية وعدد الطلاب",
            "المسافة بين مدارس (ب الشارع)",
            "امكانية نقل المدارس المستأجرة الى مدارس الملك",
            "توزيع المدارس حسب ادنى صف و اعلى صف"
        ]
        # إلغاء مُعلَّم من تشغيل سابق: يُطبَّق قبل إنشاء القائمة
        # (لو ضبطنا study_select بعد إنشائه، تُعيد الواجهة قيمته القديمة)
        if st.session_state.pop('reset_study_pending', False):
            st.session_state['study_select'] = NO_STUDY

        selected_study = st.selectbox("اختر الدراسة:", study_options, key="study_select")

        # سجلّ الدراسات التي عُرضت سابقاً (للإلغاء والمسح)
        if 'studies_shown' not in st.session_state:
            st.session_state.studies_shown = []
        if selected_study != NO_STUDY and selected_study not in st.session_state.studies_shown:
            st.session_state.studies_shown.append(selected_study)

        # ---- خياران يظهران عند اختيار دراسة ----
        if selected_study != NO_STUDY:
            n_prev = len(st.session_state.studies_shown)
            st.caption(f"الدراسة المختارة: **{selected_study}**  ·  عدد الدراسات المستخدمة: **{n_prev}**")
            cc1, cc2 = st.columns(2)
            with cc1:
                if st.button("✖ إلغاء الدراسة", key="cancel_study_btn", use_container_width=True,
                             type="secondary"):
                    st.session_state.studies_shown = [s for s in st.session_state.studies_shown
                                                      if s != selected_study]
                    st.session_state.reset_study_pending = True
                    st.rerun()
            with cc2:
                if st.button("🗑 مسح جميع الدراسات", key="clear_studies_btn", use_container_width=True,
                             type="secondary"):
                    st.session_state.studies_shown = []
                    st.session_state.reset_study_pending = True
                    st.rerun()
            st.markdown("---")

        if selected_study == "توزيع المدارس حسب المديرية":
            study_df = filtered_df['المديرية'].value_counts()
            st.markdown(f"#### توزيع المدارس حسب المديرية ({len(study_df):,} مديرية)")
            _fig_dir = px.bar(
                pd.DataFrame({'المديرية': study_df.index,
                              'عدد المدارس': study_df.values}),
                x='المديرية', y='عدد المدارس', color='عدد المدارس',
                title='عدد المدارس في كل مديرية', text='عدد المدارس')
            _fig_dir.update_layout(xaxis={'tickangle': -60},
                                   font_family='Cairo', showlegend=False)
            st.plotly_chart(_fig_dir, width='stretch', use_container_width=True)
            with st.expander("عرض كقائمة نصية", expanded=False):
                for name, count in study_df.items():
                    st.write(f"**{name}:** {int(count):,} مدرسة")
            directorate_table(filtered_df, 'جدول كل المديريات — التوزيع العام')

        elif selected_study == "توزيع المدارس حسب المحافظة":
            study_df = filtered_df['المحافظة'].value_counts()
            st.markdown("#### توزيع المدارس حسب المحافظة")
            for name, count in study_df.items():
                st.write(f"**{name}:** {count} مدرسة")

        elif selected_study == "توزيع المدارس حسب الجنس":
            study_df = filtered_df['الجنس'].value_counts()
            st.markdown("#### توزيع المدارس حسب الجنس")
            for name, count in study_df.items():
                st.write(f"**{name}:** {count} مدرسة")

        elif selected_study == "توزيع المدارس حسب نوع التعليم":
            study_df = filtered_df['نوع التعليم'].value_counts()
            st.markdown("#### توزيع المدارس حسب نوع التعليم")
            for name, count in study_df.items():
                st.write(f"**{name}:** {count} مدرسة")

        elif selected_study == "توزيع المدارس حسب الملكية":
            study_df = filtered_df['الملكية'].value_counts()
            st.markdown("#### توزيع المدارس حسب الملكية")
            for name, count in study_df.items():
                st.write(f"**{name}:** {count} مدرسة")

        elif selected_study == "تحليل الطاقة الاستيعابية":
            if 'الطاقة الاستيعابية' in filtered_df.columns:
                capacity = pd.to_numeric(filtered_df['الطاقة الاستيعابية'], errors='coerce')
                st.markdown("#### تحليل الطاقة الاستيعابية")
                st.write(f"**متوسط الطاقة الاستيعابية:** {capacity.mean():.0f}")
                st.write(f"**أعلى طاقة استيعابية:** {capacity.max():.0f}")
                st.write(f"**أدنى طاقة استيعابية:** {capacity.min():.0f}")
                st.write(f"**إجمالي الطاقة الاستيعابية:** {capacity.sum():.0f}")
            else:
                st.warning("لا توجد بيانات عن الطاقة الاستيعابية")

            # دراسة انتقال المدارس المستأجرة إلى المدارس الملك القريبة
            st.markdown("---")
            st.markdown("#### دراسة انتقال المدارس المستأجرة إلى المدارس الملك القريبة")

            # التحقق من وجود الأعمدة المطلوبة
            required_cols = ['الملكية', 'الطاقة الاستيعابية', 'مجموع الطلاب', 'x', 'y']
            missing_cols = [c for c in required_cols if c not in filtered_df.columns]
            if missing_cols:
                st.warning(f"الأعمدة التالية غير موجودة: {', '.join(missing_cols)}")
            else:
                # تحديد المدارس المستأجرة والملك
                rented_schools = filtered_df[filtered_df['الملكية'].astype(str).str.contains('مستأجر', na=False)].copy()
                owned_schools = filtered_df[filtered_df['الملكية'].astype(str).str.contains('ملك', na=False)].copy()

                if len(rented_schools) == 0 or len(owned_schools) == 0:
                    st.warning("لا توجد مدارس مستأجرة أو ملك في البيانات المفلترة")
                else:
                    # تحويل الأعمدة إلى أرقام
                    rented_schools['مجموع الطلاب'] = pd.to_numeric(rented_schools['مجموع الطلاب'], errors='coerce')
                    owned_schools['الطاقة الاستيعابية'] = pd.to_numeric(owned_schools['الطاقة الاستيعابية'], errors='coerce')
                    rented_schools['x'] = pd.to_numeric(rented_schools['x'], errors='coerce')
                    rented_schools['y'] = pd.to_numeric(rented_schools['y'], errors='coerce')
                    owned_schools['x'] = pd.to_numeric(owned_schools['x'], errors='coerce')
                    owned_schools['y'] = pd.to_numeric(owned_schools['y'], errors='coerce')

                    # حساب المسافة بين المدارس المستأجرة والملك
                    def haversine_distance(lat1, lon1, lat2, lon2):
                        import math
                        R = 6371
                        lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
                        dlat = lat2 - lat1
                        dlon = lon2 - lon1
                        a = math.sin(dlat/2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2)**2
                        c = 2 * math.asin(math.sqrt(a))
                        return R * c

                    # البحث عن المدارس الملك القريبة التي تستوعب طلاب المدارس المستأجرة
                    # حساب متجهي (أسرع بعOrders_OF_MAGNITUDE من الحلقات المتداخلة)
                    _R = 6371.0
                    def _dist_matrix(lat1, lon1, lat2, lon2):
                        p1, p2 = _math.radians(lat1), _math.radians(lat2)
                        dp = p2 - p1
                        dl = _math.radians(lon2) - _math.radians(lon1)
                        a = _math.sin(dp / 2) ** 2 + \
                            _math.cos(p1) * _math.cos(p2) * _math.sin(dl / 2) ** 2
                        return _R * 2 * _math.asin(_math.sqrt(a))

                    results = []
                    r_lat = rented_schools['y'].tolist()
                    r_lon = rented_schools['x'].tolist()
                    o_lat = owned_schools['y'].tolist()
                    o_lon = owned_schools['x'].tolist()
                    r_name = rented_schools.get('اسم المدرسة', pd.Series(['غير معروف'] * len(rented_schools))).tolist()
                    o_name = owned_schools.get('اسم المدرسة', pd.Series(['غير معروف'] * len(owned_schools))).tolist()
                    r_stu = rented_schools['مجموع الطلاب'].tolist()
                    o_cap = owned_schools['الطاقة الاستيعابية'].tolist()

                    for i in range(len(r_lat)):
                        dists = _dist_matrix(r_lat[i], r_lon[i], o_lat, o_lon)
                        for j in range(len(o_lat)):
                            d = dists[j]
                            if d <= 10 and o_cap[j] >= r_stu[i]:
                                results.append({
                                    'المدرسة المستأجرة': r_name[i],
                                    'المدرسة الملك': o_name[j],
                                    'عدد طلاب المستأجرة': r_stu[i],
                                    'الطاقة الاستيعابية للملك': o_cap[j],
                                    'المسافة (كم)': round(d, 2)
                                })

                    if results:
                        result_df = pd.DataFrame(results)
                        st.write(f"**عدد المدارس المستأجرة التي يمكن نقلها:** {len(result_df)}")
                        show_df(result_df, 'مدارس يمكن نقلها', 'نقل_المدارس')
                    else:
                        st.info("لا توجد مدارس مستأجرة يمكن نقلها إلى مدارس ملك قريبة ضمن 10 كم")

        elif selected_study == "تحليل أعداد الطلاب":
            students = pd.to_numeric(filtered_df['مجموع الطلاب'], errors='coerce')
            st.markdown("#### تحليل أعداد الطلاب")
            st.write(f"**إجمالي الطلاب:** {students.sum():.0f}")
            st.write(f"**متوسط الطلاب لكل مدرسة:** {students.mean():.0f}")
            st.write(f"**أعلى مدرسة في عدد الطلاب:** {students.max():.0f}")
            st.write(f"**أدنى مدرسة في عدد الطلاب:** {students.min():.0f}")

        elif selected_study == "العلاقة بين المديرية وعدد الطلاب":
            dir_students = (filtered_df.groupby('المديرية')['مجموع الطلاب']
                            .sum().sort_values(ascending=False))
            st.markdown(f"#### إجمالي الطلاب في كل مديرية ({len(dir_students):,})")
            _fig_ds = px.bar(
                pd.DataFrame({'المديرية': dir_students.index,
                              'إجمالي الطلاب': dir_students.values}),
                x='المديرية', y='إجمالي الطلاب', color='إجمالي الطلاب',
                text='إجمالي الطلاب', title='إجمالي الطلاب في كل مديرية')
            _fig_ds.update_layout(xaxis={'tickangle': -60},
                                  font_family='Cairo', showlegend=False)
            st.plotly_chart(_fig_ds, width='stretch', use_container_width=True)
            with st.expander("عرض كقائمة نصية", expanded=False):
                for name, count in dir_students.items():
                    st.write(f"**{name}:** {int(count):,} طالب")
            directorate_table(filtered_df,
                               'جدول كل المديريات — مع إجمالي الطلاب',
                               sort_by='إجمالي الطلاب')

        elif selected_study == "توزيع المدارس حسب ادنى صف و اعلى صف":
            st.markdown("#### توزيع المدارس حسب أدنى صف وأعلى صف")
            st.caption("كل صف مرقّم: روضة ثانية = ‎-2، روضة أولى = ‎-1، "
                       "والأول = 1 … الثاني عشر = 12. "
                       "تُعرض كل المدارس التي يقع مداها الصفّي بالكامل بين "
                       "«أدنى صف» و«أعلى صف» المختارَين: أدنى صف في المدرسة "
                       "أكبر من أو يساوي الحد الأدنى، وأعلى صف فيها أصغر من أو "
                       "يساوي الحد الأعلى.")

            _mn, _mx = "أدنى صف", "أعلى صف"
            _miss = [c for c in (_mn, _mx) if c not in filtered_df.columns]
            if _miss:
                st.warning("أعمدة ناقصة: " + "، ".join(_miss))
            else:
                _g = filtered_df.copy()
                _g['_lo'] = pd.to_numeric(_g[_mn].map(grade_number), errors='coerce')
                _g['_hi'] = pd.to_numeric(_g[_mx].map(grade_number), errors='coerce')

                gc1, gc2 = st.columns(2)
                with gc1:
                    sel_lo = st.multiselect(
                        "الصفوف المقبولة لـ «أدنى صف»", GRADE_SORT,
                        default=[], key="dist_lo",
                        help="حد أدنى للمدى: تُعرض المدارس التي يبدأ أدنى صف فيها من هذا الصف أو أعلى.")
                with gc2:
                    sel_hi = st.multiselect(
                        "الصفوف المقبولة لـ «أعلى صف»", GRADE_SORT,
                        default=[], key="dist_hi",
                        help="حد أعلى للمدى: تُعرض المدارس التي ينتهي أعلى صف فيها عند هذا الصف أو أدنى.")

                _lo_set = {GRADE_NUM[g] for g in sel_lo if g in GRADE_NUM}
                _hi_set = {GRADE_NUM[g] for g in sel_hi if g in GRADE_NUM}

                st.markdown("###### الجدول المرجعي للصفوف")
                ref = pd.DataFrame(GRADE_TABLE, columns=['الصف', 'الرقم'])
                ref = ref[ref['الصف'].isin(
                    sorted(set(_lo_set) | set(_hi_set), key=lambda n: n))] \
                    if (_lo_set or _hi_set) else ref
                show_df(ref, 'الجدول المرجعي للصفوف', 'الجدول_المرجعي_للصفوف')

                # ================= التوزيع العام (بدون شرط) =================
                _cnt_lo = pd.Series([n for n in _g['_lo'] if n is not None]) \
                    .value_counts().reindex(sorted(GRADE_NUMS), fill_value=0)
                _cnt_hi = pd.Series([n for n in _g['_hi'] if n is not None]) \
                    .value_counts().reindex(sorted(GRADE_NUMS), fill_value=0)
                _dist = pd.DataFrame({
                    'الصف': [grade_label(n) for n in sorted(GRADE_NUMS)],
                    'عدد المدارس (أدنى صف)': _cnt_lo.tolist(),
                    'عدد المدارس (أعلى صف)': _cnt_hi.tolist(),
                })
                _fig = px.bar(
                    _dist, x='الصف', y=['عدد المدارس (أدنى صف)', 'عدد المدارس (أعلى صف)'],
                    barmode='group', title='توزيع جميع المدارس (بدون شرط)',
                    labels={'value': 'عدد المدارس', 'variable': ''},
                    color_discrete_map={'عدد المدارس (أدنى صف)': '#1a9850',
                                        'عدد المدارس (أعلى صف)': '#2166ac'})
                _fig.update_layout(xaxis={'categoryorder': 'array',
                                          'categoryarray': [grade_label(n)
                                                            for n in sorted(GRADE_NUMS)]},
                                   legend_title_text='', font_family='Cairo')
                if _lo_set and _hi_set:
                    with st.expander("عرض توزيع جميع المدارس (بدون شرط)"):
                        st.plotly_chart(_fig, width='stretch', use_container_width=True)
                        show_df(_dist, 'توزيع جميع المدارس (بدون شرط)', 'توزيع_جميع_المدارس')
                else:
                    st.markdown("###### توزيع جميع المدارس (حسب الأرقام)")
                    st.plotly_chart(_fig, width='stretch', use_container_width=True)
                    show_df(_dist, 'توزيع جميع المدارس (حسب الأرقام)', 'توزيع_جميع_المدارس')

                # ================= الدراسة المطلوبة =================
                if not _lo_set or not _hi_set:
                    st.info("اختر صفاً واحداً على الأقل من كل قائمة أعلاه "
                            "لعرض نتيجة الدراسة.")
                else:
                    # المدارس التي تقع كاملة بين أدنى صف وأعلى صف المختارين:
                    # أدنى صف في المدرسة >= أدنى حد مختار ،
                    # وأعلى صف في المدرسة <= أعلى حد مختار .
                    _lo_bound = min(_lo_set)
                    _hi_bound = max(_hi_set)
                    _mask = (_g['_lo'].notna()) & (_g['_hi'].notna()) & \
                            (_g['_lo'] >= _lo_bound) & (_g['_hi'] <= _hi_bound)
                    _sel = _g[_mask].copy()

                    _total = len(_g)
                    k1, k2, k3, k4 = st.columns(4)
                    k1.metric("عدد المدارس المطابقة", f"{len(_sel):,}")
                    k2.metric("النسبة من الإجمالي",
                              f"{(len(_sel) / _total * 100):.1f}٪" if _total else "—")
                    _stu = pd.to_numeric(_sel.get('مجموع الطلاب', pd.Series(dtype=float)),
                                         errors='coerce').sum()
                    k3.metric("إجمالي طلابها",
                              f"{int(_stu):,}" if _stu == _stu else "—")
                    k4.metric("بلا صف معروف",
                              f"{int(_g['_lo'].isna().sum()):,}")

                    st.markdown("###### توزيع المدارس المطابقة")
                    _sl = _sel['_lo'].value_counts() \
                        .reindex(sorted(GRADE_NUMS), fill_value=0)
                    _sh = _sel['_hi'].value_counts() \
                        .reindex(sorted(GRADE_NUMS), fill_value=0)
                    _d2 = pd.DataFrame({
                        'الصف': [grade_label(n) for n in sorted(GRADE_NUMS)],
                        'عدد المدارس (أدنى صف)': _sl.tolist(),
                        'عدد المدارس (أعلى صف)': _sh.tolist(),
                    })
                    _f2 = px.bar(
                        _d2, x='الصف',
                        y=['عدد المدارس (أدنى صف)', 'عدد المدارس (أعلى صف)'],
                        barmode='group', title='المدارس المطابقة للشرط',
                        labels={'value': 'عدد المدارس', 'variable': ''},
                        color_discrete_map={'عدد المدارس (أدنى صف)': '#1a9850',
                                            'عدد المدارس (أعلى صف)': '#2166ac'})
                    _f2.update_layout(xaxis={'categoryorder': 'array',
                                             'categoryarray': [grade_label(n)
                                                               for n in sorted(GRADE_NUMS)]},
                                      legend_title_text='', font_family='Cairo')
                    st.plotly_chart(_f2, width='stretch', use_container_width=True)
                    show_df(_d2, 'توزيع المدارس المطابقة للشرط', 'توزيع_المدارس_المطابقة')

                    # ---- قائمة المدارس ----
                    _cols = [c for c in ['اسم المدرسة', 'المديرية', 'المحافظة', 'الجنس',
                                         'نوع التعليم', 'الملكية', 'أدنى صف', 'أعلى صف',
                                         'مجموع الطلاب', 'الطاقة الاستيعابية']
                             if c in _sel.columns]
                    _view = _sel[_cols].copy()
                    _view.insert(0, 'رقم أدنى صف', _sel['_lo'].map(grade_label))
                    _view.insert(1, 'رقم أعلى صف', _sel['_hi'].map(grade_label))
                    if 'المديرية' in _view.columns:
                        _view = _view.sort_values(['المديرية', 'اسم المدرسة'])
                    st.markdown(f"###### جدول المدارس المطابقة ({len(_view):,})")
                    show_df(_view, 'المدارس المطابقة', 'المدارس_المطابقة')

        elif selected_study == "امكانية نقل المدارس المستأجرة الى مدارس الملك":
            render_rented_to_owned(None, RENTED_TO_OWNED_RADIUS_KM)

        elif selected_study == "المسافة بين مدارس (ب الشارع)":
            st.markdown("#### المسافة بين مدارس — على الشارع")
            st.caption("اختر مدرسة أو أكثر في كل حقل. تُحسب المسافة المستقيمة "
                       "ومسافة الشارع الفعلية، ثم تُرسم المسارات على شبكة الشوارع.")

            _miss = [c for c in ['اسم المدرسة', 'x', 'y'] if c not in filtered_df.columns]
            if _miss:
                st.warning("أعمدة ناقصة: " + "، ".join(_miss))
            else:
                _d = filtered_df.copy()
                _d['اسم المدرسة'] = _d['اسم المدرسة'].astype(str)
                _d = _d[~_d['اسم المدرسة'].isin(['nan', 'None', ''])]
                _d['_x'] = pd.to_numeric(_d['x'], errors='coerce')
                _d['_y'] = pd.to_numeric(_d['y'], errors='coerce')
                _d = _d[_d['_x'].notna() & _d['_y'].notna()]
                _opts = sorted(_d['اسم المدرسة'].unique().tolist())

                if not _opts:
                    st.warning("لا توجد مدارس صالحة في البيانات المفلترة.")
                else:
                    mf1, mf2 = st.columns(2)
                    with mf1:
                        src_sel = st.multiselect(
                            "من مدرسة (اختر واحدة أو أكثر)", _opts, key="dist_from")
                    with mf2:
                        dst_sel = st.multiselect(
                            "إلى مدرسة (اختر واحدة أو أكثر)", _opts, key="dist_to")
                    st.caption(f"اخترت **{len(src_sel)}** من · **{len(dst_sel)}** إلى — "
                               f"كل زوج يُحسب كمسار مستقل.")

                    if st.button("احسب مسافات الشارع", key="dist_go", type="primary"):
                        rows, routes = [], []
                        with st.spinner("جارٍ جلب مسارات الشارع..."):
                            for a in src_sel:
                                ra = _d[_d['اسم المدرسة'] == a].iloc[0]
                                for b in dst_sel:
                                    if a == b:
                                        continue
                                    rb = _d[_d['اسم المدرسة'] == b].iloc[0]
                                    la1, lo1 = float(ra['_y']), float(ra['_x'])
                                    la2, lo2 = float(rb['_y']), float(rb['_x'])
                                    km = straight_distance_km(la1, lo1, la2, lo2)
                                    dm, ds, geom = road_route(la1, lo1, la2, lo2)
                                    rkm = None if dm is None else dm / 1000.0
                                    rows.append({
                                        'من': a,
                                        'إلى': b,
                                        'مسافة مستقيمة (كم)': round(km, 2),
                                        'مسافة الشارع (كم)': (None if rkm is None
                                                        else round(rkm, 2)),
                                        'زمن القيادة (دقيقة)': (None if ds is None
                                                        else round(ds / 60, 1)),
                                        'الفرق (كم)': (None if rkm is None
                                                  else round(rkm - km, 2)),
                                    })
                                    if geom and isinstance(geom, list) and geom \
                                            and isinstance(geom[0], list):
                                        routes.append({
                                            'from': a, 'to': b,
                                            'geom': geom,
                                            'straight_km': km,
                                            'road_km': rkm,
                                            'minutes': (None if ds is None else ds / 60.0),
                                        })

                        if rows:
                            res_df = pd.DataFrame(rows)
                            _ok = res_df['مسافة الشارع (كم)'].notna()
                            st.markdown("###### جدول المسافات")
                            show_df(res_df, 'جدول المسافات', 'جدول_المسافات')

                            if _ok.any():
                                _avg = res_df.loc[_ok, 'مسافة الشارع (كم)'].mean()
                                _str8 = res_df.loc[_ok, 'مسافة مستقيمة (كم)'].mean()
                                k1, k2, k3 = st.columns(3)
                                k1.metric("متوسط مسافة الشارع", f"{_avg:,.1f} كم")
                                k2.metric("متوسط المستقيمة", f"{_str8:,.1f} كم")
                                k3.metric("الزيادة على المستقيم",
                                          (f"{(_avg / _str8 - 1) * 100:,.0f}%"
                                           if _str8 else "—"))
                            if (~_ok).any():
                                st.warning("بعض المسارات لم يُرجعها خادم الشارع (OSRM) — "
                                           "ترى «None» في الجدول.")

                            if routes:
                                st.markdown("###### المسارات على شبكة الشوارع")
                                _lats = [q[0] for r in routes for q in r['geom']]
                                _lons = [q[1] for r in routes for q in r['geom']]
                                _mm = folium.Map(
                                    location=[sum(_lats) / len(_lats),
                                              sum(_lons) / len(_lons)],
                                    zoom_start=9, prefer_canvas=True)
                                folium.TileLayer(
                                    tiles='https://server.arcgisonline.com/ArcGIS/'
                                          'rest/services/World_Imagery/MapServer/'
                                          'tile/{z}/{y}/{x}',
                                    attr='Esri', name='ستلايت',
                                    overlay=True, control=True, show=False
                                ).add_to(_mm)
                                _pal = ['#2166ac', '#e75480', '#1a9850', '#ff8c00',
                                        '#7b3294', '#c51b7d', '#2b83ba']
                                for _i, _r in enumerate(routes):
                                    _c = _pal[_i % len(_pal)]
                                    _g = _r['geom']
                                    _rd = ("—" if _r['road_km'] is None
                                           else "%.1f كم" % _r['road_km'])
                                    _sd = "%.1f كم" % _r['straight_km']
                                    _mn = ("—" if _r['minutes'] is None
                                           else "%.0f د" % _r['minutes'])
                                    _tip = (
                                        f"<b>من:</b> {_r['from']}<br>"
                                        f"<b>إلى:</b> {_r['to']}<br>"
                                        f"<b>مسافة الشارع:</b> {_rd}<br>"
                                        f"<b>مسافة مستقيمة:</b> {_sd}<br>"
                                        f"<b>زمن القيادة:</b> {_mn}")

                                    folium.PolyLine(
                                        [[q[0], q[1]] for q in _g],
                                        color=_c, weight=5, opacity=0.9,
                                        tooltip=_tip).add_to(_mm)

                                    folium.CircleMarker(
                                        [_g[0][0], _g[0][1]], radius=8,
                                        color=_c, weight=2, opacity=1.0,
                                        fill=True, fill_color=_c,
                                        fill_opacity=0.95,
                                        popup=f"<b>من:</b> {_r['from']}"
                                        f"<br>مسافة الشارع: {_rd}").add_to(_mm)
                                    folium.CircleMarker(
                                        [_g[-1][0], _g[-1][1]], radius=8,
                                        color='#333333', weight=2, opacity=1.0,
                                        fill=True, fill_color='#ffffff',
                                        fill_opacity=0.95,
                                        popup=f"<b>إلى:</b> {_r['to']}"
                                        f"<br>مسافة الشارع: {_rd}").add_to(_mm)

                                    _mid = _g[len(_g) // 2]
                                    folium.Marker(
                                        [_mid[0], _mid[1]],
                                        icon=folium.DivIcon(
                                            html=(
                                                '<div style="font-size:12px;font-weight:800;'
                                                'color:#fff;background:%s;'
                                                'padding:3px 8px;border-radius:6px;'
                                                'border:2px solid #fff;'
                                                'white-space:nowrap;'
                                                'box-shadow:0 2px 5px rgba(0,0,0,.35)">'
                                                '%s · %s</div>' % (_c, _rd, _mn)),
                                            icon_size=(150, 24),
                                            icon_anchor=(75, 12)),
                                        tooltip=_tip).add_to(_mm)

                                folium.LayerControl().add_to(_mm)
                                try:
                                    _mm.fit_bounds([[min(_lats), min(_lons)],
                                                    [max(_lats), max(_lons)]],
                                                   padding=(40, 40))
                                except Exception:
                                    pass
                                components.html(_mm.get_root().render(),
                                                height=620, scrolling=False)
                                st.caption("المسارات على شبكة الشوارع الفعلية (OSRM). "
                                           "الدائرة الملوّنة = «من»، الدائرة البيضاء = «إلى»، "
                                           "والبطاقة في منتصف كل مسار تبيّن المسافة والزمن.")
                            else:
                                st.info("لم تُرجع مسارات شارع لعرضها.")
                        else:
                            st.info("اختر مدارس مختلفة في الحقلين لحساب المسافة.")
                    else:
                        st.info("اختر مدرسة واحدة على الأقل في كل حقل.")

        elif selected_study == NO_STUDY:
            st.markdown(
                '<div style="text-align:center; padding:18px 10px; opacity:.75">'
                '<div style="font-size:2rem">📚</div>'
                '<div style="font-weight:700; margin-top:6px">لا توجد دراسة معروضة</div>'
                '<div style="font-size:.85rem; margin-top:4px">'
                'اختر دراسة من القائمة أعلاه لعرض نتائجها</div></div>',
                unsafe_allow_html=True)
            if st.session_state.studies_shown:
                st.caption("دراسات مُلغاة سابقاً: " + " · ".join(st.session_state.studies_shown))
