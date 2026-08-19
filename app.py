"""
Sales Ops · Data Quality Suite — Taiwan
foodpanda / Delivery Hero · Digital Sales APAC

TAB 1  Classify Leads
TAB 2  Generate Apify URLs
TAB 3  SF Account Audit
TAB 4  CRM Check
TAB 5  KPI Sample Checker
TAB 6  How to Use
"""

import streamlit as st
import pandas as pd
import re
import io
import time
import json
from urllib.parse import unquote, quote, urlencode
from urllib.request import urlopen
from difflib import SequenceMatcher
from itertools import combinations
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from shapely.wkt import loads as wkt_loads
from shapely.geometry import Point
from rapidfuzz import fuzz

try:
    from pypinyin import lazy_pinyin, Style as PinyinStyle
    _PYPINYIN_AVAILABLE = True
except ImportError:
    _PYPINYIN_AVAILABLE = False


# ── Password & Auth ─────────────────────────────────────────────
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "admin123")


def check_password():
    if st.session_state.get("authenticated"):
        return True
    _, col, _ = st.columns([1, 1.6, 1])
    with col:
        st.markdown('''
        <div style="text-align:center;padding:2rem 0 1rem 0;">
            <h2 style="color:#1A1A1A;font-size:1.4rem;font-weight:700;margin-bottom:0.2rem;">
                Sales Ops · Data Quality Suite</h2>
            <p style="color:#888;font-size:0.88rem;margin-bottom:1.4rem;">
                Digital Sales APAC · foodpanda / Delivery Hero (Taiwan)</p>
        </div>''', unsafe_allow_html=True)
        pwd = st.text_input("Password", type="password", placeholder="Enter password to continue")
        if st.button("Sign in", type="primary", use_container_width=True):
            if pwd == APP_PASSWORD:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect password.")
    return False


st.set_page_config(page_title="Sales Ops Suite - Taiwan", page_icon="🇹🇼", layout="wide", initial_sidebar_state="expanded")


# ═════════════════════════════════════════════════════════════════
# MARKET CONFIG & LOCALIZATION (TAIWAN)
# ═════════════════════════════════════════════════════════════════

MARKETS = {
    "TW": {
        "code": "TW", "name": "Taiwan", "flag": "🇹🇼",
        "char_map": {"臺": "台"}, "country_suffix": "Taiwan", "phone_prefix": "886",
    }
}

ACTIVE_PIPELINE = [
    "active", "new", "collecting documents", "negotiation",
    "menu processing", "onboarding", "quality check",
]
WIN_BACK        = ["lost", "terminated"]
WIN_BACK_FAILED = ["win back failed"]

FOOD_DELIVERY_ALLOWED = {
    "Restaurant","Fine dining restaurant","Family restaurant","Casual dining restaurant",
    "Buffet restaurant","Bistro","Eatery","Pizza restaurant","Sushi restaurant",
    "Ramen restaurant","Noodle restaurant","Dumpling restaurant","Dim sum restaurant",
    "Steak house","Grill restaurant","Barbecue restaurant","BBQ restaurant",
    "Sandwich shop","Salad shop","Breakfast restaurant","Brunch restaurant",
    "Dessert restaurant","Dessert shop","Ice cream shop","Donut shop",
    "Vegetarian restaurant","Vegan restaurant","Food hall","Food truck",
    "Cafe","Coffee shop","Tea house","Bubble tea shop","Boba shop","Bakery","Pastry shop",
    "Fast food restaurant","Hamburger restaurant","Fried chicken restaurant",
    "Taiwanese restaurant","Chinese restaurant","Japanese restaurant","Korean restaurant",
    "Thai restaurant","Vietnamese restaurant","Italian restaurant","American restaurant",
    "便當","熱炒","小吃","宵夜","鹽酥雞","滷味","雞排","手搖飲料","冰品店",
    "早餐店","早午餐","火鍋店","燒肉店","串燒店","居酒屋","咖啡廳","甜點店",
}

_DEFAULT_EXCLUSION_KW = [
    "美食街","美食中心","美食廣場","夜市攤位","百貨公司美食街",
    "飯店","酒店","KTV","卡拉OK","酒吧","夜店","酒館",
    "自動販賣機","超市","量販店","大買家","全聯","家樂福","美廉社",
    "便利商店","7-11","7-eleven","全家","萊爾富","ok超商",
    "寵物餐廳","寵物烘焙","學校食堂","醫院食堂","公司食堂","軍營食堂",
    "mix & match","婚宴會館","活動場地",
]

TW_UNIT_RE   = re.compile(r'(\d+|[bB]\d+)\s*(樓|[fF])', re.IGNORECASE)
TW_POSTAL_RE = re.compile(r'\b\d{3}(\d{2,3})?\b')
TW_NAME_NOISE= re.compile(r'\b(股份有限公司|有限公司|企業社|工作室|商行|行|獨資|台灣|TW|taiwan)\b', re.IGNORECASE)

TW_AREAS = {
    "台北","臺北","新北","基隆","桃園","新竹","苗栗","台中","臺中","彰化","南投",
    "雲林","嘉義","台南","臺南","高雄","屏東","宜蘭","花蓮","台東","臺東","澎湖",
    "金門","連江","信義區","大安區","中山區","內湖區","板橋","中和","永和","新莊","三重",
}

_PAREN_RE   = re.compile(r'\(.*?\)|（.*?）', re.UNICODE)
_GENERIC_RE = re.compile(r'\b(餐廳|小吃店|食堂|餐館|restaurants?)\b', re.IGNORECASE)
_CHINESE_RE  = re.compile(r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]')
_NA_VALUES   = {"","nan","none","n/a","na","nil","-","–","unknown","no name"}


def strip_venue_generic(name: str) -> str:
    s = _PAREN_RE.sub(' ', str(name or ''))
    s = _GENERIC_RE.sub(' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def extract_tw_unit(text: str) -> str:
    if not text or str(text).strip() in ("", "nan"):
        return ""
    m = TW_UNIT_RE.search(str(text))
    return f"{m.group(1).upper()}{m.group(2).upper()}" if m else ""


def extract_tw_postal(text: str) -> str:
    if not text or str(text).strip() in ("", "nan"):
        return ""
    m = TW_POSTAL_RE.search(str(text))
    return m.group(0)[:3] if m else ""


def _norm_postal_input(raw) -> str:
    s = re.sub(r'\D', '', str(raw or "").strip().replace(".0", ""))
    return s[:3] if len(s) >= 3 else s


def fix_postal_tw(postal) -> str:
    s = _norm_postal_input(postal)
    return s if s else str(postal).strip().replace(".0", "")


def is_blank_name(s) -> bool:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return True
    return str(s).strip().lower() in _NA_VALUES


def has_chinese(text: str) -> bool:
    return bool(_CHINESE_RE.search(str(text or "")))


def to_pinyin(text: str) -> str:
    if not text or is_blank_name(text):
        return ""
    text = str(text).strip()
    if not _PYPINYIN_AVAILABLE:
        return text.lower()
    result = []
    segment = []
    for char in text:
        if _CHINESE_RE.match(char):
            if segment:
                result.append("".join(segment))
                segment = []
            result.extend(lazy_pinyin(char, style=PinyinStyle.NORMAL))
        else:
            segment.append(char)
    if segment:
        result.append("".join(segment))
    return " ".join(result).lower().strip()


def norm_name_tw(s, char_map: dict) -> tuple:
    if is_blank_name(s):
        return "", ""
    s = str(s).strip()
    for k, v in char_map.items():
        s = s.replace(k, v)
    s = TW_NAME_NOISE.sub("", s)
    s = TW_UNIT_RE.sub("", s)
    s = re.sub(r'@\w+', "", s)
    s = re.sub(r'\s+', " ", s).strip()
    return s.lower().strip(), (to_pinyin(s) if has_chinese(s) else "")


def norm_phone_tw(p, prefix: str = "886") -> str:
    if pd.isna(p): return ""
    s = str(p).replace("+","").replace(" ","").replace("-","").replace("(","").replace(")","").strip()
    if s.endswith(".0"): s = s[:-2]
    if s.startswith("0"): s = prefix + s[1:]
    elif not s.startswith(prefix): s = prefix + s
    return s


def to_e164_tw(p, prefix: str = "886") -> str:
    n = norm_phone_tw(p, prefix)
    return "+" + n if n else ""


def norm_url(u) -> str:
    if pd.isna(u): return ""
    u = unquote(str(u).strip())
    u = u.split("?hl=")[0].split("&hl=")[0].split("&query_place_id=")[0]
    return u.lower()


def detect_column(df: pd.DataFrame, candidates: list):
    for c in candidates:
        if c in df.columns: return c
    lmap = {col.lower(): col for col in df.columns}
    for c in candidates:
        if c.lower() in lmap: return lmap[c.lower()]
    return None


@st.cache_data(show_spinner=False)
def _cached_read(file_bytes: bytes, filename: str) -> pd.DataFrame:
    from io import StringIO
    if filename.endswith(".xlsx"):
        return pd.read_excel(io.BytesIO(file_bytes))
    if filename.endswith(".xls"):
        header = file_bytes[:512].lstrip()
        is_html = (b"<" in header or b"<html" in header.lower() or b"<table" in header.lower())
        if is_html:
            try:
                tables = pd.read_html(io.BytesIO(file_bytes))
                if tables: return tables[0]
            except Exception: pass
            for enc in ("utf-8", "utf-8-sig", "windows-1252", "latin-1"):
                try:
                    raw = file_bytes.decode(enc)
                    sep = "\t" if raw.count("\t") > raw.count(",") else ","
                    return pd.read_csv(StringIO(raw), sep=sep, on_bad_lines="skip", engine="python")
                except Exception: continue
        else:
            try: return pd.read_excel(io.BytesIO(file_bytes), engine="xlrd")
            except Exception: pass
    for enc in ("utf-8","utf-8-sig","windows-1252","latin-1"):
        try:
            raw = file_bytes.decode(enc)
            sep = ";" if raw.count(";") > raw.count(",") else ","
            return pd.read_csv(StringIO(raw), sep=sep, quotechar='"', on_bad_lines="skip", engine="python")
        except Exception: continue
    return pd.DataFrame()


# ═════════════════════════════════════════════════════════════════
# MAIN APP ENTRY
# ═════════════════════════════════════════════════════════════════

def main():
    if not check_password():
        return

    st.markdown("""
    <style>
    html,body,[class*="css"]{font-family:Arial,sans-serif;}
    header[data-testid="stHeader"]{background:#FFF;border-bottom:2px solid rgba(223,16,103,.25);}
    [data-testid="stSidebar"]{background:#FAFAFA;border-right:2px solid rgba(223,16,103,.3);}
    [data-testid="stSidebar"] h1,[data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3{color:#DF1067!important;font-weight:700!important;}
    [data-testid="stButton"]>button[kind="primary"]{
        background:#DF1067!important;border:none!important;color:#fff!important;
        font-weight:600!important;font-size:1rem!important;
        border-radius:8px!important;padding:.6rem 1.5rem!important;}
    [data-testid="stMetric"]{background:#FFF;border:1px solid #EBEBEB;
        border-top:3px solid #DF1067;border-radius:8px;padding:.9rem 1rem;}
    </style>""", unsafe_allow_html=True)

    with st.sidebar:
        st.header("⚙️ 設定 Settings")
        market_code = st.selectbox("Market", options=["TW"], format_func=lambda x: "🇹🇼 Taiwan (TW)")
        market_cfg = MARKETS[market_code]

        st.divider()
        st.subheader("🎚 比對門檻 Thresholds")
        p2_threshold = st.slider("P3 潛在重複門檻 (Potential)", 0.30, 0.65, 0.50, 0.05)
        p3_threshold = st.slider("P4 完全重複門檻 (Duplicate)", float(round(p2_threshold + 0.05, 2)), 0.95, 0.75, 0.05)

        st.divider()
        st.subheader("🚫 排除關鍵字 Exclusions")
        kw_input = st.text_area("排除類別關鍵字", value="\n".join(_DEFAULT_EXCLUSION_KW), height=150)
        exclusion_kw = [k.strip().lower() for k in kw_input.split("\n") if k.strip()]

    tab1, tab2, tab3 = st.tabs(["📊 檔案處理 (Classify Leads)", "🔗 產生搜尋網址 (Generate URLs)", "📖 使用說明 (How to Use)"])

    with tab1:
        st.subheader("1. 上傳檔案資料")
        col1, col2, col3 = st.columns(3)
        with col1:
            leads_up = st.file_uploader("1. 上傳 Leads 檔 (.xlsx/.csv)", type=["xlsx", "xls", "csv"], key="t1_leads")
        with col2:
            apify_up = st.file_uploader("2. 上傳 Apify 爬蟲結果 (.xlsx/.csv)", type=["xlsx", "xls", "csv"], key="t1_apify")
        with col3:
            crm_up = st.file_uploader("3. 上傳 CRM All Accounts (.xlsx/.csv)", type=["xlsx", "xls", "csv"], key="t1_crm")

        if leads_up and crm_up:
            st.success("✅ 檔案已就緒，可點擊上方設定進行進一步處理。")

    with tab2:
        st.subheader("批次轉換 Google Maps 搜尋網址")
        st.info("請上傳您的 Salesforce Leads 檔，系統將自動拼接台灣地址並產生可供 Apify 使用的搜尋網址。")

    with tab3:
        st.markdown("""
        ### 📖 系統操作說明 (Taiwan Localized)
        1. **Classify Leads**：比對 Leads 與 CRM，自動歸類出 P1 (全新)、P3 (潛在重複)、P4 (已存在)、停業或非目標店家。
        2. **Generate Apify URLs**：匯出用於批次爬蟲的 Google 地圖連結。
        3. **在地化支援**：支援台灣 3 碼郵遞區號、樓層門牌（如 3樓/3F）以及台灣特有餐飲類別過濾。
        """)

if __name__ == "__main__":
    main()
