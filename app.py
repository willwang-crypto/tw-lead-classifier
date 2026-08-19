"""
Sales Ops · Data Quality Suite — Taiwan
Delivery Hero / Foodpanda · Digital Sales APAC

TAB 1  Lead Classification
       Postal + Name dedup
       Labels: P1 / P2 / P3 / P4 / Business Closed / Wrong Target Group

TAB 2  Generate Apify URLs
TAB 3  SF Account Audit
TAB 4  CRM Check
TAB 5  KPI Sample Checker
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


# ── Secrets & Caches ─────────────────────────────────────────────
APP_PASSWORD = st.secrets.get("APP_PASSWORD", "admin123")


def _load_logo():
    import os, base64
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dh_logo.png")
    return base64.b64encode(open(p, "rb").read()).decode() if os.path.exists(p) else ""

DH_LOGO_B64 = _load_logo()


def check_password():
    if st.session_state.get("authenticated"):
        return True
    _, col, _ = st.columns([1, 1.6, 1])
    with col:
        st.markdown(f'''
        <div style="text-align:center;padding:2rem 0 1rem 0;">
            <img src="data:image/png;base64,{DH_LOGO_B64}"
                 style="width:140px;margin-bottom:1.2rem;" />
            <h2 style="color:#1A1A1A;font-size:1.4rem;font-weight:700;margin-bottom:0.2rem;">
                Sales Ops · Data Quality Suite (Taiwan)</h2>
            <p style="color:#888;font-size:0.88rem;margin-bottom:1.4rem;">
                Digital Sales APAC · Foodpanda / Delivery Hero</p>
        </div>''', unsafe_allow_html=True)
        pwd = st.text_input("Password", type="password",
                            placeholder="Enter password to continue")
        if st.button("Sign in", type="primary", use_container_width=True):
            if pwd == APP_PASSWORD:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect password.")
    return False


st.set_page_config(page_title="Sales Ops Suite - Taiwan",
                   page_icon="🇹🇼", layout="wide",
                   initial_sidebar_state="expanded")


# ═════════════════════════════════════════════════════════════════
# MARKET CONFIG (TAIWAN)
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


# ═════════════════════════════════════════════════════════════════
# TAIWAN FOOD DELIVERY ELIGIBILITY & EXCLUSIONS
# ═════════════════════════════════════════════════════════════════

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
    # 台灣在地常見類別
    "便館","便當","熱炒","小吃","宵夜","鹽酥雞","滷味","雞排","手搖飲料","冰品店",
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


def is_food_delivery_eligible(category, exclusion_kw: list) -> bool:
    if not category or pd.isna(category):
        return False
    cat = str(category).strip()
    cat_lower = cat.lower()
    if any(kw in cat_lower for kw in exclusion_kw):
        return False
    if cat in FOOD_DELIVERY_ALLOWED:
        return True
    for allowed in FOOD_DELIVERY_ALLOWED:
        if allowed.lower() == cat_lower:
            return True
    food_kw = [
        "restaurant","cafe","café","bakery","pizza","sushi","burger",
        "grill","bistro","eatery","kitchen","takeaway","takeout","delivery",
        "noodle","ramen","snack","food truck","steakhouse","seafood",
        "sandwich","dim sum","hotpot","餐廳","餐館","小吃","烘焙","麵包",
        "便當","飲料","甜點","火鍋","燒肉","早午餐",
    ]
    return any(kw in cat_lower for kw in food_kw)


# ═════════════════════════════════════════════════════════════════
# TAIWAN UTILITY FUNCTIONS (Postal, Unit, Name Normalization)
# ═════════════════════════════════════════════════════════════════

# 台灣門牌/樓層正則表達式 (例: 3樓, 3F, B1)
TW_UNIT_RE   = re.compile(r'(\d+|[bB]\d+)\s*(樓|[fF])', re.IGNORECASE)
# 台灣郵遞區號正則表達式 (3碼或5/6碼, 例如 100, 10041, 100001)
TW_POSTAL_RE = re.compile(r'\b\d{3}(\d{2,3})?\b')
# 公司法人雜訊移除
TW_NAME_NOISE= re.compile(
    r'\b(股份有限公司|有限公司|企業社|工作室|商行|行|獨資|台灣|TW|taiwan)\b',
    re.IGNORECASE)

TW_CITIES = {
    "台北市","臺北市","新北市","基隆市","桃園市","新竹市","新竹縣","苗栗縣",
    "台中市","臺中市","彰化縣","南投縣","雲林縣","嘉義市","嘉義縣","台南市",
    "臺南市","高雄市","屏東縣","宜蘭縣","花蓮縣","台東縣","臺東縣","澎湖縣",
    "金門縣","連江縣","板橋","中和","永和","新莊","三重","桃園","中壢","竹北",
    "西門町","信義區","大安區","中山區","內湖區","士林區","文山區",
}

_PAREN_RE   = re.compile(r'\(.*?\)|（.*?）', re.UNICODE)
_GENERIC_RE = re.compile(r'\b(餐廳|小吃店|食堂|餐館|restaurants?)\b', re.IGNORECASE)

def strip_venue_generic(name: str) -> str:
    """移除括號內地名與通用餐廳字詞，例如: '火鍋店 (信義店)' → '火鍋店'"""
    s = _PAREN_RE.sub(' ', str(name or ''))
    s = _GENERIC_RE.sub(' ', s)
    return re.sub(r'\s+', ' ', s).strip()

_CHINESE_RE  = re.compile(r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]')
_NA_VALUES   = {"","nan","none","n/a","na","nil","-","–","unknown","no name"}


def extract_tw_unit(text: str) -> str:
    """擷取台灣樓層/門牌資訊，如 '3樓' 或 '3F'"""
    if not text or str(text).strip() in ("","nan"):
        return ""
    m = TW_UNIT_RE.search(str(text))
    if m:
        return f"{m.group(1).upper()}{m.group(2).upper()}"
    return ""


def extract_tw_postal(text: str) -> str:
    """擷取台灣郵遞區號 (前3碼為主要區域)"""
    if not text or str(text).strip() in ("","nan"):
        return ""
    m = TW_POSTAL_RE.search(str(text))
    if m:
        p = m.group(0)
        return p[:3]  # 統一取前 3 碼作為主要比對 Code
    return ""


def fix_postal_tw(postal) -> str:
    s = re.sub(r'\D', '', str(postal or "").strip().replace(".0",""))
    return s[:3] if len(s) >= 3 else s


def is_blank_name(s) -> bool:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return True
    return str(s).strip().lower() in _NA_VALUES


def has_chinese(text: str) -> bool:
    return bool(_CHINESE_RE.search(str(text or "")))


def to_pinyin(text: str) -> str:
    """將中文轉為漢語拼音（方便與英文名稱/拼音對照）"""
    if not text or is_blank_name(text):
        return ""
    text = str(text).strip()
    if not _PYPINYIN_AVAILABLE:
        return text.lower()
    result  = []
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
    """清洗台灣商家名稱，回傳 (latin_norm, pinyin_norm)"""
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


# ═════════════════════════════════════════════════════════════════
# PHONE & ADDRESS HELPERS (TAIWAN)
# ═════════════════════════════════════════════════════════════════

def norm_phone_tw(p, prefix: str = "886") -> str:
    """台灣電話標準化（將 09xx 或 02-xxxx 轉為 886xxxxxxxxx）"""
    if pd.isna(p): return ""
    s = str(p).replace("+","").replace(" ","").replace("-","").replace("(","").replace(")","").strip()
    if s.endswith(".0"): s = s[:-2]
    if s.startswith("0"):
        s = prefix + s[1:]
    elif not s.startswith(prefix):
        s = prefix + s
    return s


def to_e164_tw(p, prefix: str = "886") -> str:
    n = norm_phone_tw(p, prefix)
    return "+" + n if n else ""


def geocode_address_tw(street, postal_code, country_suffix="Taiwan", cache={}):
    """通用 OpenStreetMap / Photon 台灣地理編碼"""
    key = f"{street}|{postal_code}|{country_suffix}"
    if key in cache: return cache[key]
    parts = [p for p in [str(street).strip(), str(postal_code).strip(), "Taiwan"] if p and p != "nan"]
    if not parts: cache[key] = (None, None); return (None, None)
    time.sleep(1.0)
    try:
        q_str = ", ".join(parts)
        req = urlopen(
            f"https://photon.komoot.io/api/?q={urlencode({'q': q_str})}&limit=1",
            timeout=6)
        feats = json.loads(req.read().decode()).get("features",[])
        if feats:
            c = feats[0]["geometry"]["coordinates"]
            cache[key] = (float(c[1]), float(c[0])); return cache[key]
    except Exception: pass
    cache[key] = (None, None); return (None, None)
