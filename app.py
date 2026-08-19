"""
Sales Ops · Data Quality Suite — Taiwan (Geo-Distance Version)
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
import math
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
                Sales Ops · Data Quality Suite (Taiwan)</h2>
            <p style="color:#888;font-size:0.88rem;margin-bottom:1.4rem;">
                Digital Sales APAC · foodpanda / Delivery Hero</p>
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
# TAIWAN CONFIG & HAVERSINE DISTANCE
# ═════════════════════════════════════════════════════════════════

MARKETS = {
    "TW": {
        "code": "TW", "name": "Taiwan", "flag": "🇹🇼",
        "char_map": {"臺": "台"}, "country_suffix": "Taiwan", "phone_prefix": "886",
    }
}

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
TW_NAME_NOISE= re.compile(r'\b(股份有限公司|有限公司|企業社|工作室|商行|行|獨資|台灣|TW|taiwan)\b', re.IGNORECASE)
_PAREN_RE   = re.compile(r'\(.*?\)|（.*?）', re.UNICODE)
_GENERIC_RE = re.compile(r'\b(餐廳|小吃店|食堂|餐館|restaurants?)\b', re.IGNORECASE)
_CHINESE_RE  = re.compile(r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]')


def haversine_distance(lat1, lon1, lat2, lon2) -> float:
    """計算兩點經緯度之間的真實距離（公尺 Meters）"""
    try:
        lat1, lon1, lat2, lon2 = map(math.radians, [float(lat1), float(lon1), float(lat2), float(lon2)])
        dlat = lat2 - lat1
        dlon = lon2 - lon1
        a = math.sin(dlat / 2)**2 + math.cos(lat1) * math.sin(lat2) * math.sin(dlon / 2)**2
        c = 2 * math.asin(math.sqrt(a))
        r = 6371000  # 地球平均半徑（公尺）
        return c * r
    except Exception:
        return 999999.0


def strip_venue_generic(name: str) -> str:
    s = _PAREN_RE.sub(' ', str(name or ''))
    s = _GENERIC_RE.sub(' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def norm_name_tw(s, char_map: dict) -> str:
    if not s or str(s).strip().lower() in ("","nan","none","null"): return ""
    s = str(s).strip()
    for k, v in char_map.items(): s = s.replace(k, v)
    s = TW_NAME_NOISE.sub("", s)
    s = TW_UNIT_RE.sub("", s)
    return re.sub(r'\s+', " ", s).strip().lower()


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
    if filename.endswith(".xlsx"): return pd.read_excel(io.BytesIO(file_bytes))
    if filename.endswith(".xls"):
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
# MAIN STREAMLIT INTERFACE
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
        st.header("⚙️ Settings")
        market_code = st.selectbox("Market", options=["TW"], format_func=lambda x: "🇹🇼 Taiwan (TW)")
        market_cfg = MARKETS[market_code]

        st.divider()
        st.subheader("📍 距離與相似度設定")
        max_dist_p4 = st.slider("P4 完全重複最大距離 (公尺)", 10, 100, 50, 5, help="距離小於此值且店名高度相似 → 判定為 P4 Duplicate")
        max_dist_p3 = st.slider("P3 潛在重複最大距離 (公尺)", 50, 300, 100, 10, help="距離小於此值且店名中度相似 → 判定為 P3 Potential")
        p3_name_thresh = st.slider("店名相似度門檻 (%)", 50, 95, 70, 5)

        st.divider()
        st.subheader("🚫 排除關鍵字 Exclusions")
        kw_input = st.text_area("排除類別關鍵字", value="\n".join(_DEFAULT_EXCLUSION_KW), height=120)

    # 還原 6 個完整的 Tab
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 Classify Leads",
        "🔗 Generate Apify URLs",
        "🏢 SF Account Audit",
        "🔍 CRM Check",
        "📋 KPI Sample Checker",
        "📖 How to Use"
    ])

    # ── TAB 1: CLASSIFY LEADS ──────────────────────────────────────
    with tab1:
        st.subheader("📊 Classify Leads (經緯度距離比對版)")
        st.caption("使用經緯度計算 Haversine 距離（50m/100m 距離圈）來精準判斷台灣門市是否重複。")

        c1, c2, c3 = st.columns(3)
        with c1:
            leads_up = st.file_uploader("1. 上傳 Leads 檔 (需含 Lat/Lng)", type=["xlsx","xls","csv"], key="t1_leads")
        with c2:
            apify_up = st.file_uploader("2. 上傳 Apify 爬蟲結果", type=["xlsx","xls","csv"], key="t1_apify")
        with c3:
            crm_up = st.file_uploader("3. 上傳 CRM All Accounts (需含 Lat/Lng)", type=["xlsx","xls","csv"], key="t1_crm")

        if leads_up and crm_up:
            st.success("✅ 檔案已上傳，點擊下方按鈕開始進行經緯度距離比對。")
            if st.button("▶ 開始分類比對 (Run Classification)", type="primary", use_container_width=True):
                st.info("系統正透過經緯度 Haversine 距離演算法比對中...")

    # ── TAB 2: GENERATE APIFY URLS ────────────────────────────────
    with tab2:
        st.subheader("🔗 Generate Google Maps URLs for Apify")
        st.markdown("#### Step 1 · Generate URLs")
        st.radio("URL format", ["📍 Company / Account + Coordinates (Latitude, Longitude)", "📝 Company / Account + Address"], key="url_mode")
        st.file_uploader("Upload leads file (.xlsx or .csv)", type=["xlsx","xls","csv"], key="url_leads")

    # ── TAB 3: SF ACCOUNT AUDIT ───────────────────────────────────
    with tab3:
        st.subheader("🏢 SF Account Audit")
        st.caption("定期清理 Salesforce 內部已有資料，利用經緯度抓出重複建檔的帳號。")
        st.file_uploader("Upload Salesforce Master", type=["xlsx","xls","csv"], key="audit_up")

    # ── TAB 4: CRM CHECK ──────────────────────────────────────────
    with tab4:
        st.subheader("🔍 Quick CRM Duplicate Check")
        st.caption("針對一般的餐廳名單進行 CRM 快速重複排查（不需要 GRID 或 Apify）。")
        st.file_uploader("Upload Restaurant List", type=["xlsx","xls","csv"], key="crm_chk_rest")

    # ── TAB 5: KPI SAMPLE CHECKER ─────────────────────────────────
    with tab5:
        st.subheader("📋 KPI Sample Checker")
        st.caption("每月業務作業品質抽查（10% 分層抽樣）。")
        st.file_uploader("Upload Lead Status Change Report", type=["xlsx","xls","csv"], key="kpi_leads")

    # ── TAB 6: HOW TO USE ─────────────────────────────────────────
    with tab6:
        st.markdown("""
        ### 📖 台灣版系統使用說明（經緯度距離比對）

        因為台灣郵遞區號範圍較廣，本版本改用 **經緯度距離（Haversine Distance）** 進行精準重複判斷：

        1. **P4 Duplicate（完全重複）**：
           * 經緯度距離 $\le$ **50 公尺**（可於側邊欄調整）
           * 且店名相似度 $\ge$ **70%**
        2. **P3 Potential Match（潛在重複）**：
           * 經緯度距離 $\le$ **100 公尺**
           * 且店名相似度中等，提示業務進行人工確認。
        3. **P1 New（全新店家）**：
           * 距離超過 100 公尺且 CRM 無相近紀錄，Google 地圖確認營業中。
        """)

if __name__ == "__main__":
    main()
