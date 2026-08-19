"""
Sales Ops · Data Quality Suite — Taiwan
Delivery Hero / Foodpanda · Digital Sales APAC

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
                Sales Ops · Data Quality Suite (Taiwan)</h2>
            <p style="color:#888;font-size:0.88rem;margin-bottom:1.4rem;">
                Digital Sales APAC · Foodpanda / Delivery Hero</p>
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


# ── Market Config (Taiwan) ───────────────────────────────────────
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
TW_POSTAL_RE = re.compile(r'\b\d{3}(\d{2,3})?\b')
TW_NAME_NOISE= re.compile(r'\b(股份有限公司|有限公司|企業社|工作室|商行|行|獨資|台灣|TW|taiwan)\b', re.IGNORECASE)
_PAREN_RE   = re.compile(r'\(.*?\)|（.*?）', re.UNICODE)
_GENERIC_RE = re.compile(r'\b(餐廳|小吃店|食堂|餐館|restaurants?)\b', re.IGNORECASE)
_CHINESE_RE  = re.compile(r'[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]')


def extract_tw_postal(text: str) -> str:
    if not text or str(text).strip() in ("", "nan"):
        return ""
    m = TW_POSTAL_RE.search(str(text))
    return m.group(0)[:3] if m else ""


def extract_tw_unit(text: str) -> str:
    if not text or str(text).strip() in ("", "nan"):
        return ""
    m = TW_UNIT_RE.search(str(text))
    return f"{m.group(1).upper()}{m.group(2).upper()}" if m else ""


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


# ── Main Application ─────────────────────────────────────────────
def main():
    if not check_password():
        return

    st.title("🇹🇼 Sales Ops · Lead Classifier (Taiwan)")
    st.write("歡迎使用台灣版銷售名單清洗與重複資料比對系統。")

    with st.sidebar:
        st.header("⚙️ 設定 Settings")
        market_code = st.selectbox("Market", options=["TW"], format_func=lambda x: "🇹🇼 Taiwan (TW)")
        market_cfg = MARKETS[market_code]
        p2_threshold = st.slider("P3 潛在重複門檻 (Potential Match)", 0.30, 0.65, 0.50, 0.05)
        p3_threshold = st.slider("P4 完全重複門檻 (Duplicate)", float(round(p2_threshold + 0.05, 2)), 0.95, 0.75, 0.05)

    tab1, tab2, tab3 = st.tabs(["📊 檔案處理 (Classify Leads)", "🔗 產生搜尋網址 (Generate URLs)", "📖 使用說明 (How to Use)"])

    with tab1:
        st.subheader("1. 上傳檔案資料")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.file_uploader("上傳 Leads 檔 (.xlsx/.csv)", type=["xlsx", "csv"], key="t1_leads")
        with col2:
            st.file_uploader("上傳 Apify 爬蟲結果 (.xlsx/.csv)", type=["xlsx", "csv"], key="t1_apify")
        with col3:
            st.file_uploader("上傳 CRM All Accounts (.xlsx/.csv)", type=["xlsx", "csv"], key="t1_crm")

    with tab2:
        st.subheader("批次轉換 Google Maps 搜尋網址")
        st.info("請在此上傳您的 Salesforce Leads 檔，系統將自動拼接地址並產生可供 Apify 使用的搜尋網址。")

    with tab3:
        st.markdown("""
        ### 📖 系統操作說明
        1. **Classify Leads**：比對 Leads 與 CRM，自動歸類出 P1 (全新)、P3 (潛在重複)、P4 (已存在)、停業或非目標店家。
        2. **Generate Apify URLs**：匯出用於批次爬蟲的 Google 地圖連結。
        """)

if __name__ == "__main__":
    main():
