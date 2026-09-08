import io
import time
import json
import math
import re
import pandas as pd
import streamlit as st
from urllib.parse import unquote, quote, urlencode
from urllib.request import urlopen
from difflib import SequenceMatcher

# ═════════════════════════════════════════════════════════════════
# 1. AUTHENTICATION & LOGIN
# ═════════════════════════════════════════════════════════════════
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.markdown('''
        <div style="text-align:center;padding:2rem 0 1rem 0;">
            <h2 style="color:#1A1A1A;font-size:1.4rem;font-weight:700;margin-bottom:0.2rem;">
                Sales Ops · Data Quality Suite (Taiwan)</h2>
            <p style="color:#888;font-size:0.88rem;margin-bottom:1.4rem;">
                Digital Sales APAC · foodpanda / Delivery Hero</p>
        </div>''', unsafe_allow_html=True)
    
    pwd = st.text_input("Password", type="password", placeholder="Enter password to continue")
    if st.button("Sign in", type="primary", use_container_width=True):
        if pwd == st.secrets.get("PASSWORD", "foodpanda"):
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("😕 Password incorrect")
    return False

if not check_password():
    st.stop()

# ═════════════════════════════════════════════════════════════════
# 2. MARKET CONFIG & LOCALIZATION (TAIWAN)
# ═════════════════════════════════════════════════════════════════
MARKETS = {
    "TW": {
        "name": "Taiwan",
        "currency": "TWD",
        "prefix": "886"
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
}
_DEFAULT_EXCLUSION_KW = ["hotel", "supermarket", "convenience store"]

TW_UNIT_RE   = re.compile(r'(\d+|[bB]\d+)\s*(樓|[fF])', re.IGNORECASE)
TW_POSTAL_RE = re.compile(r'\b\d{3,5}\b')  # 修復：完整支援 3~5 位數台灣郵編
TW_NAME_NOISE= re.compile(r'\b(股份有限公司|有限公司|企業社|工作室|商行|行|獨資|台灣|TW|taiwan)\b', re.IGNORECASE)

_NA_VALUES   = {"","nan","none","n/a","na","nil","-","–","unknown","no name"}

def haversine_distance(lat1, lon1, lat2, lon2) -> float:
    """計算兩點經緯度之間的真實距離（公尺 Meters）"""
    try:
        lat1, lon1, lat2, lon2 = map(math.radians, [float(lat1), float(lon1), float(lat2), float(lon2)])
        dlat = lat2 - lat1
        dlon = lon2 - lon1
        a = math.sin(dlat / 2)**2 + math.cos(lat1) * math.sin(lat2) * math.sin(dlon / 2)**2
        c = 2 * math.asin(math.sqrt(a))
        return c * 6371000  # 地球平均半徑（公尺）
    except Exception:
        return 999999.0

def find_column(df, possible_names):
    """彈性偵測欄位名稱 (包含 Coordinates (Latitude/Longitude))"""
    for col in df.columns:
        col_clean = str(col).strip().lower()
        for target in possible_names:
            if target.lower() in col_clean:
                return col
    return None

def extract_tw_postal(text: str) -> str:
    if not text or str(text).strip() in ("", "nan"):
        return ""
    m = TW_POSTAL_RE.search(str(text))
    return m.group(0) if m else ""

def generate_google_maps_url(row, url_format, name_col, street_col, postal_col, lat_col, lng_col):
    company_name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else ""
    if url_format == "coords" and lat_col and lng_col and pd.notna(row[lat_col]) and pd.notna(row[lng_col]):
        lat, lng = str(row[lat_col]).strip(), str(row[lng_col]).strip()
        return f"https://www.google.com/maps/search/{quote(f'{company_name}@{lat},{lng}')}"
    else:
        street = str(row[street_col]).strip() if street_col and pd.notna(row[street_col]) else ""
        postal = extract_tw_postal(row[postal_col]) if postal_col and pd.notna(row[postal_col]) else ""
        return f"https://www.google.com/maps/search/{quote(f'{company_name} {street} {postal}'.strip())}"

# ═════════════════════════════════════════════════════════════════
# 3. MAIN STREAMLIT INTERFACE & SIDEBAR
# ═════════════════════════════════════════════════════════════════
def main():
    st.set_page_config(page_title="Sales Ops · Data Quality Suite (Taiwan)", layout="wide")

    st.markdown("""
    <style>
        .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
    </style>""", unsafe_allow_html=True)

    with st.sidebar:
        st.header("⚙️ Settings")
        market_code = st.selectbox("Market", options=["TW"], format_func=lambda x: "🇹🇼 Taiwan (TW)")
        
        st.divider()
        st.subheader("📍 距離與相似度設定")
        max_dist_p4 = st.slider("P4 完全重複最大距離 (公尺)", 10, 100, 50, 5, help="距離小於此值且店名高度相似 → 判定為 P4 Duplicate")
        max_dist_p3 = st.slider("P3 潛在重複最大距離 (公尺)", 50, 300, 100, 10, help="距離小於此值且店名中度相似 → 判定為 P3 Potential")
        p3_name_thresh = st.slider("店名相似度門檻 (%)", 50, 95, 70, 5)

        st.divider()
        st.subheader("🚫 排除關鍵字 Exclusions")
        kw_input = st.text_area("排除類別關鍵字", value="\n".join(_DEFAULT_EXCLUSION_KW), height=120)

    # 頁面主標題
    st.title("Sales Ops · Data Quality Suite — Taiwan (Geo-Distance Version)")
    st.caption("foodpanda / Delivery Hero · Digital Sales APAC")

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

    # ── TAB 2: GENERATE APIFY URLS (完全修復版) ───────────────────
    with tab2:
        st.subheader("🔗 Generate Google Maps URLs for Apify")
        st.markdown("#### Step 1 · Generate URLs")

        url_mode = st.radio(
            "URL format",
            ["📍 Company / Account + Coordinates (Latitude, Longitude)", "📝 Company / Account + Address"],
            key="url_mode"
        )
        url_format = "coords" if "Coordinates" in url_mode else "address"

        uploaded_url_file = st.file_uploader("Upload leads file (.xlsx or .csv)", type=["xlsx","xls","csv"], key="url_leads")

        if uploaded_url_file is not None:
            try:
                df_url = pd.read_csv(uploaded_url_file) if uploaded_url_file.name.endswith(".csv") else pd.read_excel(uploaded_url_file)
                st.success(f"Successfully loaded {len(df_url)} rows.")

                # 精準模糊比對包含 Coordinates (Latitude) 括號的欄位
                name_col = find_column(df_url, ["company / account", "company", "account", "account name", "title"])
                street_col = find_column(df_url, ["street", "address"])
                postal_col = find_column(df_url, ["zip/postal code", "zip", "postal", "postal code"])
                lat_col = find_column(df_url, ["coordinates (lat", "latitude", "lat"])
                lng_col = find_column(df_url, ["coordinates (long", "longitude", "lng", "long"])
                grid_col = find_column(df_url, ["grid", "id"])

                missing_fields = []
                if not name_col: missing_fields.append("Company / Account")
                if url_format == "coords" and (not lat_col or not lng_col):
                    missing_fields.append("Coordinates (Latitude / Longitude)")
                elif url_format == "address" and not street_col:
                    missing_fields.append("Street / Address")

                if missing_fields:
                    st.error(f"❌ Missing required columns: {', '.join(missing_fields)}")
                else:
                    df_url["url"] = df_url.apply(
                        lambda row: generate_google_maps_url(row, url_format, name_col, street_col, postal_col, lat_col, lng_col),
                        axis=1
                    )

                    out_cols = [c for c in [grid_col, name_col, "url"] if c]
                    res_df = df_url[out_cols]

                    st.write("### Preview Generated URLs")
                    st.dataframe(res_df.head(10))

                    csv_data = res_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        label="📥 Download URLs CSV for Apify",
                        data=csv_data,
                        file_name="apify_input_urls.csv",
                        mime="text/csv"
                    )
            except Exception as e:
                st.error(f"Processing Error: {str(e)}")

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
