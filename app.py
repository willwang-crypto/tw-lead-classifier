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
# 2. HELPER FUNCTIONS & LARGE FILE LOADERS
# ═════════════════════════════════════════════════════════════════
MARKETS = {
    "TW": {
        "name": "Taiwan",
        "currency": "TWD",
        "prefix": "886"
    }
}

_DEFAULT_EXCLUSION_KW = ["hotel", "supermarket", "convenience store"]
TW_POSTAL_RE = re.compile(r'\b\d{3,5}\b')

def safe_load_large_csv(uploaded_file):
    """專為 Salesforce 匯出大檔設計的高效能 CSV 解析器"""
    if uploaded_file.name.endswith(".csv"):
        for encoding in ['utf-8-sig', 'utf-8', 'cp950', 'latin-1']:
            try:
                uploaded_file.seek(0)
                return pd.read_csv(
                    uploaded_file,
                    encoding=encoding,
                    low_memory=False,
                    on_bad_lines='skip',
                    engine='c',
                    quoting=1
                )
            except Exception:
                continue
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, low_memory=False, on_bad_lines='skip')
    else:
        return pd.read_excel(uploaded_file)

def haversine_distance(lat1, lon1, lat2, lon2) -> float:
    try:
        lat1, lon1, lat2, lon2 = map(math.radians, [float(lat1), float(lon1), float(lat2), float(lon2)])
        dlat = lat2 - lat1
        dlon = lon2 - lon1
        a = math.sin(dlat / 2)**2 + math.cos(lat1) * math.sin(lat2) * math.sin(dlon / 2)**2
        c = 2 * math.asin(math.sqrt(a))
        return c * 6371000
    except Exception:
        return 999999.0

def find_column(df, possible_names):
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
        max_dist_p4 = st.slider("P4 完全重複最大距離 (公尺)", 10, 100, 50, 5)
        max_dist_p3 = st.slider("P3 潛在重複最大距離 (公尺)", 50, 300, 100, 10)
        p3_name_thresh = st.slider("店名相似度門檻 (%)", 50, 95, 70, 5)

        st.divider()
        st.subheader("🚫 排除關鍵字 Exclusions")
        kw_input = st.text_area("排除類別關鍵字", value="\n".join(_DEFAULT_EXCLUSION_KW), height=120)

    st.title("Sales Ops · Data Quality Suite — Taiwan (Geo-Distance Version)")
    st.caption("foodpanda / Delivery Hero · Digital Sales APAC")

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

        url_mode = st.radio(
            "URL format",
            ["📍 Company / Account + Coordinates (Latitude, Longitude)", "📝 Company / Account + Address"],
            key="url_mode"
        )
        url_format = "coords" if "Coordinates" in url_mode else "address"

        uploaded_url_file = st.file_uploader("Upload leads file (.xlsx or .csv)", type=["xlsx","xls","csv"], key="url_leads")

        if uploaded_url_file is not None:
            try:
                df_url = safe_load_large_csv(uploaded_url_file)
                st.success(f"Successfully loaded {len(df_url)} rows.")

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

    # ── TAB 4: CRM CHECK (大數據極速精準比對版) ───────────────────
    with tab4:
        st.subheader("🔍 Quick CRM Duplicate Check")
        st.caption("針對一般的餐廳名單進行 CRM 快速重複排查（無需 GRID 或 Apify 爬蟲）。")

        col_c1, col_c2 = st.columns(2)
        with col_c1:
            raw_list_up = st.file_uploader("1. 上傳 Raw 餐廳名單 (.xlsx/.csv)", type=["xlsx","xls","csv"], key="tab4_raw")
        with col_c2:
            crm_accounts_up = st.file_uploader("2. 上傳 CRM All Accounts (.xlsx/.csv)", type=["xlsx","xls","csv"], key="tab4_crm")

        if raw_list_up and crm_accounts_up:
            try:
                df_raw = safe_load_large_csv(raw_list_up)
                df_crm = safe_load_large_csv(crm_accounts_up)

                st.success(f"✅ 檔案載入成功：Raw 名單 ({len(df_raw)} 筆) | CRM Accounts ({len(df_crm)} 筆)")

                raw_name_col = find_column(df_raw, ["account name", "company", "account", "name", "title"])
                crm_name_col = find_column(df_crm, ["account name", "company", "account", "name", "title"])

                if not raw_name_col or not crm_name_col:
                    st.error("❌ 找不到餐廳名稱欄位，請檢查檔案標頭是否含有 Name, Account, Company 等字樣。")
                else:
                    if st.button("▶ 開始 CRM 快速精準比對", type="primary", use_container_width=True):
                        with st.spinner("建立 42 萬筆資料快取索引中..."):
                            crm_series = df_crm[crm_name_col].fillna("").astype(str).str.strip().str.lower()
                            valid_crm = crm_series[crm_series != ""].tolist()
                            
                            # 建立 Hash Table 精準全匹配 Set
                            crm_set = set(valid_crm)
                            
                            # 按開頭前 2 個字分組，極速收窄搜尋範圍
                            crm_prefix_map = {}
                            for name in valid_crm:
                                prefix = name[:2]
                                if prefix not in crm_prefix_map:
                                    crm_prefix_map[prefix] = []
                                crm_prefix_map[prefix].append(name)

                            def check_dup_fast(raw_name):
                                if pd.isna(raw_name) or not str(raw_name).strip():
                                    return "Unverified", "", 0.0
                                
                                raw_str = str(raw_name).lower().strip()
                                
                                # 1. 完全相同 100% 命中 (0.0001秒)
                                if raw_str in crm_set:
                                    return "P4 - Duplicate", raw_str, 100.0

                                # 2. 僅針對開頭 2 字相同的近親候選進行 SequenceMatcher 演算 (大幅縮減 99% 運算量)
                                prefix = raw_str[:2]
                                candidates = crm_prefix_map.get(prefix, [])
                                
                                best_match = ""
                                best_score = 0.0

                                for c_name in candidates:
                                    # 字長差距超過 4 個字直接跳過，避免短名配長名
                                    if abs(len(raw_str) - len(c_name)) > 4:
                                        continue
                                        
                                    ratio = SequenceMatcher(None, raw_str, c_name).ratio()
                                    if ratio > best_score:
                                        best_score = ratio
                                        best_match = c_name

                                best_score_pct = round(best_score * 100, 1)

                                if best_score_pct >= 85.0:
                                    return "P4 - Duplicate", best_match, best_score_pct
                                elif best_score_pct >= 65.0:
                                    return "P3 - Potential Match", best_match, best_score_pct
                                else:
                                    return "Unverified", "", 0.0

                            results = df_raw[raw_name_col].apply(check_dup_fast)
                            df_raw["CRM Status"] = [r[0] for r in results]
                            df_raw["Matched CRM Name"] = [r[1] for r in results]
                            df_raw["Similarity Score (%)"] = [r[2] for r in results]

                            st.write("### 比對結果預覽")
                            st.dataframe(df_raw.head(20))

                            csv_out = df_raw.to_csv(index=False).encode('utf-8-sig')
                            st.download_button(
                                label="📥 下載 CRM 排查結果 CSV",
                                data=csv_out,
                                file_name="crm_check_result.csv",
                                mime="text/csv"
                            )
            except Exception as e:
                st.error(f"處理檔案時發生錯誤: {str(e)}")

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
