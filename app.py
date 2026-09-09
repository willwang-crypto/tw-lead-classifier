import pandas as pd
import streamlit as st
from difflib import SequenceMatcher
from urllib.parse import quote
import math
import re

# 1. 密碼驗證
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if st.session_state["password_correct"]:
        return True

    st.title("Sales Ops · Data Quality Suite")
    pwd = st.text_input("Password", type="password")
    if st.button("Sign in", type="primary"):
        if pwd == st.secrets.get("PASSWORD", "foodpanda"):
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("Password incorrect")
    return False

if not check_password():
    st.stop()

# 生鮮雜貨與非餐飲排除關鍵字
GROCERY_KEYWORDS = [
    "超市", "生鮮", "全聯", "家樂福", "聖德科斯", "水果行", "水果", 
    "肉品", "雜貨", "雜貨店", "便利商店", "超商", "7-11", "711", "全家", 
    "萊爾富", "ok超商", "美廉社", "有機", "農場", "水產", "Mart", "Grocery", 
    "Supermarket", "Convenience Store", "Market", "gift shop", "wedding store", "禮品"
]

def is_grocery(name_str):
    if not name_str or pd.isna(name_str): return False
    name_lower = str(name_str).lower()
    return any(kw.lower() in name_lower for kw in GROCERY_KEYWORDS)

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

def safe_load_csv(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        for enc in ['utf-8-sig', 'utf-8', 'cp950', 'latin-1']:
            try:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding=enc, low_memory=False, on_bad_lines='skip')
            except Exception:
                continue
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, low_memory=False, on_bad_lines='skip')
    else:
        return pd.read_excel(uploaded_file)

def find_column(df, possible_names):
    for col in df.columns:
        col_clean = str(col).strip().lower()
        for target in possible_names:
            if target.lower() in col_clean:
                return col
    return None

# Streamlit UI
st.set_page_config(page_title="Sales Ops Suite", layout="wide")
st.title("Sales Ops · Data Quality Suite — Taiwan")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Classify Leads", "🔗 Generate Apify URLs", "🏢 SF Account Audit",
    "🔍 CRM Check", "📋 KPI Sample Checker", "📖 How to Use"
])

# ── TAB 1: CLASSIFY LEADS (7 頁籤完整分類引擎) ────────────────────
with tab1:
    st.subheader("📊 Classify Leads (經緯度距離與 Google 分類引擎)")
    st.caption("結合經緯度 Haversine 距離與 Apify 抓取的 Google Maps 主類別進行最終判定。")

    c1, c2, c3 = st.columns(3)
    with c1:
        leads_up = st.file_uploader("1. 上傳 Unverified Leads 檔", type=["xlsx","xls","csv"], key="t1_leads")
    with c2:
        apify_up = st.file_uploader("2. 上傳含 GRID 的 Apify 結果", type=["xlsx","xls","csv"], key="t1_apify")
    with c3:
        crm_up = st.file_uploader("3. 上傳 CRM All Accounts 大檔", type=["xlsx","xls","csv"], key="t1_crm")

    if leads_up and apify_up and crm_up:
        st.success("✅ 3 個檔案皆已載入，點擊下方按鈕進行最終分類。")
        if st.button("▶ 開始分類比對 (Run Classification)", type="primary", use_container_width=True):
            with st.spinner("整合數據並執行經緯度與類別驗證中..."):
                try:
                    df_leads = safe_load_csv(leads_up)
                    df_apify = safe_load_csv(apify_up)
                    df_crm = safe_load_csv(crm_up)

                    # 欄位自動對應
                    l_name_col = find_column(df_leads, ["company / account", "company", "account", "name", "title"])
                    l_lat_col = find_column(df_leads, ["latitude", "lat", "coordinates (lat"])
                    l_lng_col = find_column(df_leads, ["longitude", "lng", "long", "coordinates (long"])
                    l_grid_col = find_column(df_leads, ["grid", "id", "account id", "lead id"])

                    a_grid_col = find_column(df_apify, ["grid", "id"])
                    a_cat_col = find_column(df_apify, ["categoryname", "categories", "primarycategory", "category"])
                    a_status_col = find_column(df_apify, ["permanentlyclosed", "temporarilyclosed", "isclosed", "status"])

                    c_name_col = find_column(df_crm, ["account name", "company", "account", "name"])
                    c_lat_col = find_column(df_crm, ["latitude", "lat"])
                    c_lng_col = find_column(df_crm, ["longitude", "lng", "long"])

                    # 建立 Apify 查表字典
                    apify_dict = {}
                    if a_grid_col and len(df_apify) > 0:
                        for _, a_row in df_apify.iterrows():
                            g_id = str(a_row[a_grid_col]).strip()
                            cat = str(a_row[a_cat_col]) if a_cat_col and pd.notna(a_row[a_cat_col]) else ""
                            is_closed = str(a_row[a_status_col]) if a_status_col and pd.notna(a_row[a_status_col]) else "False"
                            apify_dict[g_id] = {"category": cat, "is_closed": is_closed}

                    # 分類主邏輯
                    final_statuses = []
                    matched_crm_names = []
                    matched_dists = []

                    # 預處理 CRM 經緯度列表
                    crm_records = []
                    if c_name_col and c_lat_col and c_lng_col:
                        for _, c_row in df_crm.iterrows():
                            try:
                                crm_records.append((
                                    str(c_row[c_name_col]).lower().strip(),
                                    float(c_row[c_lat_col]),
                                    float(c_row[c_lng_col])
                                ))
                            except Exception:
                                continue

                    for idx, l_row in df_leads.iterrows():
                        l_name = str(l_row[l_name_col]).strip() if l_name_col and pd.notna(l_row[l_name_col]) else ""
                        l_grid = str(l_row[l_grid_col]).strip() if l_grid_col and pd.notna(l_row[l_grid_col]) else f"GRID_{idx}"

                        # 1. 檢查 Google 爬蟲類別與狀態
                        apify_info = apify_dict.get(l_grid, {})
                        cat_str = apify_info.get("category", "")
                        is_closed_str = apify_info.get("is_closed", "False")

                        if "true" in is_closed_str.lower() or "closed" in is_closed_str.lower():
                            final_statuses.append("Permanently Closed")
                            matched_crm_names.append("")
                            matched_dists.append(0)
                            continue

                        if is_grocery(l_name) or is_grocery(cat_str):
                            final_statuses.append("Wrong Target Group (WTG)")
                            matched_crm_names.append("")
                            matched_dists.append(0)
                            continue

                        # 2. 經緯度距離二次比對
                        best_status = "P1 - New Lead"
                        best_crm_name = ""
                        min_dist = 999999.0

                        try:
                            l_lat = float(l_row[l_lat_col])
                            l_lng = float(l_row[l_lng_col])

                            for c_name, c_lat, c_lng in crm_records:
                                dist = haversine_distance(l_lat, l_lng, c_lat, c_lng)
                                if dist < min_dist:
                                    min_dist = dist
                                    name_ratio = SequenceMatcher(None, l_name.lower(), c_name).ratio()
                                    
                                    if dist <= 50 and name_ratio >= 0.65:
                                        best_status = "P4 - Duplicate"
                                        best_crm_name = c_name
                                        break
                                    elif dist <= 100 and name_ratio >= 0.50:
                                        best_status = "P3 - Potential Match"
                                        best_crm_name = c_name
                        except Exception:
                            pass

                        final_statuses.append(best_status)
                        matched_crm_names.append(best_crm_name)
                        matched_dists.append(round(min_dist, 1) if min_dist < 99999 else None)

                    df_leads["Final Classification"] = final_statuses
                    df_leads["Matched CRM Name"] = matched_crm_names
                    df_leads["Min Distance (m)"] = matched_dists

                    st.write("### 📊 分類結果預覽")
                    st.dataframe(df_leads.head(20))

                    csv_out = df_leads.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 下載最終分類報告 (CSV)", data=csv_out, file_name="final_leads_classified.csv", mime="text/csv")

                except Exception as e:
                    st.error(f"分類過程發生錯誤: {str(e)}")

# ── TAB 2: GENERATE APIFY URLS & RE-ATTACH GRID ──────────────────
with tab2:
    st.subheader("🔗 Generate Google Maps URLs for Apify")
    st.markdown("#### Step 1 · Generate URLs")

    url_mode = st.radio(
        "URL format",
        ["📍 Company / Account + Coordinates", "📝 Company / Account + Address"],
        key="url_mode"
    )
    uploaded_url_file = st.file_uploader("Upload leads file (.csv/.xlsx)", type=["xlsx","xls","csv"], key="url_leads")

    if uploaded_url_file is not None:
        try:
            df_url = safe_load_csv(uploaded_url_file)
            st.success(f"Successfully loaded {len(df_url)} rows.")

            name_col = find_column(df_url, ["company / account", "company", "account", "account name", "name", "title"])
            street_col = find_column(df_url, ["street", "address", "地址"])
            grid_col = find_column(df_url, ["grid", "id", "account id", "lead id", "no"])

            if not name_col:
                st.error("❌ Missing required column: Company / Account Name")
            else:
                def make_url(r):
                    c_name = str(r[name_col]).strip() if pd.notna(r[name_col]) else ""
                    c_addr = str(r[street_col]).strip() if street_col and pd.notna(r[street_col]) else ""
                    return f"https://www.google.com/maps/search/{quote(f'{c_name} {c_addr}'.strip())}"

                df_url["url"] = df_url.apply(make_url, axis=1)
                
                if not grid_col:
                    df_url["GRID"] = [f"GRID_{i+1:06d}" for i in range(len(df_url))]
                    grid_col = "GRID"

                out_cols = [c for c in [grid_col, name_col, "url"] if c in df_url.columns]
                res_df = df_url[out_cols]

                st.dataframe(res_df.head(10))
                csv_data = res_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 Download URLs CSV for Apify", data=csv_data, file_name="apify_input_urls.csv", mime="text/csv")
        except Exception as e:
            st.error(f"Error: {str(e)}")

    st.divider()
    st.markdown("#### Step 2B · Re-attach GRID / Info to Apify Export")
    st.caption("將 Apify 爬蟲完成下載的 CSV 檔黏回原始識別碼，方便最後匯入 Tab 1。")

    col_a, col_b = st.columns(2)
    with col_a:
        orig_url_file = st.file_uploader("1. 上傳剛才下載的 apify_input_urls.csv 或 Unverified 名單", type=["csv", "xlsx"], key="s2b_orig")
    with col_b:
        apify_raw_file = st.file_uploader("2. 上傳從 Apify 爬好的結果 CSV", type=["csv", "xlsx"], key="s2b_raw")

    if orig_url_file and apify_raw_file:
        try:
            df_orig = safe_load_csv(orig_url_file)
            df_apify = safe_load_csv(apify_raw_file)

            grid_col = find_column(df_orig, ["grid", "id", "account id", "lead id"])
            orig_url_col = find_column(df_orig, ["url", "input_url", "searchurl"])
            apify_url_col = find_column(df_apify, ["searchstring", "inputurl", "url", "search_url"])

            if grid_col:
                target_id_series = df_orig[grid_col].astype(str)
            else:
                target_id_series = df_orig.iloc[:, 0].astype(str)

            if apify_url_col and orig_url_col:
                url_to_grid = dict(zip(df_orig[orig_url_col].astype(str), target_id_series))
                df_apify["GRID"] = df_apify[apify_url_col].astype(str).map(url_to_grid)
            else:
                df_apify["GRID"] = target_id_series.values[:len(df_apify)]

            st.success("✅ 識別碼（GRID/ID）補回完成！")
            st.dataframe(df_apify.head(10))
            
            csv_grid_out = df_apify.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 下載含 GRID 的 Apify 結果 CSV", data=csv_grid_out, file_name="apify_results_with_grid.csv", mime="text/csv")
        except Exception as e:
            st.error(f"處理失敗: {str(e)}")

# ── TAB 4: CRM CHECK ─────────────────────────────────────────────
with tab4:
    st.subheader("🔍 Quick CRM Duplicate Check")
    st.caption("自動排除生鮮雜貨/非餐廳店家 + 店名地址雙重比對。")

    col1, col2 = st.columns(2)
    with col1:
        raw_list_up = st.file_uploader("1. 上傳 Raw 餐廳名單 (.csv)", type=["csv", "xlsx"], key="t4_raw")
    with col2:
        crm_accounts_up = st.file_uploader("2. 上傳 CRM All Accounts (.csv)", type=["csv", "xlsx"], key="t4_crm")

    if raw_list_up and crm_accounts_up:
        try:
            df_raw = safe_load_csv(raw_list_up)
            df_crm = safe_load_csv(crm_accounts_up)

            st.success(f"✅ 載入成功：Raw 名單 ({len(df_raw)} 筆) | CRM Accounts ({len(df_crm)} 筆)")

            raw_name_col = find_column(df_raw, ["account name", "company", "account", "name", "title"])
            crm_name_col = find_column(df_crm, ["account name", "company", "account", "name", "title"])
            raw_addr_col = find_column(df_raw, ["street", "address", "地址"])
            crm_addr_col = find_column(df_crm, ["street", "address", "地址"])

            if not raw_name_col or not crm_name_col:
                st.error("❌ 找不到餐廳名稱欄位，請檢查標頭。")
            elif not raw_addr_col or not crm_addr_col:
                st.error("❌ 找不到地址欄位，請檢查標頭。")
            else:
                if st.button("▶ 開始 CRM 比對 (含生鮮雜貨過濾)", type="primary"):
                    with st.spinner("排查生鮮雜貨與比對中..."):
                        crm_names = df_crm[crm_name_col].fillna("").astype(str).str.strip().str.lower().tolist()
                        crm_addrs = df_crm[crm_addr_col].fillna("").astype(str).str.strip().str.lower().tolist()
                        crm_pairs = list(zip(crm_names, crm_addrs))

                        def check_single(row):
                            r_name = str(row[raw_name_col]).lower().strip() if pd.notna(row[raw_name_col]) else ""
                            r_addr = str(row[raw_addr_col]).lower().strip() if pd.notna(row[raw_addr_col]) else ""

                            if not r_name:
                                return "Unverified", "", 0.0

                            if is_grocery(r_name):
                                return "Wrong Target Group (Grocery)", "", 0.0

                            prefix = r_name[:2]
                            for c_name, c_addr in crm_pairs:
                                if not c_name.startswith(prefix):
                                    continue
                                
                                name_score = SequenceMatcher(None, r_name, c_name).ratio()
                                if name_score >= 0.70:
                                    addr_score = SequenceMatcher(None, r_addr, c_addr).ratio() if (r_addr and c_addr) else 0.0
                                    addr_contains = (r_addr in c_addr or c_addr in r_addr) if len(r_addr) > 5 and len(c_addr) > 5 else False

                                    if addr_score >= 0.60 or addr_contains:
                                        return "P4 - Duplicate", c_name, round(((name_score + max(addr_score, 0.8))/2)*100, 1)

                            return "Unverified", "", 0.0

                        results = df_raw.apply(check_single, axis=1)
                        df_raw["CRM Status"] = [r[0] for r in results]
                        df_raw["Matched CRM Name"] = [r[1] for r in results]
                        df_raw["Similarity Score (%)"] = [r[2] for r in results]

                        wtg_count = (df_raw["CRM Status"] == "Wrong Target Group (Grocery)").sum()
                        p4_count = (df_raw["CRM Status"] == "P4 - Duplicate").sum()
                        unv_count = (df_raw["CRM Status"] == "Unverified").sum()

                        st.write(f"### 📊 排查結果摘要：Unverified 新餐飲店 ({unv_count} 筆) | P4 重複 ({p4_count} 筆) | 🚫 生鮮雜貨過濾 ({wtg_count} 筆)")
                        st.dataframe(df_raw.head(20))

                        csv_out = df_raw.to_csv(index=False).encode('utf-8-sig')
                        st.download_button("📥 下載 CRM 排查結果 CSV", data=csv_out, file_name="crm_result_filtered.csv", mime="text/csv")
        except Exception as e:
            st.error(f"錯誤: {str(e)}")
