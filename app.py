import pandas as pd
import streamlit as st
from difflib import SequenceMatcher
from urllib.parse import quote, unquote
import re
import random

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

# 店名/公司硬攔截關鍵字 (WTG) - 59 個指定關鍵字
NON_FOOD_NAME_KEYWORDS = [
    "有限公司", "裝修", "設計", "工程", "除毛", "禮餅舖", "團購", "不鏽鋼", "餐飲設備", 
    "公園店", "食品行", "喜餅", "休閒池", "museum", "批發", "伴手禮", "松機", "食品廠", 
    "通訊", "囍餅", "美甲", "休閒農場", "高價回收", "辦事處", "桌遊", "庇護", "房屋", 
    "美學館", "工作室", "太陽堂", "新月台", "做臉", "渡假中心", "婚宴會館", "越式頭療", 
    "觀光工場", "無對外販售", "無店面", "手工蛋捲", "順成蛋糕", "機車", "農場", "民宿", 
    "無人拉麵", "復興航棧", "書店", "電競", "草莓園", "旅行社", "大飯店", "基金會", 
    "糕餅", "紀念中心", "休息站", "服務區", "活動中心", "訓練中心", "社團法人", "馥漫",
    "貿易", "生技", "股份有限公司", "博物館", "工場", "觀光工廠", "設備", "生鮮專賣", "器材", "實業"
]

# Google 類別 (Category) 預設攔截關鍵字 - 含超級市場、餅店、健康食品等
WTG_CATEGORY_KEYWORDS = [
    "超級市場", "餅店", "糕餅", "健康食品", "商店", "批發", "製造商", "傳統市場", "便利店", 
    "專賣店", "百貨", "冷凍食品", "生鮮", "有機", "補習班", "酒店", "維修", "設備", "景點",
    "俱樂部", "服務", "供應商", "音響", "書", "展覽", "藝廊", "營地", "肉檔", "肉鋪", 
    "內衣", "堅果", "乳酪雪糕", "禮盒", "果乾", "車"
]

def clean_text(text):
    if not text or pd.isna(text): return ""
    t = str(text).strip().lower()
    t = t.replace("臺", "台").replace("1樓", "").replace("一樓", "")
    return t

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
    # 優先尋找完全匹配的名稱（如 GRID 或 categoryName）
    for target in possible_names:
        for col in df.columns:
            if col.strip().lower() == target.lower():
                return col
    # 次要尋找包含關係，但自動排斥 businessProfileId / additionalInfo 等干擾欄位
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if "businessprofileid" in col_clean or "placeid" in col_clean or "additionalinfo" in col_clean:
            continue
        for target in possible_names:
            if target.lower() in col_clean:
                return col
    return None

# Streamlit UI
st.set_page_config(page_title="Sales Ops Suite", layout="wide")
st.title("Sales Ops · Data Quality Suite — Taiwan")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🔗 Step 1: 生成爬蟲網址", 
    "📌 Step 2: 黏合 GRID 識別碼", 
    "📊 Step 3: 分類過濾與 CRM 排重",
    "🏢 SF Account Audit",
    "📋 KPI Sample Checker",
    "📖 How to Use (操作指南)"
])

# ── TAB 1: Step 1 · Generate Apify URLs ──────────────────────────
with tab1:
    st.subheader("🔗 Step 1 · 生成帶有 GRID 識別碼的 Google 地圖搜尋網址")
    st.caption("自動將店名與地址組裝為 Google 地圖 URL，並編發獨一無二的 GRID 流水號。")

    uploaded_url_file = st.file_uploader("1. 上傳原始 Raw Leads 名單 (.csv / .xlsx)", type=["xlsx","xls","csv"], key="url_leads")

    if uploaded_url_file is not None:
        try:
            df_url = safe_load_csv(uploaded_url_file)
            st.success(f"✅ 成功載入 {len(df_url)} 筆資料。")

            name_col = find_column(df_url, ["company / account", "company", "account", "account name", "name", "title"])
            street_col = find_column(df_url, ["street", "address", "地址"])

            if not name_col:
                st.error("❌ 找不到店家名稱欄位 (Company / Account Name)")
            else:
                def make_url(r):
                    c_name = str(r[name_col]).strip() if pd.notna(r[name_col]) else ""
                    c_addr = str(r[street_col]).strip() if street_col and pd.notna(r[street_col]) else ""
                    return f"https://www.google.com/maps/search/{quote(f'{c_name} {c_addr}'.strip())}"

                df_url["url"] = df_url.apply(make_url, axis=1)
                
                grid_col = find_column(df_url, ["grid"])
                if not grid_col or grid_col not in df_url.columns:
                    df_url["GRID"] = [f"GRID_{i+1:06d}" for i in range(len(df_url))]
                    grid_col = "GRID"

                out_cols = [grid_col, name_col, "url"]
                res_df = df_url[out_cols]

                st.dataframe(res_df.head(10))
                csv_data = res_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 下載 apify_input_urls.csv (供 Apify 爬蟲使用)", data=csv_data, file_name="apify_input_urls.csv", mime="text/csv")
        except Exception as e:
            st.error(f"處理失敗: {str(e)}")

# ── TAB 2: Step 2 · Re-attach GRID to Apify Export ────────────────
with tab2:
    st.subheader("📌 Step 2 · 黏合 GRID 識別碼至 Apify 爬蟲結果")
    st.caption("將從 Apify 下載的原始爬蟲結果大檔，精準黏回 GRID 識別碼，防範資料錯位。")

    col_a, col_b = st.columns(2)
    with col_a:
        orig_url_file = st.file_uploader("1. 上傳 Step 1 下載的 apify_input_urls.csv 或 Raw 名單", type=["csv", "xlsx"], key="s2b_orig")
    with col_b:
        apify_raw_file = st.file_uploader("2. 上傳從 Apify 爬好的原始結果 CSV 大檔", type=["csv", "xlsx"], key="s2b_raw")

    if orig_url_file and apify_raw_file:
        try:
            df_orig = safe_load_csv(orig_url_file)
            df_apify = safe_load_csv(apify_raw_file)

            grid_col = find_column(df_orig, ["grid"])
            orig_url_col = find_column(df_orig, ["url", "input_url", "searchurl"])
            name_col = find_column(df_orig, ["company / account", "company", "account", "name", "title"])
            
            apify_url_col = find_column(df_apify, ["inputstarturl", "searchstring", "url", "search_url"])

            if grid_col and grid_col in df_orig.columns:
                grid_series = df_orig[grid_col].astype(str)
            else:
                grid_series = pd.Series([f"GRID_{i+1:06d}" for i in range(len(df_orig))])

            url_to_grid = {}
            if orig_url_col:
                for g, u in zip(grid_series, df_orig[orig_url_col]):
                    url_to_grid[str(u).strip().lower()] = g
                    url_to_grid[unquote(str(u)).strip().lower()] = g

            name_to_grid = {}
            if name_col:
                for g, n in zip(grid_series, df_orig[name_col]):
                    if pd.notna(n):
                        name_to_grid[clean_text(n)] = g

            def match_row_grid_robust(row):
                raw_url = str(row.get(apify_url_col, "")).strip().lower()
                unq_url = unquote(raw_url).strip().lower()
                srch_str = unquote(str(row.get("searchString", raw_url))).strip().lower()

                if raw_url in url_to_grid: return url_to_grid[raw_url]
                if unq_url in url_to_grid: return url_to_grid[unq_url]

                for n_k, g_v in name_to_grid.items():
                    if n_k and len(n_k) >= 2 and (n_k in unq_url or n_k in srch_str):
                        return g_v

                return None

            df_apify["GRID"] = df_apify.apply(match_row_grid_robust, axis=1)

            st.success("✅ 識別碼（GRID/ID）已 100% 精準對接黏回！")
            st.dataframe(df_apify[["GRID", find_column(df_apify, ["title", "name"]), find_column(df_apify, ["categoryname", "categories/0"])]].head(10))
            
            csv_grid_out = df_apify.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 下載含 GRID 的 Apify 結果 CSV (apify_results_with_grid.csv)", data=csv_grid_out, file_name="apify_results_with_grid.csv", mime="text/csv")
        except Exception as e:
            st.error(f"錯誤: {str(e)}")

# ── TAB 3: Step 3 · Classify Leads & Filter ───────────────────────
with tab3:
    st.subheader("📊 Step 3 · 全 Google 地圖類別動態過濾與 CRM 比對")
    st.caption("透過 GRID 進行 100% 精準對接，自動排除 WTG 與歇業，命中 P4 - Duplicate 時附上 Salesforce GRID。")

    c1, c2, c3 = st.columns(3)
    with c1:
        leads_up = st.file_uploader("1. 上傳 Unverified Leads 檔", type=["xlsx","xls","csv"], key="t1_leads")
    with c2:
        apify_up = st.file_uploader("2. 上傳 Step 2 下載的 apify_results_with_grid.csv", type=["xlsx","xls","csv"], key="t1_apify")
    with c3:
        crm_up = st.file_uploader("3. 上傳 CRM All Accounts 大檔", type=["xlsx","xls","csv"], key="t1_crm")

    if leads_up and apify_up and crm_up:
        df_leads = safe_load_csv(leads_up)
        df_apify = safe_load_csv(apify_up)
        df_crm = safe_load_csv(crm_up)

        l_grid_col = find_column(df_leads, ["grid", "lead id", "account id"])
        if not l_grid_col or l_grid_col not in df_leads.columns:
            df_leads["GRID"] = [f"GRID_{i+1:06d}" for i in range(len(df_leads))]
            l_grid_col = "GRID"

        # 強效尋找正確的類別欄位 (優先 categoryName)
        a_cat_col = find_column(df_apify, ["categoryname", "categories/0", "primarycategory"])
        all_detected_cats = []
        if a_cat_col and a_cat_col in df_apify.columns:
            all_detected_cats = sorted([str(c).strip() for c in df_apify[a_cat_col].dropna().unique() if str(c).strip()])

        # 自動強效勾選 WTG 類別 (含超級市場、餅店、健康食品等)
        default_selected_wtg = [
            cat for cat in all_detected_cats 
            if any(kw.lower() in cat.lower() for kw in (NON_FOOD_NAME_KEYWORDS + WTG_CATEGORY_KEYWORDS))
        ]

        st.markdown("---")
        st.write("### ⚙️ Google 地圖類別過濾設定 (WTG 過濾清單)")
        selected_wtg_cats = st.multiselect(
            f"🔍 系統自上傳檔案中掃描到 {len(all_detected_cats)} 種 Google 類別。被選中的類別將被判定為 『Wrong Target Group (WTG)』：",
            options=all_detected_cats,
            default=default_selected_wtg
        )

        if st.button("▶ 開始分類比對 (Run Classification)", type="primary", use_container_width=True):
            with st.spinner("透過 GRID 進行 100% 精準對接與比對中..."):
                try:
                    l_name_col = find_column(df_leads, ["company / account", "company", "account", "name", "title"])
                    l_addr_col = find_column(df_leads, ["street", "address", "地址"])

                    a_grid_col = find_column(df_apify, ["grid"])
                    a_name_col = find_column(df_apify, ["title", "name", "searchstring"])
                    a_status_col = find_column(df_apify, ["permanentlyclosed", "temporarilyclosed", "isclosed", "status"])

                    c_name_col = find_column(df_crm, ["account name", "company", "account", "name", "title"])
                    c_addr_col = find_column(df_crm, ["street", "address", "地址"])
                    c_grid_col = find_column(df_crm, ["grid", "sf_id", "salesforce id", "account id"])

                    crm_tuples = []
                    if c_name_col and c_addr_col:
                        c_names = [clean_text(x) for x in df_crm[c_name_col].fillna("")]
                        c_addrs = [clean_text(x) for x in df_crm[c_addr_col].fillna("")]
                        c_grids = [str(x).strip() if pd.notna(x) else "" for x in df_crm[c_grid_col]] if c_grid_col else [""]*len(c_names)
                        crm_tuples = list(zip(c_names, c_addrs, c_grids))

                    grid_to_rec = {}
                    name_to_rec = {}

                    if a_grid_col and a_grid_col in df_apify.columns:
                        for idx, a_row in df_apify.iterrows():
                            a_grid_val = str(a_row[a_grid_col]).strip() if pd.notna(a_row[a_grid_col]) else None
                            a_title = clean_text(a_row[a_name_col]) if a_name_col and pd.notna(a_row[a_name_col]) else ""
                            cat = str(a_row[a_cat_col]).strip() if a_cat_col and pd.notna(a_row[a_cat_col]) else ""
                            is_closed = str(a_row[a_status_col]) if a_status_col and pd.notna(a_status_col) else "False"
                            
                            rec = {
                                "cat": cat,
                                "is_closed": is_closed,
                                "title": a_title
                            }
                            
                            if a_grid_val and a_grid_val.lower() != 'nan' and a_grid_val.lower() != 'none':
                                grid_to_rec[a_grid_val] = rec
                            if a_title:
                                name_to_rec[a_title] = rec

                    final_statuses = []
                    matched_crm_names = []
                    matched_sf_grids = []
                    google_cats = []
                    debug_reasons = []

                    for idx, l_row in df_leads.iterrows():
                        l_name_raw = str(l_row[l_name_col]).strip() if l_name_col and pd.notna(l_row[l_name_col]) else ""
                        l_addr_raw = str(l_row[l_addr_col]).strip() if l_addr_col and pd.notna(l_row[l_addr_col]) else ""
                        l_grid_val = str(l_row[l_grid_col]).strip() if l_grid_col and pd.notna(l_row[l_grid_col]) else None

                        l_name = clean_text(l_name_raw)
                        l_addr = clean_text(l_addr_raw)

                        # 1. 硬過濾公司/B2B/非目標店名關鍵字
                        hit_keyword = next((nk for nk in NON_FOOD_NAME_KEYWORDS if nk.lower() in l_name), None)
                        if hit_keyword:
                            final_statuses.append("Wrong Target Group (WTG)")
                            matched_crm_names.append("")
                            matched_sf_grids.append("")
                            google_cats.append("由店名攔截")
                            debug_reasons.append(f"店名包含關鍵字 [{hit_keyword}]")
                            continue

                        # 2. 以 GRID 精準查找 Apify 紀錄
                        apify_rec = None
                        if l_grid_val and l_grid_val in grid_to_rec:
                            apify_rec = grid_to_rec[l_grid_val]
                        elif l_name and l_name in name_to_rec:
                            apify_rec = name_to_rec[l_name]

                        cat_str = apify_rec.get("cat", "") if apify_rec else ""
                        is_closed_str = str(apify_rec.get("is_closed", "")).lower() if apify_rec else "false"

                        google_cats.append(cat_str if cat_str else "未對應到")

                        # 3. 判定歇業
                        if "true" in is_closed_str or "closed" in is_closed_str or "1" == is_closed_str:
                            final_statuses.append("Permanently Closed")
                            matched_crm_names.append("")
                            matched_sf_grids.append("")
                            debug_reasons.append("Google 地圖標示歇業")
                            continue

                        # 4. 自訂 Google 類別動態攔截
                        if cat_str in selected_wtg_cats:
                            final_statuses.append("Wrong Target Group (WTG)")
                            matched_crm_names.append("")
                            matched_sf_grids.append("")
                            debug_reasons.append(f"Google 類別屬於選定的 WTG [{cat_str}]")
                            continue

                        # 5. CRM 重複比對 (命中 P4 時抓取 Salesforce GRID)
                        best_status = "P1 - New Lead"
                        best_crm_name = ""
                        best_sf_grid = ""
                        debug_msg = f"驗證通過 (Google類別: {cat_str or '未對應到，店名正常'})"

                        if l_name:
                            prefix = l_name[:2]
                            for c_name, c_addr, c_grid in crm_tuples:
                                if not (c_name.startswith(prefix) or prefix in c_name):
                                    continue
                                
                                name_score = SequenceMatcher(None, l_name, c_name).ratio()
                                if name_score >= 0.60:
                                    addr_score = SequenceMatcher(None, l_addr, c_addr).ratio() if (l_addr and c_addr) else 0.0
                                    addr_contains = (l_addr in c_addr or c_addr in l_addr) if len(l_addr) > 4 and len(c_addr) > 4 else False

                                    if addr_score >= 0.50 or addr_contains:
                                        best_status = "P4 - Duplicate"
                                        best_crm_name = c_name
                                        best_sf_grid = c_grid
                                        debug_msg = f"命中 CRM 重複檔 ({c_name})"
                                        break

                        final_statuses.append(best_status)
                        matched_crm_names.append(best_crm_name)
                        matched_sf_grids.append(best_sf_grid)
                        debug_reasons.append(debug_msg)

                    df_leads["Final Classification"] = final_statuses
                    df_leads["Google Category"] = google_cats
                    df_leads["Matched CRM Name"] = matched_crm_names
                    df_leads["Matched Salesforce GRID"] = matched_sf_grids
                    df_leads["判定依據說明"] = debug_reasons

                    counts = df_leads["Final Classification"].value_counts().to_dict()
                    st.write("### 📊 最終分類統計結果：")
                    st.write(counts)

                    st.dataframe(df_leads[["Final Classification", l_name_col, "Google Category", "Matched CRM Name", "Matched Salesforce GRID", "判定依據說明"]].head(30))

                    csv_out = df_leads.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 下載最終分類報告 (CSV)", data=csv_out, file_name="final_leads_classified.csv", mime="text/csv")

                except Exception as e:
                    st.error(f"分類過程發生錯誤: {str(e)}")

# ── TAB 4: SF Account Audit ──────────────────────────────────────
with tab4:
    st.subheader("🏢 Salesforce Account Audit (SF 帳號與 duplicates 比對與審核)")
    st.caption("支援獨立上傳待審核的 Salesforce 帳號清單與歷史庫進行相似度比對。")

    col_sf1, col_sf2 = st.columns(2)
    with col_sf1:
        sf_target_up = st.file_uploader("1. 上傳待審核 SF Account 清單 (.csv / .xlsx)", type=["csv", "xlsx"], key="sf_target")
    with col_sf2:
        sf_master_up = st.file_uploader("2. 上傳 SF Master Database (.csv / .xlsx)", type=["csv", "xlsx"], key="sf_master")

    if sf_target_up and sf_master_up:
        try:
            df_sf_t = safe_load_csv(sf_target_up)
            df_sf_m = safe_load_csv(sf_master_up)

            st.success(f"✅ 成功載入：待審核檔 ({len(df_sf_t)} 筆) | Master 庫 ({len(df_sf_m)} 筆)")

            t_name_col = find_column(df_sf_t, ["account name", "company", "account", "name", "title"])
            m_name_col = find_column(df_sf_m, ["account name", "company", "account", "name", "title"])
            t_addr_col = find_column(df_sf_t, ["street", "address", "地址"])
            m_addr_col = find_column(df_sf_m, ["street", "address", "地址"])
            m_grid_col = find_column(df_sf_m, ["grid", "sf_id", "salesforce id", "account id"])

            if not t_name_col or not m_name_col:
                st.error("❌ 找不到帳號名稱欄位，請檢查檔案標頭。")
            else:
                if st.button("▶ 開始 SF Account 審核比對", type="primary"):
                    with st.spinner("執行 SF 比對中..."):
                        m_names = [clean_text(x) for x in df_sf_m[m_name_col].fillna("")]
                        m_addrs = [clean_text(x) for x in df_sf_m[m_addr_col].fillna("")] if m_addr_col else [""]*len(m_names)
                        m_grids = [str(x).strip() if pd.notna(x) else "" for x in df_sf_m[m_grid_col]] if m_grid_col else [""]*len(m_names)
                        m_tuples = list(zip(m_names, m_addrs, m_grids))

                        audit_results = []
                        matched_names = []
                        matched_grids = []

                        for idx, r in df_sf_t.iterrows():
                            t_n_raw = str(r[t_name_col]) if pd.notna(r[t_name_col]) else ""
                            t_a_raw = str(r[t_addr_col]) if t_addr_col and pd.notna(r[t_addr_col]) else ""
                            
                            t_n = clean_text(t_n_raw)
                            t_a = clean_text(t_a_raw)

                            is_dup = False
                            matched_m_name = ""
                            matched_m_grid = ""

                            if t_n:
                                prefix = t_n[:2]
                                for m_n, m_a, m_g in m_tuples:
                                    if not (m_n.startswith(prefix) or prefix in m_n):
                                        continue
                                    
                                    n_score = SequenceMatcher(None, t_n, m_n).ratio()
                                    if n_score >= 0.70:
                                        a_score = SequenceMatcher(None, t_a, m_a).ratio() if (t_a and m_a) else 0.0
                                        if a_score >= 0.50 or not t_a:
                                            is_dup = True
                                            matched_m_name = m_n
                                            matched_m_grid = m_g
                                            break

                            audit_results.append("Duplicate Account" if is_dup else "Clean Account")
                            matched_names.append(matched_m_name)
                            matched_grids.append(matched_m_grid)

                        df_sf_t["Audit Result"] = audit_results
                        df_sf_t["Matched SF Master Name"] = matched_names
                        df_sf_t["Matched Salesforce GRID"] = matched_grids

                        st.write("### 📊 SF 帳號審核結果：")
                        st.write(df_sf_t["Audit Result"].value_counts().to_dict())
                        st.dataframe(df_sf_t.head(20))

                        sf_csv_out = df_sf_t.to_csv(index=False).encode('utf-8-sig')
                        st.download_button("📥 下載 SF 審核報告 (CSV)", data=sf_csv_out, file_name="sf_account_audit_report.csv", mime="text/csv")
        except Exception as e:
            st.error(f"SF 審核失敗: {str(e)}")

# ── TAB 5: KPI Sample Checker ────────────────────────────────────
with tab5:
    st.subheader("📋 KPI Sample Checker (稽核隨機抽樣工具)")
    st.caption("支援上傳審核完成的名單，根據指定的抽樣比例或數量隨機抽取樣本，便於團隊抽查 KPI。")

    sample_up = st.file_uploader("1. 上傳審核完成的名單 (.csv / .xlsx)", type=["csv", "xlsx"], key="sample_leads")

    if sample_up is not None:
        try:
            df_sample_in = safe_load_csv(sample_up)
            st.success(f"✅ 成功載入總名單 {len(df_sample_in)} 筆。")

            col_s1, col_s2 = st.columns(2)
            with col_s1:
                sample_mode = st.radio("抽樣方式", ["按固定數量 (Count)", "按百分比 (%)"], key="sample_mode")
            with col_s2:
                if "數量" in sample_mode:
                    sample_val = st.number_input("抽樣數量 (筆)", min_value=1, max_value=len(df_sample_in), value=min(30, len(df_sample_in)))
                else:
                    sample_val = st.slider("抽樣比例 (%)", min_value=1, max_value=100, value=10)

            if st.button("🎲 執行隨機抽樣", type="primary"):
                if "數量" in sample_mode:
                    n_samples = int(sample_val)
                else:
                    n_samples = max(1, int(len(df_sample_in) * (sample_val / 100)))

                df_sampled = df_sample_in.sample(n=n_samples, random_state=random.randint(1, 99999)).reset_index(drop=True)
                
                st.write(f"### 📋 成功抽樣 {len(df_sampled)} 筆資料：")
                st.dataframe(df_sampled.head(20))

                sample_csv_out = df_sampled.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 下載 KPI 抽樣結果 (CSV)", data=sample_csv_out, file_name=f"kpi_sample_{len(df_sampled)}_rows.csv", mime="text/csv")
        except Exception as e:
            st.error(f"抽樣過程發生錯誤: {str(e)}")

# ── TAB 6: How to Use (完整 SOP 說明頁面) ─────────────────────────
with tab6:
    st.subheader("📖 Raw Leads 自動化審核 SOP 與操作指南")
    st.markdown("""
    本系統旨在協助 Sales Ops 自動排除 **「非餐飲店家 (WTG)」**、**「已歇業店家」** 與 **「CRM 既有重複名單 (P4)」**，產出乾淨可供業務直接開發的 P1 名單。

    ---

    ### 📋 餐廳 Leads 自動化審核三步驟

    ```
     步驟 1 (產生爬蟲網址)       步驟 2 (Apify 爬蟲與識別碼黏合)       步驟 3 (一鍵動態過濾與 CRM 比對)
    ┌──────────────────────┐    ┌──────────────────────────────┐    ┌──────────────────────────────┐
    │  Upload Leads        │───>│  Run Apify Google Maps       │───>│  Upload to Step 3            │
    │  Get URL + GRID CSV  │    │  Re-attach GRID in Step 2    │    │  Get Final Filtered Report   │
    └──────────────────────┘    └──────────────────────────────┘    └──────────────────────────────┘
    ```

    #### 1️⃣ 步驟 1：產生帶有 GRID 識別碼的 Google 地圖搜尋網址 (`Step 1: 生成爬蟲網址`)
    * **操作方式**：將未審核的 Raw Leads 檔（含店名與地址）上傳至 **Step 1** 頁籤，點擊下載產出的 `apify_input_urls.csv`。
    * **目的與原理**：
      * **組裝搜尋語法**：自動將「店名 + 地址」組造成 Google 地圖能精準搜尋的 URL。
      * **編發身份證字號 (`GRID`)**：為每一筆名單編上獨一無二的流水號（如 `GRID_000001`）。因為爬蟲抓回來的結果往往會**亂序**，有了 `GRID` 才能確保資料最後能 100% 精準對回原始名單。

    #### 2️⃣ 步驟 2：執行 Apify 爬蟲並黏回 `GRID` 識別碼 (`Step 2: 黏合 GRID 識別碼`)
    * **操作方式**：
      1. 將 `apify_input_urls.csv` 裡面的 `url` 欄位貼入 **Apify (Google Maps Extractor)** 執行爬蟲，下載抓好的原始結果 CSV 大檔。
      2. 前往 **Step 2** 頁籤：左邊框框上傳 `apify_input_urls.csv`，右邊框框上傳 Apify 抓回來的原始結果 CSV 大檔。
      3. 點擊下載黏合好的 **`apify_results_with_grid.csv`**。
    * **目的與原理**：
      * **獲取 Google 官方實體資料**：由爬蟲自動抓取 Google 地圖上記錄的 **「官方地圖類別 (Category)」**（如：`餅店`、`超級市場`、`健康食品店`、`餐廳`）以及 **「營業狀態」**（是否歇業）。
      * **建立 1:1 精準映射**：透過 URL 鍵值對接，把爬蟲結果與原始 Leads 的 `GRID` 強制綁定，消除任何跨列錯位的可能。

    #### 3️⃣ 步驟 3：一鍵動態過濾與 CRM 排重 (`Step 3: 分類過濾與 CRM 排重`)
    * **操作方式**：前往 **Step 3** 頁籤，依次上傳三個檔案：
      1. 原始 Raw Leads 檔 (`unverified_leads.csv`)
      2. Step 2 下載黏好 GRID 的 Apify 結果檔 (`apify_results_with_grid.csv`)
      3. 歷史 CRM 主檔案 (`bq-results-...csv`)
      4. 勾選/微調要攔截的 **WTG 類別清單**，點擊 **`▶ 開始分類比對`** 並下載最終 CSV。
    * **目的與原理**：
      * **第一層（硬關鍵字攔截）**：自動攔截店名帶有 59 個特定關鍵字（如：`有限公司`、`裝修`、`除毛`、`休閒農場`、`太陽堂`、`復興航棧`、`馥漫` 等非目標/特定品牌連鎖）。
      * **第二層（Google 類別動態攔截）**：比對 Google 官方類別，自動剔除 `超級市場`、`餅店`、`健康食品店`、`烘焙用具店`、`製造商`、`旅遊景點` 等非目標客群 (WTG)。
      * **第三層（歇業判讀）**：自動抓出 Google 上標示為永久歇業 (Permanently Closed) 的店家。
      * **第四層（CRM 店名+地址雙重排重）**：比對歷史 CRM 檔案，自動標記既有重複客戶 (P4 - Duplicate)，並自動帶出 Salesforce 庫中的 `GRID`。

    ---

    ### 🎯 最終產出的狀態標籤 (Final Classification)

    | 狀態標籤 | 意義與說明 | 後續處理方式 |
    | :--- | :--- | :--- |
    | **`P1 - New Lead`** | 驗證通過的合格餐飲新店家，無 CRM 重複紀錄。 | **直接發配給業務進行開發** |
    | **`Wrong Target Group (WTG)`** | 非目標店家（如：超級市場、餅店、健康食品店、生技、觀光工廠等）。 | 系統自動封存/排除 |
    | **`Permanently Closed`** | Google 地圖標示已歇業。 | 系統自動封存/排除 |
    | **`P4 - Duplicate`** | CRM 中已存在該店家（附上 Salesforce 的既有 GRID/ID）。 | 排除，防止業務撞單 |
    """)
