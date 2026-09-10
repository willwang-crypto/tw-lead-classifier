import pandas as pd
import streamlit as st
from difflib import SequenceMatcher
from urllib.parse import quote, unquote
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

# 店名/公司硬攔截關鍵字
NON_FOOD_NAME_KEYWORDS = [
    "貿易", "生技", "股份有限公司", "有限公司", "博物館", "museum", 
    "工場", "觀光工廠", "批發", "設備", "生鮮專賣", "器材", "實業"
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

# ── TAB 1: CLASSIFY LEADS (動態全 Google 類別選擇器版) ──────────────
with tab1:
    st.subheader("📊 Classify Leads (全 Google 地圖類別動態分類引擎)")
    st.caption("系統會自動抓取檔案中所有的 Google 類別，您可以自由勾選/剔除哪些類別屬於非目標 (WTG)。")

    c1, c2, c3 = st.columns(3)
    with c1:
        leads_up = st.file_uploader("1. 上傳 Unverified Leads 檔", type=["xlsx","xls","csv"], key="t1_leads")
    with c2:
        apify_up = st.file_uploader("2. 上傳 Apify 結果檔", type=["xlsx","xls","csv"], key="t1_apify")
    with c3:
        crm_up = st.file_uploader("3. 上傳 CRM All Accounts 大檔", type=["xlsx","xls","csv"], key="t1_crm")

    if leads_up and apify_up and crm_up:
        df_leads = safe_load_csv(leads_up)
        df_apify = safe_load_csv(apify_up)
        df_crm = safe_load_csv(crm_up)

        # 抓取 Apify 類別欄位
        a_cat_col = find_column(df_apify, ["categoryname", "categories/0", "primarycategory", "category"])
        
        # 抓出 Apify 檔案中出現過的所有 Google 類別
        all_detected_cats = []
        if a_cat_col and a_cat_col in df_apify.columns:
            all_detected_cats = sorted([str(c).strip() for c in df_apify[a_cat_col].dropna().unique() if str(c).strip()])

        # 自動生成非餐飲預設選擇清單
        wtg_default_keywords = [
            "餅店", "糕餅", "製造商", "批發", "商店", "景點", "傳統市場", "專賣", "公司", "博物館", 
            "俱樂部", "服務", "供應商", "維修", "音響", "設備", "書", "展覽", "藝廊", "營地", 
            "補習班", "肉檔", "肉鋪", "百貨", "內衣", "冷凍食品", "生鮮", "有機", "堅果", "健康食品", "乳酪雪糕"
        ]
        
        default_selected_wtg = [
            cat for cat in all_detected_cats 
            if any(kw in cat for kw in wtg_default_keywords) and not any(f in cat for f in ["餐廳", "小吃", "熟食", "麵店", "咖啡", "飲料", "火鍋", "早午餐"])
        ]

        st.markdown("---")
        st.write("### ⚙️ Google 地圖類別過濾設定 (WTG 過濾清單)")
        selected_wtg_cats = st.multiselect(
            f"🔍 系統自上傳檔案中掃描到 {len(all_detected_cats)} 種 Google 類別。被選中的類別將被判定為 『Wrong Target Group (WTG)』：",
            options=all_detected_cats,
            default=default_selected_wtg
        )

        if st.button("▶ 開始分類比對 (Run Classification)", type="primary", use_container_width=True):
            with st.spinner("執行全類別動態判定與 CRM 比對中..."):
                try:
                    l_name_col = find_column(df_leads, ["company / account", "company", "account", "name", "title"])
                    l_addr_col = find_column(df_leads, ["street", "address", "地址"])

                    a_name_col = find_column(df_apify, ["title", "name", "searchstring"])
                    a_search_col = find_column(df_apify, ["searchstring", "inputstarturl", "url"])
                    a_status_col = find_column(df_apify, ["permanentlyclosed", "temporarilyclosed", "isclosed", "status"])

                    c_name_col = find_column(df_crm, ["account name", "company", "account", "name", "title"])
                    c_addr_col = find_column(df_crm, ["street", "address", "地址"])

                    crm_pairs = []
                    if c_name_col and c_addr_col:
                        c_names = [clean_text(x) for x in df_crm[c_name_col].fillna("")]
                        c_addrs = [clean_text(x) for x in df_crm[c_addr_col].fillna("")]
                        crm_pairs = list(zip(c_names, c_addrs))

                    apify_records = []
                    for _, a_row in df_apify.iterrows():
                        a_title = clean_text(a_row[a_name_col]) if a_name_col and pd.notna(a_row[a_name_col]) else ""
                        a_search = unquote(str(a_row[a_search_col])).lower() if a_search_col and pd.notna(a_row[a_search_col]) else ""
                        cat = str(a_row[a_cat_col]).strip() if a_cat_col and pd.notna(a_row[a_cat_col]) else ""
                        is_closed = str(a_row[a_status_col]) if a_status_col and pd.notna(a_row[a_status_col]) else "False"
                        
                        apify_records.append({
                            "title": a_title,
                            "search": a_search,
                            "cat": cat,
                            "is_closed": is_closed
                        })

                    final_statuses = []
                    matched_crm_names = []
                    google_cats = []
                    debug_reasons = []

                    for idx, l_row in df_leads.iterrows():
                        l_name_raw = str(l_row[l_name_col]).strip() if l_name_col and pd.notna(l_row[l_name_col]) else ""
                        l_addr_raw = str(l_row[l_addr_col]).strip() if l_addr_col and pd.notna(l_row[l_addr_col]) else ""

                        l_name = clean_text(l_name_raw)
                        l_addr = clean_text(l_addr_raw)

                        # 1. 第一層：店名本體硬過濾 (B2B/生技/貿易/工場)
                        if any(nk in l_name for nk in NON_FOOD_NAME_KEYWORDS):
                            final_statuses.append("Wrong Target Group (WTG)")
                            matched_crm_names.append("")
                            google_cats.append("由店名攔截")
                            debug_reasons.append("店名包含非餐飲/B2B關鍵字")
                            continue

                        # 2. 搜尋 Apify 結果
                        best_apify_match = None
                        best_score = 0.0

                        for a_rec in apify_records:
                            if l_name and l_name in a_rec["search"]:
                                best_apify_match = a_rec
                                break
                            if l_name and a_rec["title"]:
                                score = SequenceMatcher(None, l_name, a_rec["title"]).ratio()
                                if score > best_score and score >= 0.50:
                                    best_score = score
                                    best_apify_match = a_rec

                        cat_str = best_apify_match["cat"] if best_apify_match else ""
                        is_closed_str = str(best_apify_match["is_closed"]).lower() if best_apify_match else "false"

                        google_cats.append(cat_str if cat_str else "未對應到")

                        # 3. 判定歇業
                        if "true" in is_closed_str or "closed" in is_closed_str or "1" == is_closed_str:
                            final_statuses.append("Permanently Closed")
                            matched_crm_names.append("")
                            debug_reasons.append("Google 地圖標示歇業")
                            continue

                        # 4. 第二層：自訂 Google 類別動態攔截
                        if cat_str in selected_wtg_cats:
                            final_statuses.append("Wrong Target Group (WTG)")
                            matched_crm_names.append("")
                            debug_reasons.append(f"Google 類別屬於選定的 WTG [{cat_str}]")
                            continue

                        # 5. CRM 重複比對
                        best_status = "P1 - New Lead"
                        best_crm_name = ""
                        debug_msg = f"驗證通過 (Google類別: {cat_str or '未對應到，店名正常'})"

                        if l_name:
                            prefix = l_name[:2]
                            for c_name, c_addr in crm_pairs:
                                if not (c_name.startswith(prefix) or prefix in c_name):
                                    continue
                                
                                name_score = SequenceMatcher(None, l_name, c_name).ratio()
                                if name_score >= 0.60:
                                    addr_score = SequenceMatcher(None, l_addr, c_addr).ratio() if (l_addr and c_addr) else 0.0
                                    addr_contains = (l_addr in c_addr or c_addr in l_addr) if len(l_addr) > 4 and len(c_addr) > 4 else False

                                    if addr_score >= 0.50 or addr_contains:
                                        best_status = "P4 - Duplicate"
                                        best_crm_name = c_name
                                        debug_msg = f"命中 CRM 重複檔 ({c_name})"
                                        break

                        final_statuses.append(best_status)
                        matched_crm_names.append(best_crm_name)
                        debug_reasons.append(debug_msg)

                    df_leads["Final Classification"] = final_statuses
                    df_leads["Google Category"] = google_cats
                    df_leads["Matched CRM Name"] = matched_crm_names
                    df_leads["判定依據說明"] = debug_reasons

                    counts = df_leads["Final Classification"].value_counts().to_dict()
                    st.write("### 📊 最終分類統計結果：")
                    st.write(counts)

                    st.dataframe(df_leads[["Final Classification", l_name_col, "Google Category", "Matched CRM Name", "判定依據說明"]].head(30))

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
        ["📝 Company / Account + Address"],
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
            st.error(f"錯誤: {str(e)}")

# ── TAB 4: CRM CHECK ─────────────────────────────────────────────
with tab4:
    st.subheader("🔍 Quick CRM Duplicate Check")
    st.caption("自動排除純零售超商 + 店名地址雙重比對。")

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
                if st.button("▶ 開始 CRM 比對 (純地址版)", type="primary"):
                    with st.spinner("比對中..."):
                        crm_names = [clean_text(x) for x in df_crm[crm_name_col].fillna("")]
                        crm_addrs = [clean_text(x) for x in df_crm[crm_addr_col].fillna("")]
                        crm_pairs = list(zip(crm_names, crm_addrs))

                        def check_single(row):
                            r_name_raw = str(row[raw_name_col]) if pd.notna(row[raw_name_col]) else ""
                            r_addr_raw = str(row[raw_addr_col]) if pd.notna(row[raw_addr_col]) else ""
                            
                            r_name = clean_text(r_name_raw)
                            r_addr = clean_text(r_addr_raw)

                            if not r_name:
                                return "Unverified", "", 0.0

                            prefix = r_name[:2]
                            for c_name, c_addr in crm_pairs:
                                if not (c_name.startswith(prefix) or prefix in c_name):
                                    continue
                                
                                name_score = SequenceMatcher(None, r_name, c_name).ratio()
                                if name_score >= 0.60:
                                    addr_score = SequenceMatcher(None, r_addr, c_addr).ratio() if (r_addr and c_addr) else 0.0
                                    addr_contains = (r_addr in c_addr or c_addr in r_addr) if len(r_addr) > 4 and len(c_addr) > 4 else False

                                    if addr_score >= 0.50 or addr_contains:
                                        return "P4 - Duplicate", c_name, round(((name_score + max(addr_score, 0.8))/2)*100, 1)

                            return "Unverified", "", 0.0

                        results = df_raw.apply(check_single, axis=1)
                        df_raw["CRM Status"] = [r[0] for r in results]
                        df_raw["Matched CRM Name"] = [r[1] for r in results]
                        df_raw["Similarity Score (%)"] = [r[2] for r in results]

                        p4_count = (df_raw["CRM Status"] == "P4 - Duplicate").sum()
                        unv_count = (df_raw["CRM Status"] == "Unverified").sum()

                        st.write(f"### 📊 排查結果摘要：Unverified 新餐飲店 ({unv_count} 筆) | P4 重複 ({p4_count} 筆)")
                        st.dataframe(df_raw.head(20))

                        csv_out = df_raw.to_csv(index=False).encode('utf-8-sig')
                        st.download_button("📥 下載 CRM 排查結果 CSV", data=csv_out, file_name="crm_result_filtered.csv", mime="text/csv")
        except Exception as e:
            st.error(f"錯誤: {str(e)}")
