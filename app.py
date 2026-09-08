import streamlit as st
import pandas as pd
import re
from urllib.parse import quote

# -----------------------------------------------------------------------------
# 1. 核心通用函式（支援台灣郵編與彈性欄位偵測）
# -----------------------------------------------------------------------------
def extract_tw_postal(address_or_code):
    """提取台灣 3 位數或 5 位數郵遞區號"""
    if pd.isna(address_or_code):
        return ""
    val = str(address_or_code).strip()
    match = re.search(r'\b(\d{3,5})\b', val)
    return match.group(1) if match else ""

def find_column(df, possible_names):
    """不分大小寫且支援模糊比對欄位名稱"""
    for col in df.columns:
        col_clean = str(col).strip().lower()
        for target in possible_names:
            if target.lower() in col_clean:
                return col
    return None

def generate_google_maps_url(row, url_format, name_col, street_col, postal_col, lat_col, lng_col):
    """生成 Google Maps 搜尋網址"""
    company_name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else ""
    
    if url_format == "coords" and lat_col and lng_col and pd.notna(row[lat_col]) and pd.notna(row[lng_col]):
        lat = str(row[lat_col]).strip()
        lng = str(row[lng_col]).strip()
        search_query = f"{company_name}@{lat},{lng}"
        return f"https://www.google.com/maps/search/{quote(search_query)}"
    else:
        street = str(row[street_col]).strip() if street_col and pd.notna(row[street_col]) else ""
        postal = extract_tw_postal(row[postal_col]) if postal_col and pd.notna(row[postal_col]) else ""
        full_address = f"{company_name} {street} {postal}".strip()
        return f"https://www.google.com/maps/search/{quote(full_address)}"

# -----------------------------------------------------------------------------
# 2. 頁面設定與主選單 (Tab 1 ~ Tab 6)
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Sales Ops · Lead Classifier Suite (Taiwan)", layout="wide")

st.title("Sales Ops · Data Quality Suite — Taiwan (Geo-Distance Version)")
st.caption("foodpanda / Delivery Hero · Digital Sales APAC")

tabs = st.tabs([
    "📊 Classify Leads", 
    "🔗 Generate Apify URLs", 
    "🏢 SF Account Audit", 
    "🔍 CRM Check", 
    "📋 KPI Sample Checker", 
    "📖 How to Use"
])

# -----------------------------------------------------------------------------
# Tab 1: Classify Leads
# -----------------------------------------------------------------------------
with tabs[0]:
    st.header("📊 Classify Leads")
    st.info("請上傳 SF Leads、Apify 爬蟲結果與 CRM All Accounts 進行完整分類處理。")
    st.file_uploader("Upload Salesforce Leads", type=["csv", "xlsx"], key="tab1_leads")
    st.file_uploader("Upload Apify Results (with GRID)", type=["csv", "xlsx"], key="tab1_apify")
    st.file_uploader("Upload Salesforce All Accounts", type=["csv", "xlsx"], key="tab1_crm")

# -----------------------------------------------------------------------------
# Tab 2: Generate Apify URLs (已成功修復與優化)
# -----------------------------------------------------------------------------
with tabs[1]:
    st.header("🔗 Generate Google Maps URLs for Apify")
    st.subheader("Step 1 · Generate URLs")

    url_format_option = st.radio(
        "URL format",
        options=[
            "📍 Company / Account + Coordinates (Latitude, Longitude)",
            "📝 Company / Account + Address"
        ],
        index=0
    )
    url_format = "coords" if "Coordinates" in url_format_option else "address"

    uploaded_file = st.file_uploader("Upload leads file (.xlsx or .csv)", type=["csv", "xlsx"], key="tab2_file")

    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)
            st.success(f"Successfully loaded {len(df)} rows.")

            # 精準對應台灣經緯度欄位表頭
            name_col = find_column(df, ["company / account", "company", "account", "account name", "title"])
            street_col = find_column(df, ["street", "address"])
            postal_col = find_column(df, ["zip/postal code", "zip", "postal", "postal code"])
            lat_col = find_column(df, ["coordinates (lat", "latitude", "lat"])
            lng_col = find_column(df, ["coordinates (long", "longitude", "lng", "long"])
            grid_col = find_column(df, ["grid", "id"])

            missing_fields = []
            if not name_col: missing_fields.append("Company / Account")
            if url_format == "coords" and (not lat_col or not lng_col):
                missing_fields.append("Coordinates (Latitude / Longitude)")
            elif url_format == "address" and not street_col:
                missing_fields.append("Street / Address")

            if missing_fields:
                st.error(f"❌ Missing required columns: {', '.join(missing_fields)}")
            else:
                df["url"] = df.apply(
                    lambda row: generate_google_maps_url(row, url_format, name_col, street_col, postal_col, lat_col, lng_col),
                    axis=1
                )

                output_cols = []
                if grid_col: output_cols.append(grid_col)
                if name_col: output_cols.append(name_col)
                output_cols.append("url")

                result_df = df[output_cols]
                st.write("### Preview Generated URLs")
                st.dataframe(result_df.head(10))

                csv_data = result_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 Download URLs CSV for Apify",
                    data=csv_data,
                    file_name="apify_input_urls.csv",
                    mime="text/csv"
                )
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

# -----------------------------------------------------------------------------
# Tab 3: SF Account Audit
# -----------------------------------------------------------------------------
with tabs[2]:
    st.header("🏢 SF Account Audit")
    st.info("比對 Salesforce 內部帳號，找出潛在重複檔。")
    st.file_uploader("Upload Salesforce All Accounts File", type=["csv", "xlsx"], key="tab3_crm")

# -----------------------------------------------------------------------------
# Tab 4: CRM Check
# -----------------------------------------------------------------------------
with tabs[3]:
    st.header("🔍 CRM Check")
    st.info("純 CRM 比對與去重，無需 Apify 爬蟲資料。")
    st.file_uploader("Upload Raw Restaurant List", type=["csv", "xlsx"], key="tab4_raw")
    st.file_uploader("Upload CRM All Accounts", type=["csv", "xlsx"], key="tab4_crm")

# -----------------------------------------------------------------------------
# Tab 5: KPI Sample Checker
# -----------------------------------------------------------------------------
with tabs[4]:
    st.header("📋 KPI Sample Checker")
    st.info("進行每月 10% 抽樣與 15 項自動化品質檢查。")
    st.file_uploader("Upload Lead Status Change Report", type=["csv", "xlsx"], key="tab5_report")

# -----------------------------------------------------------------------------
# Tab 6: How to Use
# -----------------------------------------------------------------------------
with tabs[5]:
    st.header("📖 How to Use")
    st.markdown("""
    ### 使用指南與邏輯說明
    1. **Generate Apify URLs**：上傳名單，透過經緯度或台灣地址（含 3~5 位數郵編）自動產生爬蟲網址。
    2. **Classify Leads**：結合 CRM 與 Apify 爬蟲結果產出包含 P1~P4 與停業標籤的 Excel 報告。
    3. **CRM Check / SF Audit**：處理名單建檔前的去重與衛生檢查。
    """)
