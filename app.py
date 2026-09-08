import streamlit as st
import pandas as pd
import re
from urllib.parse import quote

# -----------------------------------------------------------------------------
# 1. Helper Functions (支援台灣 3~5 位數郵編與彈性欄位比對)
# -----------------------------------------------------------------------------
def extract_tw_postal(address_or_code):
    """提取台灣 3 位數或 5 位數郵遞區號"""
    if pd.isna(address_or_code):
        return ""
    val = str(address_or_code).strip()
    # 搜尋獨立的 3 至 5 位數字
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
    """生成 Google Maps 搜尋與 Apify 爬蟲網址"""
    company_name = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else ""
    
    if url_format == "coords" and lat_col and lng_col and pd.notna(row[lat_col]) and pd.notna(row[lng_col]):
        lat = str(row[lat_col]).strip()
        lng = str(row[lng_col]).strip()
        # 使用經緯度 + 店名進行 Google Maps 搜尋網址拼接
        search_query = f"{company_name}@{lat},{lng}"
        return f"https://www.google.com/maps/search/{quote(search_query)}"
    else:
        street = str(row[street_col]).strip() if street_col and pd.notna(row[street_col]) else ""
        postal = extract_tw_postal(row[postal_col]) if postal_col and pd.notna(row[postal_col]) else ""
        full_address = f"{company_name} {street} {postal}".strip()
        return f"https://www.google.com/maps/search/{quote(full_address)}"

# -----------------------------------------------------------------------------
# 2. Streamlit UI (Tab 2 - Generate Apify URLs)
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Lead Classifier (Taiwan Version)", layout="wide")
st.title("🔗 Generate Google Maps URLs for Apify")
st.subheader("Step 1 · Generate URLs")

# 選擇模式
url_format_option = st.radio(
    "URL format",
    options=[
        "Company / Account + Coordinates (Latitude, Longitude)",
        "Company / Account + Address"
    ],
    index=0
)

url_format = "coords" if "Coordinates" in url_format_option else "address"

# 檔案上傳器
uploaded_file = st.file_uploader("Upload leads file (.xlsx or .csv)", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        # 讀取檔案
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.success(f"Successfully loaded {len(df)} rows.")

        # 自動偵測欄位
        name_col = find_column(df, ["company / account", "company", "account", "account name", "title"])
        street_col = find_column(df, ["street", "address"])
        postal_col = find_column(df, ["zip/postal code", "zip", "postal", "postal code"])
        lat_col = find_column(df, ["coordinates (lat", "latitude", "lat"])
        lng_col = find_column(df, ["coordinates (long", "longitude", "lng", "long"])
        grid_col = find_column(df, ["grid", "id"])

        # 檢驗必要欄位
        missing_fields = []
        if not name_col: missing_fields.append("Company / Account")
        if url_format == "coords" and (not lat_col or not lng_col):
            missing_fields.append("Coordinates (Latitude / Longitude)")
        elif url_format == "address" and not street_col:
            missing_fields.append("Street / Address")

        if missing_fields:
            st.error(f"❌ Missing required columns for this mode: {', '.join(missing_fields)}")
            st.warning("Please check your file header names.")
        else:
            # 產生網址
            df["url"] = df.apply(
                lambda row: generate_google_maps_url(row, url_format, name_col, street_col, postal_col, lat_col, lng_col),
                axis=1
            )

            # 選擇輸出欄位 (Apify 格式所需)
            output_cols = []
            if grid_col: output_cols.append(grid_col)
            if name_col: output_cols.append(name_col)
            output_cols.append("url")

            result_df = df[output_cols]

            st.write("### Preview Generated URLs")
            st.dataframe(result_df.head(10))

            # 轉換為 CSV 下載按鈕
            csv_data = result_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 Download URLs CSV for Apify",
                data=csv_data,
                file_name="apify_input_urls.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
