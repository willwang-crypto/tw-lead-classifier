# ── TAB 4: CRM CHECK (店名 + 地址雙重精準比對版) ───────────────────
    with tab4:
        st.subheader("🔍 Quick CRM Duplicate Check")
        st.caption("針對一般的餐廳名單進行 CRM 快速重複排查（需同時通過店名與地址比對，不同地址之分店一律歸為 Unverified）。")

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

                # 自動抓取店名欄位
                raw_name_col = find_column(df_raw, ["account name", "company", "account", "name", "title"])
                crm_name_col = find_column(df_crm, ["account name", "company", "account", "name", "title"])
                
                # 自動抓取地址欄位
                raw_addr_col = find_column(df_raw, ["street", "address", "地址", "formatted restaurant address"])
                crm_addr_col = find_column(df_crm, ["street", "address", "地址", "formatted restaurant address"])

                if not raw_name_col or not crm_name_col:
                    st.error("❌ 找不到餐廳名稱欄位，請檢查檔案標頭是否含有 Name, Account, Company 等字樣。")
                elif not raw_addr_col or not crm_addr_col:
                    st.error("❌ 找不到地址欄位，請檢查檔案標頭是否含有 Address, Street, 地址 等字樣。")
                else:
                    if st.button("▶ 開始 CRM 精準比對 (店名 + 地址)", type="primary", use_container_width=True):
                        with st.spinner("進行店名與地址雙重比對中..."):
                            # 預先處理 CRM 資料
                            df_crm['clean_name'] = df_crm[crm_name_col].fillna("").astype(str).str.strip().str.lower()
                            df_crm['clean_addr'] = df_crm[crm_addr_col].fillna("").astype(str).str.strip().str.lower()
                            
                            valid_crm_df = df_crm[(df_crm['clean_name'] != "")].copy()

                            def check_dup_with_address(row):
                                raw_name = str(row[raw_name_col]).lower().strip() if pd.notna(row[raw_name_col]) else ""
                                raw_addr = str(row[raw_addr_col]).lower().strip() if pd.notna(row[raw_addr_col]) else ""

                                if not raw_name:
                                    return "Unverified", "", 0.0

                                best_match_name = ""
                                best_score = 0.0
                                is_duplicate_address = False

                                for _, crm_row in valid_crm_df.iterrows():
                                    c_name = crm_row['clean_name']
                                    c_addr = crm_row['clean_addr']

                                    # 1. 計算店名相似度
                                    name_ratio = SequenceMatcher(None, raw_name, c_name).ratio()

                                    # 如果店名高度相似 (>= 70%)
                                    if name_ratio >= 0.70:
                                        # 2. 計算地址相似度
                                        addr_ratio = SequenceMatcher(None, raw_addr, c_addr).ratio() if (raw_addr and c_addr) else 0.0
                                        
                                        # 特殊比對：地址包含主要門牌數字與路名 (如 "吳興街279號" 在 "台北市信義區吳興街279號" 裡)
                                        addr_contains = (raw_addr in c_addr or c_addr in raw_addr) if len(raw_addr) > 5 and len(c_addr) > 5 else False

                                        # 如果【店名相同】且【地址也相同/高相似(>=65%)】，才認定為重複 (P4)
                                        if addr_ratio >= 0.65 or addr_contains:
                                            total_score = round(((name_ratio + max(addr_ratio, 0.8)) / 2) * 100, 1)
                                            return "P4 - Duplicate", c_name, total_score
                                        else:
                                            # 店名相同，但【地址不同】 -> 認定為不同分店，降級為 Unverified！
                                            continue

                                    if name_ratio > best_score:
                                        best_score = name_ratio
                                        best_match_name = c_name

                                # 如果只有店名像，但地址完全不一樣，直接判定為 Unverified (新分店)
                                return "Unverified", "", 0.0

                            results = df_raw.apply(check_dup_with_address, axis=1)
                            df_raw["CRM Status"] = [r[0] for r in results]
                            df_raw["Matched CRM Name"] = [r[1] for r in results]
                            df_raw["Similarity Score (%)"] = [r[2] for r in results]

                            st.write("### 比對結果預覽")
                            st.dataframe(df_raw.head(20))

                            csv_out = df_raw.to_csv(index=False).encode('utf-8-sig')
                            st.download_button(
                                label="📥 下載 CRM 排查結果 CSV",
                                data=csv_out,
                                file_name="crm_check_result_address.csv",
                                mime="text/csv"
                            )
            except Exception as e:
                st.error(f"處理檔案時發生錯誤: {str(e)}")
