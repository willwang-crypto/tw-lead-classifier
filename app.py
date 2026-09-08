# ── TAB 4: CRM CHECK (高精準度比對版) ──────────────────────────────
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
                    if st.button("▶ 開始 CRM 嚴格比對", type="primary", use_container_width=True):
                        with st.spinner("進行高精準度名稱相似度演算中，請稍候..."):
                            # 清洗並清洗 CRM 餐廳名稱
                            crm_series = df_crm[crm_name_col].fillna("").astype(str).str.strip().str.lower()
                            crm_names = list(set(crm_series[crm_series != ""].tolist()))

                            def check_dup_strict(raw_name):
                                if pd.isna(raw_name) or not str(raw_name).strip():
                                    return "Unverified", "", 0.0
                                
                                raw_str = str(raw_name).lower().strip()
                                
                                best_match = ""
                                best_score = 0.0

                                for c_name in crm_names:
                                    # 計算 SequenceMatcher 相似度比例 (0.0 ~ 1.0)
                                    ratio = SequenceMatcher(None, raw_str, c_name).ratio()
                                    
                                    # 加上字串長度懲罰：如果兩者長度相差太大（例如單字 "藏" vs "藏起來手作甜點"），強行壓低分數
                                    len_diff = abs(len(raw_str) - len(c_name))
                                    if len_diff > 4:
                                        ratio = ratio * 0.7
                                        
                                    if ratio > best_score:
                                        best_score = ratio
                                        best_match = c_name

                                best_score_pct = round(best_score * 100, 1)

                                # 設定嚴格判定門檻
                                if best_score_pct >= 85.0:
                                    return "P4 - Duplicate", best_match, best_score_pct
                                elif best_score_pct >= 65.0:
                                    return "P3 - Potential Match", best_match, best_score_pct
                                else:
                                    return "Unverified", "", 0.0

                            results = df_raw[raw_name_col].apply(check_dup_strict)
                            df_raw["CRM Status"] = [r[0] for r in results]
                            df_raw["Matched CRM Name"] = [r[1] for r in results]
                            df_raw["Similarity Score (%)"] = [r[2] for r in results]

                            st.write("### 比對結果預覽")
                            st.dataframe(df_raw.head(20))

                            csv_out = df_raw.to_csv(index=False).encode('utf-8-sig')
                            st.download_button(
                                label="📥 下載 CRM 排查結果 CSV",
                                data=csv_out,
                                file_name="crm_check_result_strict.csv",
                                mime="text/csv"
                            )
            except Exception as e:
                st.error(f"處理檔案時發生錯誤: {str(e)}")
