import streamlit as st
import pandas as pd
import numpy as np
import io

# 內建常見長照給付標準表（供無費用欄位的純人工班表查價使用）
DEFAULT_PRICE_TABLE = {
    "AA01": {"plain": 1700, "indigenous": 2040},
    "AA02": {"plain": 2000, "indigenous": 2400},
    "AA03": {"plain": 600,  "indigenous": 720},
    "AA04": {"plain": 600,  "indigenous": 720},
    "AA05": {"plain": 1200, "indigenous": 1440},
    "AA06": {"plain": 1200, "indigenous": 1440},
    "AA07": {"plain": 1200, "indigenous": 1440},
    "AA08": {"plain": 385,  "indigenous": 460},
    "AA09": {"plain": 500,  "indigenous": 500},
    "GA09": {"plain": 500,  "indigenous": 500},
    "BA01": {"plain": 260,  "indigenous": 310},
    "BA02": {"plain": 195,  "indigenous": 235},
    "BA03": {"plain": 35,   "indigenous": 40},
    "BA04": {"plain": 130,  "indigenous": 155},
    "BA05": {"plain": 310,  "indigenous": 370},
    "BA07": {"plain": 325,  "indigenous": 390},
    "BA08": {"plain": 500,  "indigenous": 600},
    "BA10": {"plain": 155,  "indigenous": 185},
    "BA11": {"plain": 195,  "indigenous": 235},
    "BA12": {"plain": 195,  "indigenous": 235},
    "BA13": {"plain": 195,  "indigenous": 235},
    "BA14": {"plain": 685,  "indigenous": 820},
    "BA15": {"plain": 195,  "indigenous": 235},
    "BA16": {"plain": 130,  "indigenous": 155},
    "BA17": {"plain": 150,  "indigenous": 180},
    "BA18": {"plain": 200,  "indigenous": 240},
    "BA20": {"plain": 175,  "indigenous": 210},
    "BA22": {"plain": 130,  "indigenous": 155},
    "BA23": {"plain": 200,  "indigenous": 240},
    "BA24": {"plain": 220,  "indigenous": 265},
}

def detect_header_row(df_raw):
    """自動尋找包含真正欄位名稱的那一列（解決系統匯出表頭有合併大標題的問題）"""
    keywords = ["個案", "案主", "姓名", "服務人員", "居服員", "照服員", "代碼", "項目", "碼別", "次數", "金額", "核銷", "身分證", "小計"]
    for idx in range(min(15, len(df_raw))):
        row_values = [str(val).strip() for val in df_raw.iloc[idx].values if pd.notna(val)]
        match_count = sum(1 for kw in keywords if any(kw in val for val in row_values))
        if match_count >= 2:
            df_clean = df_raw.iloc[idx + 1:].copy()
            df_clean.columns = [str(c).strip() if pd.notna(c) else f"欄位_{i}" for i, c in enumerate(df_raw.iloc[idx].values)]
            return df_clean.reset_index(drop=True)
    df_clean = df_raw.iloc[1:].copy()
    df_clean.columns = [str(c).strip() if pd.notna(c) else f"欄位_{i}" for i, c in enumerate(df_raw.iloc[0].values)]
    return df_clean.reset_index(drop=True)

def read_uploaded_files(uploaded_files):
    """支援單檔或多檔自動合併讀取，並自動定位真實標題列"""
    dfs = []
    for f in uploaded_files:
        f.seek(0)
        if f.name.endswith('.csv'):
            df_raw = pd.read_csv(f, header=None)
        else:
            df_raw = pd.read_excel(f, header=None)
        df_clean = detect_header_row(df_raw)
        dfs.append(df_clean)
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def auto_match_col(columns, keywords):
    """智慧自動匹配欄位名稱"""
    for kw in keywords:
        for col in columns:
            if kw in str(col):
                return col
    return columns[0] if len(columns) > 0 else None

def render_home_care_payroll():
    st.title("💰 居家服務薪資拆帳與核銷對帳引擎")
    st.caption("支援【雙模自動計薪】：可自動讀取 Jubo 費用清冊（強制抓取平地費用算 B 碼），亦支援純代碼班表查價計算，並自動拆帳多人共照 A 碼。")

    # ==========================================
    # 1. 動態參數引擎設定區
    # ==========================================
    with st.expander("⚙️ 機構拆帳內規設定（已預設：A碼55分/含原民、B碼64分/全平地）", expanded=True):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**1️⃣ 第一層：公司 vs 居服員分潤**")
            a_share_pct = st.slider("A碼 居服員分潤比例 (%)", 0, 100, 50, step=5, help="55拆帳請設 50%")
            b_share_pct = st.slider("B碼 居服員分潤比例 (%)", 0, 100, 60, step=5, help="64拆帳請設 60%")
            
        with col2:
            st.markdown("**2️⃣ 原民區 / 平地計價規則**")
            a_use_indigenous = st.checkbox("A碼 依核銷/區域套用原民加成", value=True)
            b_use_indigenous = st.checkbox("B碼 套用原民加成（未勾選則全採平地價）", value=False)
            split_method = st.selectbox(
                "多人共照 A碼 分配依據",
                ["依服務次數比例", "依人頭均分"],
                index=0
            )
            
        with col3:
            st.markdown("**3️⃣ 特定代碼公定單價覆寫**")
            custom_aa09 = st.number_input("AA09 公定單價 (元/天)", value=500, step=50)
            custom_ga09 = st.number_input("GA09 公定單價 (元/天)", value=500, step=50)
            round_to_int = st.checkbox("金額四捨五入至整數", value=True)

    price_map = DEFAULT_PRICE_TABLE.copy()
    price_map["AA09"] = {"plain": custom_aa09, "indigenous": custom_aa09}
    price_map["GA09"] = {"plain": custom_ga09, "indigenous": custom_ga09}

    st.divider()

    # ==========================================
    # 2. 檔案上傳區（支援多檔合併）
    # ==========================================
    u_col1, u_col2 = st.columns(2)
    with u_col1:
        service_files = st.file_uploader(
            "📂 上傳【服務清冊 / 班表】(必填，可多選合併)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True
        )
    with u_col2:
        a_claim_files = st.file_uploader(
            "📂 上傳【A碼核銷表】(選填，若無則直接依服務清冊A碼計算)",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True
        )

    if not service_files:
        st.info("👆 請先上傳 Jubo 或照管系統匯出的「服務清冊」Excel 檔以開始計算。")
        return

    df_service = read_uploaded_files(service_files)
    df_a_claim = read_uploaded_files(a_claim_files) if a_claim_files else pd.DataFrame()

    # ==========================================
    # 3. 智慧欄位對應（雙模自動偵測）
    # ==========================================
    st.subheader("🔗 欄位對齊確認")
    s_cols = list(df_service.columns)
    m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns(5)
    
    with m_col1:
        col_client = st.selectbox("個案姓名/ID", s_cols, index=s_cols.index(auto_match_col(s_cols, ["個案姓名", "個案", "案主", "姓名"])))
    with m_col2:
        col_worker = st.selectbox("服務人員姓名", s_cols, index=s_cols.index(auto_match_col(s_cols, ["服務人員", "居服員", "照服員", "員工"])))
    with m_col3:
        col_code = st.selectbox("服務代碼 / 項目類別", s_cols, index=s_cols.index(auto_match_col(s_cols, ["服務代碼", "項目代碼", "代碼", "服務項目類別", "項目"])))
    with m_col4:
        col_count = st.selectbox("服務次數/數量", s_cols, index=s_cols.index(auto_match_col(s_cols, ["次數", "數量", "總次", "天數"])))
    with m_col5:
        fee_options = ["(無，依代碼查表計算)"] + s_cols
        # 優先匹配平地金額「申請(補助)費用」，避開「原民區或離島申請(補助)費用」
        exact_plain_fee = next((c for c in s_cols if str(c).strip() == "申請(補助)費用"), None)
        default_fee = exact_plain_fee if exact_plain_fee else auto_match_col(s_cols, ["申請(補助)費用", "補助費用", "申請費用", "小計", "實際補助"])
        default_fee_idx = fee_options.index(default_fee) if default_fee in s_cols and any(k in str(default_fee) for k in ["費用", "金額", "小計", "補助"]) else 0
        col_direct_fee = st.selectbox("B碼計薪基準欄位 (平地費用)", fee_options, index=default_fee_idx, help="Jubo清冊請選「申請(補助)費用」即為純平地價；無此欄則自動依代碼查表")

    region_options = ["(無此欄位，全算平地)"] + s_cols
    default_reg = auto_match_col(s_cols, ["原民", "行政區", "居住行政區", "區域", "地區"])
    default_reg_idx = region_options.index(default_reg) if default_reg in s_cols and any(k in str(default_reg) for k in ["原民", "行政區", "區域"]) else 0
    col_region = st.selectbox("原民/平地判斷欄位 (供無費用欄位查表時使用)", region_options, index=default_reg_idx)

    if not df_a_claim.empty:
        st.markdown("**A碼核銷表欄位對應：**")
        ac_cols = list(df_a_claim.columns)
        ac1, ac2, ac3 = st.columns(3)
        with ac1:
            ac_client = st.selectbox("核銷表-個案姓名/ID", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["個案姓名", "個案", "案主", "姓名"])))
        with ac2:
            ac_code = st.selectbox("核銷表-A碼項目", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["服務代碼", "代碼", "項目", "碼別"])))
        with ac3:
            ac_amount = st.selectbox("核銷表-核銷總金額", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["小計", "核銷金額", "申報金額", "總金額", "金額", "費用"])))

    # ==========================================
    # 4. 核心計算邏輯執行
    # ==========================================
    if st.button("🚀 開始自動拆帳與對帳計算", type="primary", use_container_width=True):
        work_df = df_service.copy()
        # 過濾掉系統報表最底部的「總計 / 合計」列
        work_df = work_df[work_df[col_client].notna() & work_df[col_worker].notna()].copy()
        work_df = work_df[~work_df[col_client].astype(str).str.contains("總計|合計|小計", na=False)].copy()

        work_df[col_code] = work_df[col_code].astype(str).str.strip().str.upper()
        work_df[col_count] = pd.to_numeric(work_df[col_count], errors='coerce').fillna(1)
        work_df[col_client] = work_df[col_client].astype(str).str.strip()
        work_df[col_worker] = work_df[col_worker].astype(str).str.strip()

        if col_region != "(無此欄位，全算平地)":
            work_df["is_indigenous"] = work_df[col_region].astype(str).str.contains("原|山地|離島|牡丹|獅子|春日|泰武|來義|瑪家|三地門|霧臺|滿州", na=False)
        else:
            work_df["is_indigenous"] = False

        # ------------------------------------------
        # Part A：雙模計算 B 碼薪資
        # ------------------------------------------
        if col_direct_fee != "(無，依代碼查表計算)":
            # 【模式一】清冊已含平地費用欄位：排除 A 碼後，直接取該欄金額 × B碼分潤比例
            b_df = work_df[~work_df[col_code].str.startswith(("AA", "GA", "A碼", "個案管理"))].copy()
            b_df["核銷總額(平地基準)"] = pd.to_numeric(b_df[col_direct_fee], errors='coerce').fillna(0)
            b_df["居服員實領"] = b_df["核銷總額(平地基準)"] * (b_share_pct / 100.0)
        else:
            # 【模式二】純代碼班表：依代碼查單價 × 次數 × B碼分潤比例
            b_df = work_df[work_df[col_code].str.startswith("B")].copy()
            def calc_b_row(row):
                code = row[col_code]
                price_info = price_map.get(code, {"plain": 0, "indigenous": 0})
                unit_price = price_info["indigenous"] if (b_use_indigenous and row["is_indigenous"]) else price_info["plain"]
                gross = unit_price * row[col_count]
                worker_pay = gross * (b_share_pct / 100.0)
                return pd.Series([unit_price, gross, worker_pay])

            if not b_df.empty:
                b_df[["適用單價", "核銷總額(平地基準)", "居服員實領"]] = b_df.apply(calc_b_row, axis=1)

        if not b_df.empty:
            if round_to_int:
                b_df["居服員實領"] = b_df["居服員實領"].round(0)
            b_df["公司留存"] = b_df["核銷總額(平地基準)"] - b_df["居服員實領"]

        # ------------------------------------------
        # Part B：計算 A 碼拆帳 (多人共照依服務次數比例切分)
        # ------------------------------------------
        worker_client_counts = (
            work_df.groupby([col_client, col_worker])[col_count]
            .sum()
            .reset_index()
            .rename(columns={col_count: "個人服務次數"})
        )
        client_total_counts = (
            worker_client_counts.groupby(col_client)["個人服務次數"]
            .sum()
            .reset_index()
            .rename(columns={"個人服務次數": "該案總服務次數"})
        )
        ratio_df = pd.merge(worker_client_counts, client_total_counts, on=col_client)
        ratio_df["分配比例"] = ratio_df["個人服務次數"] / ratio_df["該案總服務次數"]

        if not df_a_claim.empty:
            a_base = df_a_claim[[ac_client, ac_code, ac_amount]].copy()
            a_base.columns = [col_client, col_code, "核銷總額"]
            a_base = a_base[a_base[col_client].notna()].copy()
            a_base = a_base[~a_base[col_client].astype(str).str.contains("總計|合計|小計", na=False)].copy()
            a_base[col_client] = a_base[col_client].astype(str).str.strip()
            a_base[col_code] = a_base[col_code].astype(str).str.strip().str.upper()
            a_base["核銷總額"] = pd.to_numeric(a_base["核銷總額"], errors='coerce').fillna(0)
        else:
            a_raw = work_df[work_df[col_code].str.startswith(("AA", "GA"))].copy()
            def calc_a_gross(row):
                code = row[col_code]
                price_info = price_map.get(code, {"plain": 0, "indigenous": 0})
                unit_price = price_info["indigenous"] if (a_use_indigenous and row["is_indigenous"]) else price_info["plain"]
                return unit_price * row[col_count]
            a_raw["核銷總額"] = a_raw.apply(calc_a_gross, axis=1)
            a_base = a_raw.groupby([col_client, col_code])["核銷總額"].sum().reset_index()

        a_split_df = pd.merge(a_base, ratio_df, on=col_client, how="left")
        a_split_df[col_worker] = a_split_df[col_worker].fillna("⚠️ 無班表紀錄(需手動指定)")
        a_split_df["分配比例"] = a_split_df["分配比例"].fillna(1.0)
        a_split_df["居服員分潤池"] = a_split_df["核銷總額"] * (a_share_pct / 100.0)
        
        if split_method == "依人頭均分":
            headcount = a_split_df.groupby([col_client, col_code])[col_worker].transform("count")
            a_split_df["分配比例"] = 1.0 / headcount

        a_split_df["居服員實領"] = a_split_df["居服員分潤池"] * a_split_df["分配比例"]
        if round_to_int:
            a_split_df["居服員實領"] = a_split_df["居服員實領"].round(0)
        a_split_df["公司留存"] = (a_split_df["核銷總額"] * a_split_df["分配比例"]) - a_split_df["居服員實領"]

        # ------------------------------------------
        # Part C：彙總每位居服員薪資總表
        # ------------------------------------------
        b_summary = b_df.groupby(col_worker)[["核銷總額(平地基準)", "居服員實領"]].sum().reset_index().rename(
            columns={"核銷總額(平地基準)": "B碼計薪總額(平地)", "居服員實領": "B碼實領薪資"}
        ) if not b_df.empty else pd.DataFrame(columns=[col_worker, "B碼計薪總額(平地)", "B碼實領薪資"])

        a_summary = a_split_df.groupby(col_worker)[["核銷總額", "居服員實領"]].sum().reset_index().rename(
            columns={"核銷總額": "A碼核銷總額", "居服員實領": "A碼實領薪資"}
        ) if not a_split_df.empty else pd.DataFrame(columns=[col_worker, "A碼核銷總額", "A碼實領薪資"])

        final_payroll = pd.merge(b_summary, a_summary, on=col_worker, how="outer").fillna(0)
        final_payroll["本月應發總薪資(A+B)"] = final_payroll["B碼實領薪資"] + final_payroll["A碼實領薪資"]

        # ==========================================
        # 5. 畫面呈現與 Excel 報表下載
        # ==========================================
        st.success("✅ 計算完成！已自動套用【B碼純平地64分 ＋ A碼多人次數比例55分】規則：")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("總計應發居服員薪資", f"${final_payroll['本月應發總薪資(A+B)'].sum():,.0f}")
        m2.metric("A碼居服員實領總額", f"${final_payroll['A碼實領薪資'].sum():,.0f}")
        m3.metric("B碼居服員實領總額", f"${final_payroll['B碼實領薪資'].sum():,.0f}")

        tab1, tab2, tab3 = st.tabs(["📊 居服員薪資總表", "🔍 A碼多人拆帳明細", "📋 B碼計算明細"])
        with tab1:
            st.dataframe(final_payroll, use_container_width=True)
        with tab2:
            st.dataframe(a_split_df, use_container_width=True)
        with tab3:
            st.dataframe(b_df, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_payroll.to_excel(writer, sheet_name="1_居服員薪資總表", index=False)
            a_split_df.to_excel(writer, sheet_name="2_A碼多人拆帳明細", index=False)
            if not b_df.empty:
                b_df.to_excel(writer, sheet_name="3_B碼服務計薪明細", index=False)
        
        st.download_button(
            label="📥 下載完整對帳 Excel 報表（含薪資總表＋A/B碼明細）",
            data=output.getvalue(),
            file_name="居家服務薪資拆帳與核銷對帳表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )