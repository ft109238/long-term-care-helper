import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    """自動尋找包含真正欄位名稱的那一列"""
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

def build_formatted_workbook(final_payroll, a_split_df, b_df, col_client, col_worker, col_code, col_count, month_label, a_share_pct, b_share_pct, target_workers=None):
    """產生精美格式 Excel：target_workers=None 時匯出全機構總表+全員分頁；指定單人時匯出該員獨立薪資單"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wb = writer.book
        # 若為全機構匯出，第一頁放總表；若為單人匯出，直接只放該員個人薪資明細表
        if target_workers is None:
            final_payroll.to_excel(writer, sheet_name="0_全機構薪資總表", index=False)
            workers = [w for w in final_payroll[col_worker].dropna().unique() if not str(w).startswith("⚠️")]
        else:
            workers = target_workers

        font_title = Font(name="微軟正黑體", size=14, bold=True)
        font_sec = Font(name="微軟正黑體", size=11, bold=True, color="1F4E78")
        font_header = Font(name="微軟正黑體", size=10, bold=True)
        font_body = Font(name="微軟正黑體", size=10)
        font_total = Font(name="微軟正黑體", size=10, bold=True)
        font_grand = Font(name="微軟正黑體", size=11, bold=True, color="C00000")

        fill_header_a = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        fill_header_b = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        fill_total = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        fill_grand = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        thin_side = Side(border_style="thin", color="808080")
        border_box = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        align_center = Alignment(horizontal="center", vertical="center")
        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        for worker in workers:
            safe_sheet_name = str(worker)[:30].replace("/", "_").replace("\\", "_").replace("*", "").replace("?", "").replace("[", "").replace("]", "")
            ws = wb.create_sheet(title=safe_sheet_name)

            # 設定 A4 列印版面（橫向/縱向自動縮放至 1 頁寬，方便行政直接列印）
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True

            ws.merge_cells("A1:G1")
            ws["A1"] = f"{month_label} {worker} 薪資服務明細"
            ws["A1"].font = font_title
            ws["A1"].alignment = align_center
            ws.row_dimensions[1].height = 28

            # 上半部：A 碼薪資明細表
            ws["A3"] = "【一、A碼個案管理費拆帳明細】"
            ws["A3"].font = font_sec

            a_headers = [
                "個案姓名",
                "個案總服務次數",
                "服務員服務次數",
                "A碼服務項目",
                "A碼金額",
                "個人分攤A碼金額",
                f"A碼拆帳{a_share_pct}%"
            ]
            row_idx = 4
            for c_i, h_text in enumerate(a_headers, start=1):
                cell = ws.cell(row=row_idx, column=c_i, value=h_text)
                cell.font = font_header
                cell.fill = fill_header_a
                cell.alignment = align_center
                cell.border = border_box
            ws.row_dimensions[row_idx].height = 22

            w_a_df = a_split_df[a_split_df[col_worker] == worker] if not a_split_df.empty else pd.DataFrame()
            a_sum_pay = 0.0

            if not w_a_df.empty:
                for _, r in w_a_df.iterrows():
                    row_idx += 1
                    personal_gross = float(r["核銷總額"]) * float(r["分配比例"])
                    personal_pay = float(r["居服員實領"])
                    a_sum_pay += personal_pay

                    vals = [
                        r[col_client],
                        round(float(r["該案總服務次數"]), 1) if pd.notna(r.get("該案總服務次數")) else 0,
                        round(float(r["個人服務次數"]), 1) if pd.notna(r.get("個人服務次數")) else 0,
                        r[col_code],
                        round(float(r["核銷總額"]), 0),
                        round(personal_gross, 1),
                        round(personal_pay, 0)
                    ]
                    for c_i, val in enumerate(vals, start=1):
                        cell = ws.cell(row=row_idx, column=c_i, value=val)
                        cell.font = font_body
                        cell.border = border_box
                        if c_i in [1, 4]:
                            cell.alignment = align_center
                        elif c_i in [2, 3]:
                            cell.alignment = align_center
                            cell.number_format = "#,##0.#"
                        elif c_i == 6:
                            cell.alignment = align_right
                            cell.number_format = "#,##0.0"
                        else:
                            cell.alignment = align_right
                            cell.number_format = "#,##0"
            else:
                row_idx += 1
                ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=7)
                cell = ws.cell(row=row_idx, column=1, value="本月無 A 碼分攤紀錄")
                cell.font = font_body
                cell.alignment = align_center
                for c_i in range(1, 8):
                    ws.cell(row=row_idx, column=c_i).border = border_box

            row_idx += 1
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)
            ws.cell(row=row_idx, column=1, value="A碼總計").alignment = align_right
            ws.cell(row=row_idx, column=7, value=round(a_sum_pay, 0)).number_format = "#,##0"
            ws.cell(row=row_idx, column=7).alignment = align_right
            for c_i in range(1, 8):
                c = ws.cell(row=row_idx, column=c_i)
                c.font = font_total
                c.fill = fill_total
                c.border = border_box

            # 下半部：B 碼服務薪資明細表（採平地計薪基準）
            row_idx += 2
            ws.cell(row=row_idx, column=1, value="【二、B碼照顧服務費明細（依公司規定採平地標準拆帳）】").font = font_sec

            row_idx += 1
            b_headers = [
                "個案姓名",
                "服務項目類別",
                "服務次數(分攤)",
                "平地計薪金額",
                f"公司拆帳{b_share_pct}%"
            ]
            for c_i, h_text in enumerate(b_headers, start=1):
                cell = ws.cell(row=row_idx, column=c_i, value=h_text)
                cell.font = font_header
                cell.fill = fill_header_b
                cell.alignment = align_center
                cell.border = border_box
            ws.row_dimensions[row_idx].height = 22

            w_b_df = b_df[b_df[col_worker] == worker] if not b_df.empty else pd.DataFrame()
            b_sum_pay = 0.0

            if not w_b_df.empty:
                b_grouped = (
                    w_b_df.groupby([col_client, col_code], as_index=False)
                    .agg({
                        col_count: "sum",
                        "核銷總額(平地基準)": "sum",
                        "居服員實領": "sum"
                    })
                )
                for _, r in b_grouped.iterrows():
                    row_idx += 1
                    b_pay = float(r["居服員實領"])
                    b_sum_pay += b_pay
                    vals = [
                        r[col_client],
                        r[col_code],
                        round(float(r[col_count]), 1),
                        round(float(r["核銷總額(平地基準)"]), 0),
                        round(b_pay, 0)
                    ]
                    for c_i, val in enumerate(vals, start=1):
                        cell = ws.cell(row=row_idx, column=c_i, value=val)
                        cell.font = font_body
                        cell.border = border_box
                        if c_i == 1:
                            cell.alignment = align_center
                        elif c_i == 2:
                            cell.alignment = align_left
                        elif c_i == 3:
                            cell.alignment = align_center
                            cell.number_format = "#,##0.#"
                        else:
                            cell.alignment = align_right
                            cell.number_format = "#,##0"
            else:
                row_idx += 1
                ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=5)
                ws.cell(row=row_idx, column=1, value="本月無 B 碼服務紀錄").alignment = align_center
                for c_i in range(1, 6):
                    ws.cell(row=row_idx, column=c_i).border = border_box

            row_idx += 1
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
            ws.cell(row=row_idx, column=1, value="B碼總計").alignment = align_right
            ws.cell(row=row_idx, column=5, value=round(b_sum_pay, 0)).number_format = "#,##0"
            ws.cell(row=row_idx, column=5).alignment = align_right
            for c_i in range(1, 6):
                c = ws.cell(row=row_idx, column=c_i)
                c.font = font_total
                c.fill = fill_total
                c.border = border_box

            row_idx += 1
            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=4)
            ws.cell(row=row_idx, column=1, value="本月應發總薪資（A碼總計 + B碼總計）").alignment = align_right
            ws.cell(row=row_idx, column=5, value=round(a_sum_pay + b_sum_pay, 0)).number_format = "#,##0"
            ws.cell(row=row_idx, column=5).alignment = align_right
            for c_i in range(1, 6):
                c = ws.cell(row=row_idx, column=c_i)
                c.font = font_grand
                c.fill = fill_grand
                c.border = border_box
            ws.row_dimensions[row_idx].height = 24

            col_widths = {1: 14, 2: 26, 3: 16, 4: 18, 5: 16, 6: 18, 7: 16}
            for col_i, width in col_widths.items():
                ws.column_dimensions[get_column_letter(col_i)].width = width

        # 若為單人匯出，刪除 openpyxl 預設留下的空白預設工作表（如果有的話）
        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb["Sheet"]

    return output.getvalue()

def render_home_care_payroll():
    st.title("💰 居家服務薪資拆帳與核銷對帳引擎")
    st.caption("自動拆解多人串接姓名，支援【單人獨立薪資單下載列印】與【全機構一人一分頁整批匯出】。")

    with st.expander("⚙️ 機構拆帳內規與報表抬頭設定", expanded=True):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**1️⃣ 薪資月份與分潤成數**")
            month_label = st.text_input("薪資年月抬頭（顯示於個人薪資單）", value="11506")
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
        exact_plain_fee = next((c for c in s_cols if str(c).strip() == "申請(補助)費用"), None)
        default_fee = exact_plain_fee if exact_plain_fee else auto_match_col(s_cols, ["申請(補助)費用", "補助費用", "申請費用", "小計", "實際補助"])
        default_fee_idx = fee_options.index(default_fee) if default_fee in s_cols and any(k in str(default_fee) for k in ["費用", "金額", "小計", "補助"]) else 0
        col_direct_fee = st.selectbox("B碼計薪基準欄位 (平地費用)", fee_options, index=default_fee_idx)

    region_options = ["(無此欄位，全算平地)"] + s_cols
    default_reg = auto_match_col(s_cols, ["原民", "行政區", "居住行政區", "區域", "地區"])
    default_reg_idx = region_options.index(default_reg) if default_reg in s_cols and any(k in str(default_reg) for k in ["原民", "行政區", "區域"]) else 0
    col_region = st.selectbox("原民/平地判斷欄位 (供無費用欄位查表時使用)", region_options, index=default_reg_idx)

    if not df_a_claim.empty:
        st.markdown("**A碼核銷表欄位對應：**")
        ac_cols = list(df_a_claim.columns)
        ac1, ac2, ac3, ac4 = st.columns(4)
        with ac1:
            ac_client = st.selectbox("核銷表-個案姓名/ID", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["個案姓名", "個案", "案主", "姓名"])))
        with ac2:
            ac_code = st.selectbox("核銷表-A碼項目", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["服務代碼", "代碼", "項目", "碼別"])))
        with ac3:
            ac_qty = st.selectbox("核銷表-數量/天數", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["數量", "次數", "天數"])))
        with ac4:
            ac_amount = st.selectbox("核銷表-核銷總金額", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["小計", "核銷金額", "申報金額", "總金額", "金額", "費用"])))

    if st.button("🚀 開始自動拆帳並產生個人薪資單", type="primary", use_container_width=True):
        work_df = df_service.copy()
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

        work_df["_worker_list"] = work_df[col_worker].apply(
            lambda x: [w.strip() for w in re.split(r"[、,，/]+", str(x)) if w.strip()]
        )
        work_df["_split_n"] = work_df["_worker_list"].apply(lambda lst: max(len(lst), 1))
        work_df[col_count] = work_df[col_count] / work_df["_split_n"]
        if col_direct_fee != "(無，依代碼查表計算)":
            work_df[col_direct_fee] = pd.to_numeric(work_df[col_direct_fee], errors='coerce').fillna(0) / work_df["_split_n"]

        work_df = work_df.explode("_worker_list").reset_index(drop=True)
        work_df[col_worker] = work_df["_worker_list"]

        # Part A：雙模計算 B 碼薪資
        if col_direct_fee != "(無，依代碼查表計算)":
            b_df = work_df[~work_df[col_code].str.startswith(("AA", "GA", "A碼", "個案管理"))].copy()
            b_df["核銷總額(平地基準)"] = b_df[col_direct_fee]
            b_df["居服員實領"] = b_df["核銷總額(平地基準)"] * (b_share_pct / 100.0)
        else:
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

        # Part B：計算 A 碼拆帳
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
            a_base = df_a_claim[[ac_client, ac_code, ac_qty, ac_amount]].copy()
            a_base.columns = [col_client, col_code, "_a_qty", "清冊核銷原額"]
            a_base = a_base[a_base[col_client].notna() & a_base[col_code].notna()].copy()
            a_base = a_base[~a_base[col_client].astype(str).str.contains("總計|合計|小計", na=False)].copy()
            a_base[col_client] = a_base[col_client].astype(str).str.strip()
            a_base[col_code] = a_base[col_code].astype(str).str.strip().str.upper()
            a_base["_a_qty"] = pd.to_numeric(a_base["_a_qty"], errors='coerce').fillna(1)
            a_base["清冊核銷原額"] = pd.to_numeric(a_base["清冊核銷原額"], errors='coerce').fillna(0)

            # 判斷是否為定額實領碼 (AA09 / GA09)
            is_fixed_code = a_base[col_code].str.startswith(("AA09", "GA09"))
            
            # 一般 A 碼：居服員分潤池 = 清冊核銷原額 × 50% (a_share_pct)
            # AA09 / GA09：不走 55 分帳，居服員分潤池直接就是 500元 × 天數
            a_base["核銷總額"] = a_base["清冊核銷原額"]
            a_base["居服員總分潤池"] = a_base["清冊核銷原額"] * (a_share_pct / 100.0)
            
            mask_aa09 = a_base[col_code].str.startswith("AA09")
            mask_ga09 = a_base[col_code].str.startswith("GA09")
            a_base.loc[mask_aa09, "居服員總分潤池"] = a_base.loc[mask_aa09, "_a_qty"] * custom_aa09
            a_base.loc[mask_ga09, "居服員總分潤池"] = a_base.loc[mask_ga09, "_a_qty"] * custom_ga09

            a_base = a_base.groupby([col_client, col_code], as_index=False)[["核銷總額", "居服員總分潤池"]].sum()
        else:
            a_raw = work_df[work_df[col_code].str.startswith(("AA", "GA"))].copy()
            def calc_a_pool(row):
                code = row[col_code]
                if code.startswith("AA09"):
                    gross = 770 * row[col_count]
                    pool = custom_aa09 * row[col_count]
                elif code.startswith("GA09"):
                    gross = 770 * row[col_count]
                    pool = custom_ga09 * row[col_count]
                else:
                    price_info = price_map.get(code, {"plain": 0, "indigenous": 0})
                    unit_price = price_info["indigenous"] if (a_use_indigenous and row["is_indigenous"]) else price_info["plain"]
                    gross = unit_price * row[col_count]
                    pool = gross * (a_share_pct / 100.0)
                return pd.Series([gross, pool])
            a_raw[["核銷總額", "居服員總分潤池"]] = a_raw.apply(calc_a_pool, axis=1)
            a_base = a_raw.groupby([col_client, col_code], as_index=False)[["核銷總額", "居服員總分潤池"]].sum()

        a_split_df = pd.merge(a_base, ratio_df, on=col_client, how="left")
        a_split_df[col_worker] = a_split_df[col_worker].fillna("⚠️ 無班表紀錄(需手動指定)")
        a_split_df["分配比例"] = a_split_df["分配比例"].fillna(1.0)
        
        if split_method == "依人頭均分":
            headcount = a_split_df.groupby([col_client, col_code])[col_worker].transform("count")
            a_split_df["分配比例"] = 1.0 / headcount

        # 直接用「居服員總分潤池 × 多人分配比例」算出實領，AA09/GA09 不會再被打折！
        a_split_df["居服員實領"] = a_split_df["居服員總分潤池"] * a_split_df["分配比例"]
        if round_to_int:
            a_split_df["居服員實領"] = a_split_df["居服員實領"].round(0)
        a_split_df["公司留存"] = (a_split_df["核銷總額"] * a_split_df["分配比例"]) - a_split_df["居服員實領"]

        # Part C：彙總每位居服員薪資總表
        b_summary = b_df.groupby(col_worker)[["核銷總額(平地基準)", "居服員實領"]].sum().reset_index().rename(
            columns={"核銷總額(平地基準)": "B碼計薪總額(平地)", "居服員實領": "B碼實領薪資"}
        ) if not b_df.empty else pd.DataFrame(columns=[col_worker, "B碼計薪總額(平地)", "B碼實領薪資"])

        a_summary = a_split_df.groupby(col_worker)[["核銷總額", "居服員實領"]].sum().reset_index().rename(
            columns={"核銷總額": "A碼核銷總額", "居服員實領": "A碼實領薪資"}
        ) if not a_split_df.empty else pd.DataFrame(columns=[col_worker, "A碼核銷總額", "A碼實領薪資"])

        final_payroll = pd.merge(b_summary, a_summary, on=col_worker, how="outer").fillna(0)
        final_payroll["本月應發總薪資(A+B)"] = final_payroll["B碼實領薪資"] + final_payroll["A碼實領薪資"]

        all_excel_bytes = build_formatted_workbook(
            final_payroll, a_split_df, b_df,
            col_client, col_worker, col_code, col_count,
            month_label, a_share_pct, b_share_pct,
            target_workers=None
        )

        st.session_state["payroll_result"] = {
            "final_payroll": final_payroll,
            "a_split_df": a_split_df,
            "b_df": b_df,
            "all_excel_bytes": all_excel_bytes,
            "col_client": col_client,
            "col_worker": col_worker,
            "col_code": col_code,
            "col_count": col_count,
            "month_label": month_label,
            "a_share_pct": a_share_pct,
            "b_share_pct": b_share_pct
        }

    if "payroll_result" in st.session_state:
        res = st.session_state["payroll_result"]
        final_payroll = res["final_payroll"]
        a_split_df = res["a_split_df"]
        b_df = res["b_df"]
        c_client = res["col_client"]
        c_worker = res["col_worker"]
        c_code = res["col_code"]
        c_count = res["col_count"]
        m_label = res["month_label"]

        valid_workers = [w for w in final_payroll[c_worker].tolist() if not str(w).startswith("⚠️")]
        st.success(f"✅ 計算完成！共為 {len(valid_workers)} 位居服員完成拆帳核算：")
        
        # 全機構總計
        st.markdown("#### 🏢 全機構薪資總計")
        m1, m2, m3 = st.columns(3)
        m1.metric("全機構應發總薪資 (A+B)", f"${final_payroll['本月應發總薪資(A+B)'].sum():,.0f}")
        m2.metric("全機構 A碼實領總額", f"${final_payroll['A碼實領薪資'].sum():,.0f}")
        m3.metric("全機構 B碼實領總額", f"${final_payroll['B碼實領薪資'].sum():,.0f}")

        st.divider()

        # 單人預覽與單人獨立列印下載
        preview_list = valid_workers + [w for w in final_payroll[c_worker].tolist() if str(w).startswith("⚠️")]
        preview_worker = st.selectbox("🔍 選擇居服員檢視個人明細與下載單人薪資單：", preview_list)
        
        # 計算該位居服員的個人總額卡片
        w_row = final_payroll[final_payroll[c_worker] == preview_worker].iloc[0]
        w_a_total = w_row["A碼實領薪資"]
        w_b_total = w_row["B碼實領薪資"]
        w_grand_total = w_row["本月應發總薪資(A+B)"]

        st.markdown(f"#### 👤 【{preview_worker}】個人薪資總計")
        wm1, wm2, wm3 = st.columns(3)
        wm1.metric(f"{preview_worker} 本月應發總薪資", f"${w_grand_total:,.0f}")
        wm2.metric("A碼拆帳總計", f"${w_a_total:,.0f}")
        wm3.metric("B碼平地拆帳總計", f"${w_b_total:,.0f}")

        p_col1, p_col2 = st.columns(2)
        with p_col1:
            st.markdown(f"**{preview_worker} — A碼分攤明細**")
            w_a_view = a_split_df[a_split_df[c_worker] == preview_worker][[c_client, "該案總服務次數", "個人服務次數", c_code, "核銷總額", "居服員實領"]].copy()
            w_a_view = w_a_view.rename(columns={"核銷總額": "A碼核銷金額", "居服員實領": f"A碼實領({res['a_share_pct']}%)"})
            st.dataframe(w_a_view, use_container_width=True, hide_index=True)
        with p_col2:
            st.markdown(f"**{preview_worker} — B碼平地計薪明細**")
            if not b_df.empty:
                w_b_view = b_df[b_df[c_worker] == preview_worker].groupby([c_client, c_code], as_index=False).agg({c_count: "sum", "核銷總額(平地基準)": "sum", "居服員實領": "sum"})
                w_b_view = w_b_view.rename(columns={c_count: "服務次數", "核銷總額(平地基準)": "平地計薪金額", "居服員實領": f"B碼實領({res['b_share_pct']}%)"})
                st.dataframe(w_b_view, use_container_width=True, hide_index=True)

        # 產生該位居服員專屬的單人 Excel 檔
        single_excel_bytes = build_formatted_workbook(
            final_payroll, a_split_df, b_df,
            c_client, c_worker, c_code, c_count,
            m_label, res["a_share_pct"], res["b_share_pct"],
            target_workers=[preview_worker]
        )

        # 雙下載按鈕：左邊下載單人專屬檔（可轉PDF或直接列印），右邊下載全機構整批檔
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                label=f"🖨️ 下載【{m_label} {preview_worker} 薪資服務明細.xlsx】（單人獨立檔）",
                data=single_excel_bytes,
                file_name=f"{m_label}_{preview_worker}_薪資服務明細.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with dl_col2:
            st.download_button(
                label=f"📥 下載【{m_label} 全機構居服員薪資服務明細.xlsx】（一人一分頁）",
                data=res["all_excel_bytes"],
                file_name=f"{m_label}_全機構居服員薪資服務明細.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )