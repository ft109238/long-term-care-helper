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

def build_formatted_workbook(final_payroll, a_split_df, b_df, col_client, col_worker, col_code, col_count, org_name, month_label, a_share_pct, b_share_pct, target_workers=None):
    """100% 還原機構專屬版型（上方 A碼服務明細 ＋ 左下薪資明細表 ＋ 右下個案BA碼統計細項 + 自動連動公式）"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wb = writer.book
        if target_workers is None:
            final_payroll.to_excel(writer, sheet_name="0_全機構薪資總表", index=False)
            workers = [w for w in final_payroll[col_worker].dropna().unique() if not str(w).startswith("⚠️")]
        else:
            workers = target_workers

        # 字型與樣式設定
        f_org = Font(name="微軟正黑體", size=12, bold=True)
        f_sub = Font(name="微軟正黑體", size=12, bold=True)
        f_worker = Font(name="微軟正黑體", size=11, bold=True)
        f_th = Font(name="微軟正黑體", size=10, bold=True)
        f_td = Font(name="微軟正黑體", size=10)
        f_td_bold = Font(name="微軟正黑體", size=10, bold=True)
        f_red_tot = Font(name="微軟正黑體", size=11, bold=True)
        f_right_title = Font(name="微軟正黑體", size=10, bold=True)

        fill_pink = PatternFill(start_color="E6B8B7", end_color="E6B8B7", fill_type="solid")
        fill_peach = PatternFill(start_color="FCD5B4", end_color="FCD5B4", fill_type="solid")
        fill_gray_hdr = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        fill_gray_sub = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

        thin = Side(border_style="thin", color="000000")
        med = Side(border_style="medium", color="000000")
        gray_thin = Side(border_style="thin", color="808080")

        b_grid = Border(left=thin, right=thin, top=thin, bottom=thin)
        b_right_grid = Border(left=gray_thin, right=gray_thin, top=gray_thin, bottom=gray_thin)

        al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
        al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
        al_r = Alignment(horizontal="right", vertical="center")

        for worker in workers:
            safe_name = str(worker)[:30].replace("/", "_").replace("\\", "_").replace("*", "").replace("?", "").replace("[", "").replace("]", "")
            ws = wb.create_sheet(title=safe_name)

            # A4 單頁寬自動縮放列印設定
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1
            ws.sheet_properties.pageSetUpPr.fitToPage = True

            # =========================================================
            # 上半部：A碼服務明細 (Columns A ~ G)
            # =========================================================
            ws.merge_cells("A1:G1")
            ws["A1"] = org_name
            ws["A1"].font = f_org
            ws["A1"].alignment = al_c

            ws.merge_cells("A2:G2")
            ws["A2"] = "A碼服務明細"
            ws["A2"].font = f_sub
            ws["A2"].alignment = al_c

            ws.merge_cells("A3:B3")
            ws["A3"] = f"居服員: {worker}"
            ws["A3"].font = f_worker
            ws["A3"].alignment = al_l

            a_headers = ["服務代碼", "服務項目類別", "個案姓名", "給付價格", "數量", "小計", "服務日期"]
            r_idx = 4
            for c_i, h in enumerate(a_headers, start=1):
                cell = ws.cell(row=r_idx, column=c_i, value=h)
                cell.font = f_th
                cell.alignment = al_c
                cell.border = Border(left=thin, right=thin, top=med, bottom=med)
            ws.row_dimensions[r_idx].height = 22

            w_a_df = a_split_df[a_split_df[col_worker] == worker].copy() if not a_split_df.empty else pd.DataFrame()
            
            # 拆成「一般 A 碼 (AA01~AA08)」與「假日碼 (AA09/GA09)」兩區塊
            if not w_a_df.empty:
                df_a_reg = w_a_df[~w_a_df[col_code].str.startswith(("AA09", "GA09"))].copy()
                df_a_hol = w_a_df[w_a_df[col_code].str.startswith(("AA09", "GA09"))].copy()
            else:
                df_a_reg = pd.DataFrame()
                df_a_hol = pd.DataFrame()

            # 1) 一般 A 碼明細列
            reg_start_row = r_idx + 1
            if not df_a_reg.empty:
                for _, r in df_a_reg.iterrows():
                    r_idx += 1
                    row_vals = [
                        r[col_code],
                        r.get("_a_name", ""),
                        r[col_client],
                        round(float(r["核銷總額"]), 0),
                        round(float(r.get("_a_qty", 1)), 1),
                        round(float(r["居服員實領"]), 0),
                        r.get("_a_dates", "")
                    ]
                    for c_i, v in enumerate(row_vals, start=1):
                        c = ws.cell(row=r_idx, column=c_i, value=v)
                        c.font = f_td
                        c.border = b_grid
                        c.alignment = al_l if c_i in [2, 7] else al_c
                        if c_i in [4, 6]:
                            c.number_format = "#,##0"
                    ws.row_dimensions[r_idx].height = 24
            else:
                r_idx += 1
                for c_i in range(1, 8):
                    c = ws.cell(row=r_idx, column=c_i, value="-" if c_i == 1 else "")
                    c.font = f_td
                    c.border = b_grid
                    c.alignment = al_c

            reg_end_row = r_idx
            # 一般 A 碼粉紅底總計列
            r_idx += 1
            reg_tot_row = r_idx
            for c_i in range(1, 8):
                c = ws.cell(row=r_idx, column=c_i)
                c.border = b_grid
                if c_i == 6:
                    c.value = f"=SUM(F{reg_start_row}:F{reg_end_row})"
                    c.font = f_red_tot
                    c.fill = fill_pink
                    c.alignment = al_c
                    c.number_format = "#,##0"
            ws.row_dimensions[r_idx].height = 22

            # 2) 假日碼 (AA09 / GA09) 明細列
            hol_start_row = r_idx + 1
            if not df_a_hol.empty:
                for _, r in df_a_hol.iterrows():
                    r_idx += 1
                    qty_val = float(r.get("_a_qty", 1))
                    unit_val = round(float(r["居服員實領"]) / qty_val, 0) if qty_val > 0 else 500
                    row_vals = [
                        r[col_code],
                        r.get("_a_name", "例假日服務"),
                        r[col_client],
                        unit_val,
                        round(qty_val, 1),
                        round(float(r["居服員實領"]), 0),
                        r.get("_a_dates", "")
                    ]
                    for c_i, v in enumerate(row_vals, start=1):
                        c = ws.cell(row=r_idx, column=c_i, value=v)
                        c.font = f_td
                        c.border = b_grid
                        c.alignment = al_l if c_i in [2, 7] else al_c
                        if c_i in [4, 6]:
                            c.number_format = "#,##0"
                    ws.row_dimensions[r_idx].height = 20
            else:
                r_idx += 1
                for c_i in range(1, 8):
                    c = ws.cell(row=r_idx, column=c_i, value="-" if c_i == 1 else "")
                    c.font = f_td
                    c.border = b_grid
                    c.alignment = al_c

            hol_end_row = r_idx
            # 假日碼粉紅底總計列
            r_idx += 1
            hol_tot_row = r_idx
            for c_i in range(1, 8):
                c = ws.cell(row=r_idx, column=c_i)
                c.border = Border(left=thin, right=thin, top=thin, bottom=med)
                if c_i == 6:
                    c.value = f"=SUM(F{hol_start_row}:F{hol_end_row})"
                    c.font = f_red_tot
                    c.fill = fill_pink
                    c.alignment = al_c
                    c.number_format = "#,##0"
            ws.row_dimensions[r_idx].height = 24

            # =========================================================
            # 下半部左側：薪資明細表 (Columns A ~ D)
            # =========================================================
            r_idx += 1
            ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=4)
            ws.cell(row=r_idx, column=1, value=org_name).font = f_th
            ws.cell(row=r_idx, column=1).alignment = al_l

            r_idx += 1
            ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=4)
            # 將 11506 轉成 115年06月薪資明細表
            m_display = f"{month_label[:3]}年{month_label[3:]}月薪資明細表" if len(month_label) == 5 and month_label.isdigit() else f"{month_label} 薪資明細表"
            ws.cell(row=r_idx, column=1, value=m_display).font = f_th
            ws.cell(row=r_idx, column=1).alignment = al_c

            r_idx += 1
            name_row = r_idx
            ws.cell(row=name_row, column=1, value="姓名:").font = f_td_bold
            ws.cell(row=name_row, column=1).alignment = al_l
            ws.cell(row=name_row, column=2, value=worker).font = f_sub
            ws.cell(row=name_row, column=2).alignment = al_c

            # 右側表格從 name_row 同一列開始對齊！「個案服務拆帳統計細項」(Columns E ~ G)
            right_start_row = name_row
            ws.merge_cells(start_row=right_start_row, start_column=5, end_row=right_start_row, end_column=7)
            rc_title = ws.cell(row=right_start_row, column=5, value="個案服務拆帳統計細項")
            rc_title.font = f_right_title
            rc_title.fill = fill_gray_hdr
            rc_title.alignment = al_c
            for ci in range(5, 8):
                ws.cell(row=right_start_row, column=ci).border = b_right_grid

            # 左表標題列
            left_hdr_row = name_row + 1
            for ci, txt in [(1, ""), (2, "項目"), (3, "金額"), (4, "備註")]:
                c = ws.cell(row=left_hdr_row, column=ci, value=txt)
                c.font = f_th
                c.alignment = al_c
                c.border = Border(left=thin, right=thin, top=med, bottom=thin)

            # 右表標題列
            ws.cell(row=left_hdr_row, column=5, value="").fill = fill_gray_sub
            ws.cell(row=left_hdr_row, column=6, value="個案姓名").font = f_th
            ws.cell(row=left_hdr_row, column=6).fill = fill_gray_sub
            ws.cell(row=left_hdr_row, column=6).alignment = al_c
            ws.cell(row=left_hdr_row, column=7, value="BA碼").font = f_th
            ws.cell(row=left_hdr_row, column=7).fill = fill_gray_sub
            ws.cell(row=left_hdr_row, column=7).alignment = al_c
            for ci in range(5, 8):
                ws.cell(row=left_hdr_row, column=ci).border = b_right_grid

            # 準備右下角個案 B 碼彙總資料（依個案姓名加總實領金額）
            w_b_df = b_df[b_df[col_worker] == worker] if not b_df.empty else pd.DataFrame()
            if not w_b_df.empty:
                client_b_summary = w_b_df.groupby(col_client, as_index=False, sort=False)["居服員實領"].sum()
            else:
                client_b_summary = pd.DataFrame(columns=[col_client, "居服員實領"])

            # 寫入右側 1~20 列明細與底部「薪資小計 (A)」
            r_data_start = left_hdr_row + 1
            for i in range(20):
                curr_r = r_data_start + i
                c_idx_cell = ws.cell(row=curr_r, column=5, value=i + 1)
                c_idx_cell.font = f_td
                c_idx_cell.alignment = al_c
                c_idx_cell.border = b_right_grid

                if i < len(client_b_summary):
                    c_name_val = client_b_summary.iloc[i][col_client]
                    c_amt_val = round(float(client_b_summary.iloc[i]["居服員實領"]), 0)
                else:
                    c_name_val = ""
                    c_amt_val = ""

                c_name_cell = ws.cell(row=curr_r, column=6, value=c_name_val)
                c_name_cell.font = f_td
                c_name_cell.alignment = al_c
                c_name_cell.border = b_right_grid

                c_amt_cell = ws.cell(row=curr_r, column=7, value=c_amt_val)
                c_amt_cell.font = f_td
                c_amt_cell.alignment = al_r
                c_amt_cell.number_format = '#,##0" 元"' if c_amt_val != "" else '@" 元"'
                c_amt_cell.border = b_right_grid

            r_data_end = r_data_start + 19
            r_sub_row = r_data_end + 1
            ws.merge_cells(start_row=r_sub_row, start_column=5, end_row=r_sub_row, end_column=6)
            ws.cell(row=r_sub_row, column=5, value="薪資小計 (A) :").font = f_th
            ws.cell(row=r_sub_row, column=5).alignment = al_r
            sub_val_cell = ws.cell(row=r_sub_row, column=7, value=f"=SUM(G{r_data_start}:G{r_data_end})")
            sub_val_cell.font = f_td_bold
            sub_val_cell.alignment = al_r
            sub_val_cell.number_format = '#,##0" 元"'
            for ci in range(5, 8):
                ws.cell(row=r_sub_row, column=ci).fill = fill_gray_hdr
                ws.cell(row=r_sub_row, column=ci).border = b_right_grid

            # 寫入左下角「應領金額」12 列 + 「應扣金額」6 列 + 「實領薪資」1 列（共 19 列）
            earn_start = left_hdr_row + 1
            earn_items = [
                ("工作薪資", f"=G{r_sub_row}"),          # 自動連動右側 B碼小計
                ("偏鄉交通補助", 0),
                ("偏鄉獎勵津貼", 0),
                ("交通費", 0),
                ("轉場津貼", 0),
                ("AA05.06.07", f"=F{reg_tot_row}"),     # 自動連動上方一般A碼總計
                ("A碼假日(AA09)", f"=F{hol_tot_row}"),  # 自動連動上方假日碼總計
                ("AA11 (身障/失智)", 0),
                ("喘息", 0),
                ("其它", 0),
            ]
            for idx_e, (label, val) in enumerate(earn_items):
                cr = earn_start + idx_e
                ws.cell(row=cr, column=2, value=label).font = f_td
                ws.cell(row=cr, column=2).alignment = al_l
                vc = ws.cell(row=cr, column=3, value=val)
                vc.font = f_td
                vc.alignment = al_c
                vc.number_format = "#,##0;-#,##0;\"-\""
                for ci in range(1, 5):
                    ws.cell(row=cr, column=ci).border = b_grid

            earn_tot_row = earn_start + len(earn_items)
            ws.cell(row=earn_tot_row, column=2, value="總計").font = f_td_bold
            ws.cell(row=earn_tot_row, column=2).alignment = al_l
            et_cell = ws.cell(row=earn_tot_row, column=3, value=f"=SUM(C{earn_start}:C{earn_tot_row-1})")
            et_cell.font = f_td_bold
            et_cell.alignment = al_c
            et_cell.number_format = "#,##0"
            for ci in range(1, 5):
                ws.cell(row=earn_tot_row, column=ci).border = b_grid

            # 合併左側「應領金額」垂直儲存格
            ws.merge_cells(start_row=earn_start, start_column=1, end_row=earn_tot_row, end_column=1)
            ws.cell(row=earn_start, column=1, value="應領金額").font = f_th
            ws.cell(row=earn_start, column=1).alignment = al_c

            # 應扣金額區塊
            ded_start = earn_tot_row + 1
            ded_items = [
                ("勞保費", 0),
                ("健保費", 0),
                ("法扣", 0),
                ("其他", 0),
                ("福利金", 0),
            ]
            for idx_d, (label, val) in enumerate(ded_items):
                cr = ded_start + idx_d
                ws.cell(row=cr, column=2, value=label).font = f_td
                ws.cell(row=cr, column=2).alignment = al_l
                vc = ws.cell(row=cr, column=3, value=val)
                vc.font = f_td
                vc.alignment = al_c
                vc.number_format = "#,##0"
                for ci in range(1, 5):
                    ws.cell(row=cr, column=ci).border = b_grid

            ded_tot_row = ded_start + len(ded_items)
            ws.cell(row=ded_tot_row, column=2, value="總計").font = f_td_bold
            ws.cell(row=ded_tot_row, column=2).alignment = al_l
            dt_cell = ws.cell(row=ded_tot_row, column=3, value=f"=SUM(C{ded_start}:C{ded_tot_row-1})")
            dt_cell.font = f_td_bold
            dt_cell.alignment = al_c
            dt_cell.number_format = "#,##0"
            for ci in range(1, 5):
                ws.cell(row=ded_tot_row, column=ci).border = b_grid

            ws.merge_cells(start_row=ded_start, start_column=1, end_row=ded_tot_row, end_column=1)
            ws.cell(row=ded_start, column=1, value="應扣金額").font = f_th
            ws.cell(row=ded_start, column=1).alignment = al_c

            # 最底列：實領薪資 (=應領總計 - 應扣總計)
            net_row = ded_tot_row + 1
            ws.merge_cells(start_row=net_row, start_column=1, end_row=net_row, end_column=2)
            ws.cell(row=net_row, column=1, value="實領薪資").font = f_th
            ws.cell(row=net_row, column=1).alignment = al_c
            net_cell = ws.cell(row=net_row, column=3, value=f"=C{earn_tot_row}-C{ded_tot_row}")
            net_cell.font = f_red_tot
            net_cell.fill = fill_peach
            net_cell.alignment = al_c
            net_cell.number_format = "#,##0"
            for ci in range(1, 5):
                ws.cell(row=net_row, column=ci).border = Border(left=thin, right=thin, top=thin, bottom=med)
            ws.row_dimensions[net_row].height = 24

            # 欄寬設定（完美還原比例）
            col_widths = {1: 11, 2: 24, 3: 13, 4: 11, 5: 8, 6: 14, 7: 42}
            for col_i, width in col_widths.items():
                ws.column_dimensions[get_column_letter(col_i)].width = width

        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb["Sheet"]

    return output.getvalue()

def render_home_care_payroll():
    st.title("💰 居家服務薪資拆帳與核銷對帳引擎")
    st.caption("100% 還原機構專屬【A碼服務明細 ＋ 薪資明細表 ＋ 個案服務拆帳統計細項】雙欄自動連動報表。")

    with st.expander("⚙️ 機構抬頭、薪資年月與拆帳內規設定", expanded=True):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**1️⃣ 機構名稱與薪資年月**")
            org_name = st.text_input("機構全銜", value="南門長照社團法人附設屏東縣私立南門綜合長照機構")
            month_label = st.text_input("薪資年月代碼（如 11506）", value="11506")
            
        with col2:
            st.markdown("**2️⃣ A/B 碼分潤與定額津貼**")
            a_share_pct = st.slider("一般 A碼 居服員分潤 (%)", 0, 100, 50, step=5)
            b_share_pct = st.slider("B碼 居服員分潤 (%)", 0, 100, 60, step=5)
            custom_aa09 = st.number_input("AA09 假日定額實領 (元/天)", value=500, step=50)
            custom_ga09 = st.number_input("GA09 假日定額實領 (元/天)", value=500, step=50)
            
        with col3:
            st.markdown("**3️⃣ 計價與進位規則**")
            a_use_indigenous = st.checkbox("一般 A碼 依核銷表計算", value=True)
            b_use_indigenous = st.checkbox("B碼 套用原民加成（未勾選全算平地）", value=False)
            split_method = st.selectbox("多人共照一般 A碼 分配依據", ["依服務次數比例", "依人頭均分"], index=0)
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
        st.info("👆 請先上傳 Jubo 匯出的「服務清冊」與「A碼清冊」Excel 檔以開始計算。")
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
    col_region = st.selectbox("原民/平地判斷欄位", region_options, index=default_reg_idx)

    if not df_a_claim.empty:
        st.markdown("**A碼核銷表欄位對應：**")
        ac_cols = list(df_a_claim.columns)
        ac1, ac2, ac3, ac4, ac5 = st.columns(5)
        with ac1:
            ac_client = st.selectbox("核銷表-個案姓名/ID", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["個案姓名", "個案", "案主", "姓名"])))
        with ac2:
            ac_code = st.selectbox("核銷表-A碼項目", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["服務代碼", "代碼", "項目", "碼別"])))
        with ac3:
            ac_qty = st.selectbox("核銷表-數量/天數", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["數量", "次數", "天數"])))
        with ac4:
            ac_amount = st.selectbox("核銷表-核銷總金額", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["小計", "核銷金額", "申報金額", "總金額", "金額", "費用"])))
        with ac5:
            ac_worker_opts = ["(無，全依服務清冊比例分攤)"] + ac_cols
            def_ac_w = auto_match_col(ac_cols, ["服務人員", "居服員", "照服員", "人員"])
            def_ac_w_idx = ac_worker_opts.index(def_ac_w) if def_ac_w in ac_cols and any(k in str(def_ac_w) for k in ["服務人員", "居服員", "人員"]) else 0
            ac_worker = st.selectbox("核銷表-指定服務人員", ac_worker_opts, index=def_ac_w_idx)

        # 自動偵測 A碼核銷表的「服務項目類別(中文名稱)」與「服務日期」供報表完美呈現
        ac_name_col = next((c for c in ac_cols if "項目類別" in str(c) or "項目名稱" in str(c)), None)
        ac_date_col = next((c for c in ac_cols if "日期" in str(c)), None)

    if st.button("🚀 開始自動拆帳並產生個人薪資明細表", type="primary", use_container_width=True):
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

        # Part A：計算 B 碼薪資
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
            a_raw_claim = pd.DataFrame()
            a_raw_claim[col_client] = df_a_claim[ac_client]
            a_raw_claim[col_code] = df_a_claim[ac_code]
            a_raw_claim["_a_qty"] = df_a_claim[ac_qty]
            a_raw_claim["清冊核銷原額"] = df_a_claim[ac_amount]
            a_raw_claim["_direct_worker"] = df_a_claim[ac_worker] if ac_worker != "(無，全依服務清冊比例分攤)" else ""
            a_raw_claim["_a_name"] = df_a_claim[ac_name_col].fillna("") if ac_name_col else ""
            a_raw_claim["_a_dates"] = df_a_claim[ac_date_col].fillna("") if ac_date_col else ""

            a_raw_claim = a_raw_claim[a_raw_claim[col_client].notna() & a_raw_claim[col_code].notna()].copy()
            a_raw_claim = a_raw_claim[~a_raw_claim[col_client].astype(str).str.contains("總計|合計|小計", na=False)].copy()
            a_raw_claim[col_client] = a_raw_claim[col_client].astype(str).str.strip()
            a_raw_claim[col_code] = a_raw_claim[col_code].astype(str).str.strip().str.upper()
            a_raw_claim["_a_qty"] = pd.to_numeric(a_raw_claim["_a_qty"], errors='coerce').fillna(1)
            a_raw_claim["清冊核銷原額"] = pd.to_numeric(a_raw_claim["清冊核銷原額"], errors='coerce').fillna(0)
            a_raw_claim["_direct_worker"] = a_raw_claim["_direct_worker"].fillna("").astype(str).str.strip()
            a_raw_claim.loc[a_raw_claim["_direct_worker"].isin(["nan", "None", ""]), "_direct_worker"] = ""
            a_raw_claim["_a_name"] = a_raw_claim["_a_name"].astype(str).str.strip()
            a_raw_claim["_a_dates"] = a_raw_claim["_a_dates"].astype(str).str.strip()

            a_raw_claim["核銷總額"] = a_raw_claim["清冊核銷原額"]
            a_raw_claim["居服員總分潤池"] = a_raw_claim["清冊核銷原額"] * (a_share_pct / 100.0)

            mask_aa09 = a_raw_claim[col_code].str.startswith("AA09")
            mask_ga09 = a_raw_claim[col_code].str.startswith("GA09")
            a_raw_claim.loc[mask_aa09, "居服員總分潤池"] = a_raw_claim.loc[mask_aa09, "_a_qty"] * custom_aa09
            a_raw_claim.loc[mask_ga09, "居服員總分潤池"] = a_raw_claim.loc[mask_ga09, "_a_qty"] * custom_ga09

            is_single_assigned = (a_raw_claim["_direct_worker"] != "") & (~a_raw_claim["_direct_worker"].str.contains(r"[、,，/]"))
            assigned_df = a_raw_claim[is_single_assigned].copy()
            unassigned_df = a_raw_claim[~is_single_assigned].copy()

            split_parts = []
            def join_unique_str(s):
                vals = [str(x).strip() for x in s if str(x).strip() and str(x).strip() != "nan"]
                return ",".join(dict.fromkeys(vals))

            if not assigned_df.empty:
                assigned_grouped = assigned_df.groupby([col_client, col_code, "_direct_worker"], as_index=False).agg({
                    "核銷總額": "sum",
                    "居服員總分潤池": "sum",
                    "_a_qty": "sum",
                    "_a_name": join_unique_str,
                    "_a_dates": join_unique_str
                }).rename(columns={"_direct_worker": col_worker})
                assigned_grouped = pd.merge(assigned_grouped, ratio_df[[col_client, col_worker, "該案總服務次數", "個人服務次數"]], on=[col_client, col_worker], how="left")
                assigned_grouped["分配比例"] = 1.0
                split_parts.append(assigned_grouped)

            if not unassigned_df.empty:
                unassigned_grouped = unassigned_df.groupby([col_client, col_code], as_index=False).agg({
                    "核銷總額": "sum",
                    "居服員總分潤池": "sum",
                    "_a_qty": "sum",
                    "_a_name": join_unique_str,
                    "_a_dates": join_unique_str
                })
                merged_unassigned = pd.merge(unassigned_grouped, ratio_df, on=col_client, how="left")
                merged_unassigned[col_worker] = merged_unassigned[col_worker].fillna("⚠️ 無班表紀錄(需手動指定)")
                merged_unassigned["分配比例"] = merged_unassigned["分配比例"].fillna(1.0)
                if split_method == "依人頭均分":
                    headcount = merged_unassigned.groupby([col_client, col_code])[col_worker].transform("count")
                    merged_unassigned["分配比例"] = 1.0 / headcount
                split_parts.append(merged_unassigned)

            a_split_df = pd.concat(split_parts, ignore_index=True) if split_parts else pd.DataFrame()
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
            a_raw["_a_qty"] = a_raw[col_count]
            a_raw["_a_name"] = ""
            a_raw["_a_dates"] = ""
            a_base = a_raw.groupby([col_client, col_code], as_index=False)[["核銷總額", "居服員總分潤池", "_a_qty"]].sum()
            a_base["_a_name"] = ""
            a_base["_a_dates"] = ""
            a_split_df = pd.merge(a_base, ratio_df, on=col_client, how="left")
            a_split_df[col_worker] = a_split_df[col_worker].fillna("⚠️ 無班表紀錄(需手動指定)")
            a_split_df["分配比例"] = a_split_df["分配比例"].fillna(1.0)

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
            org_name, month_label, a_share_pct, b_share_pct,
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
            "org_name": org_name,
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
        st.success(f"✅ 計算完成！已為 {len(valid_workers)} 位居服員套用機構專屬格式並建立自動連動公式：")
        
        st.markdown("#### 🏢 全機構薪資總計")
        m1, m2, m3 = st.columns(3)
        m1.metric("全機構應發總薪資 (A+B)", f"${final_payroll['本月應發總薪資(A+B)'].sum():,.0f}")
        m2.metric("全機構 A碼實領總額", f"${final_payroll['A碼實領薪資'].sum():,.0f}")
        m3.metric("全機構 B碼實領總額", f"${final_payroll['B碼實領薪資'].sum():,.0f}")

        st.divider()

        preview_list = valid_workers + [w for w in final_payroll[c_worker].tolist() if str(w).startswith("⚠️")]
        preview_worker = st.selectbox("🔍 選擇居服員檢視個人明細與下載單人薪資明細表：", preview_list)
        
        w_row = final_payroll[final_payroll[c_worker] == preview_worker].iloc[0]
        w_a_total = w_row["A碼實領薪資"]
        w_b_total = w_row["B碼實領薪資"]
        w_grand_total = w_row["本月應發總薪資(A+B)"]

        st.markdown(f"#### 👤 【{preview_worker}】個人薪資總計")
        wm1, wm2, wm3 = st.columns(3)
        wm1.metric(f"{preview_worker} 本月應發總薪資", f"${w_grand_total:,.0f}")
        wm2.metric("A碼總計 (一般+假日)", f"${w_a_total:,.0f}")
        wm3.metric("B碼工作薪資小計(A)", f"${w_b_total:,.0f}")

        p_col1, p_col2 = st.columns(2)
        with p_col1:
            st.markdown(f"**{preview_worker} — A碼服務明細**")
            w_a_view = a_split_df[a_split_df[c_worker] == preview_worker][[c_code, "_a_name", c_client, "_a_qty", "居服員實領", "_a_dates"]].copy()
            w_a_view = w_a_view.rename(columns={"_a_name": "服務項目類別", "_a_qty": "數量", "居服員實領": "小計", "_a_dates": "服務日期"})
            st.dataframe(w_a_view, use_container_width=True, hide_index=True)
        with p_col2:
            st.markdown(f"**{preview_worker} — 個案服務拆帳統計細項 (BA碼)**")
            if not b_df.empty:
                w_b_view = b_df[b_df[c_worker] == preview_worker].groupby(c_client, as_index=False, sort=False)["居服員實領"].sum()
                w_b_view = w_b_view.rename(columns={"居服員實領": "BA碼金額(元)"})
                st.dataframe(w_b_view, use_container_width=True, hide_index=True)

        single_excel_bytes = build_formatted_workbook(
            final_payroll, a_split_df, b_df,
            c_client, c_worker, c_code, c_count,
            res["org_name"], m_label, res["a_share_pct"], res["b_share_pct"],
            target_workers=[preview_worker]
        )

        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                label=f"🖨️ 下載【{m_label} {preview_worker} 薪資明細表.xlsx】（完全還原機構版型）",
                data=single_excel_bytes,
                file_name=f"{m_label}_{preview_worker}_薪資服務明細.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with dl_col2:
            st.download_button(
                label=f"📥 下載【{m_label} 全機構居服員薪資明細表.xlsx】（一人一分頁）",
                data=res["all_excel_bytes"],
                file_name=f"{m_label}_全機構居服員薪資服務明細.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )