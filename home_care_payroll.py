import streamlit as st
import pdfplumber
import pandas as pd
import numpy as np
import io
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 公定平地給付單價表（主代碼與 -1/-2 自用備餐子代碼皆採平地全額單價）
DEFAULT_PRICE_TABLE = {
    "AA01": {"plain": 1700, "indigenous": 2040},
    "AA02": {"plain": 2000, "indigenous": 2400},
    "AA03": {"plain": 600,  "indigenous": 720},
    "AA04": {"plain": 600,  "indigenous": 720},
    "AA05": {"plain": 1200, "indigenous": 1440},
    "AA06": {"plain": 1200, "indigenous": 1440},
    "AA07": {"plain": 760,  "indigenous": 910},
    "AA08": {"plain": 385,  "indigenous": 460},
    "AA09": {"plain": 500,  "indigenous": 500},
    "AA11": {"plain": 500,  "indigenous": 600},
    "GA09": {"plain": 500,  "indigenous": 500},
    "BA01": {"plain": 260,  "indigenous": 310},
    "BA02": {"plain": 195,  "indigenous": 235},
    "BA03": {"plain": 35,   "indigenous": 40},
    "BA04": {"plain": 130,  "indigenous": 155},
    "BA05": {"plain": 310,  "indigenous": 370},
    "BA05-1": {"plain": 310, "indigenous": 370},
    "BA06": {"plain": 260,  "indigenous": 310},
    "BA07": {"plain": 325,  "indigenous": 390},
    "BA08": {"plain": 500,  "indigenous": 600},
    "BA09": {"plain": 155,  "indigenous": 185},
    "BA10": {"plain": 155,  "indigenous": 185},
    "BA11": {"plain": 195,  "indigenous": 235},
    "BA12": {"plain": 195,  "indigenous": 235},
    "BA13": {"plain": 195,  "indigenous": 235},
    "BA14": {"plain": 685,  "indigenous": 820},
    "BA15": {"plain": 195,  "indigenous": 235},
    "BA15-1": {"plain": 195, "indigenous": 235},
    "BA16": {"plain": 130,  "indigenous": 155},
    "BA17": {"plain": 150,  "indigenous": 180},
    "BA17-1": {"plain": 150, "indigenous": 180},
    "BA17-2": {"plain": 150, "indigenous": 180},
    "BA18": {"plain": 200,  "indigenous": 240},
    "BA19": {"plain": 150,  "indigenous": 180},
    "BA20": {"plain": 175,  "indigenous": 210},
    "BA21": {"plain": 155,  "indigenous": 185},
    "BA22": {"plain": 130,  "indigenous": 155},
    "BA23": {"plain": 200,  "indigenous": 240},
    "BA24": {"plain": 220,  "indigenous": 265},
    "GA01": {"plain": 770,  "indigenous": 925},
    "GA02": {"plain": 770,  "indigenous": 925},
    "GA03": {"plain": 770,  "indigenous": 925},
    "GA04": {"plain": 770,  "indigenous": 925},
    "GA05": {"plain": 770,  "indigenous": 925},
    "GA06": {"plain": 770,  "indigenous": 925},
    "GA07": {"plain": 770,  "indigenous": 925},
}

def make_unique_columns(raw_cols):
    """確保所有欄位名稱唯一，遇到重複名稱自動加上 _2, _3 避免 pandas 報錯"""
    seen = {}
    unique_cols = []
    for i, c in enumerate(raw_cols):
        name = str(c).strip() if pd.notna(c) and str(c).strip() != "" else f"欄位_{i}"
        if name in seen:
            seen[name] += 1
            unique_cols.append(f"{name}_{seen[name]}")
        else:
            seen[name] = 1
            unique_cols.append(name)
    return unique_cols

def detect_header_row(df_raw):
    """自動尋找包含真正欄位名稱的那一列，並保證欄位名稱不重複"""
    keywords = ["個案", "案主", "姓名", "服務人員", "居服員", "照服員", "代碼", "項目", "碼別", "次數", "數量", "日期", "金額", "小計"]
    for idx in range(min(15, len(df_raw))):
        row_values = [str(val).strip() for val in df_raw.iloc[idx].values if pd.notna(val)]
        match_count = sum(1 for kw in keywords if any(kw in val for val in row_values))
        if match_count >= 2:
            df_clean = df_raw.iloc[idx + 1:].copy()
            df_clean.columns = make_unique_columns(df_raw.iloc[idx].values)
            return df_clean.reset_index(drop=True)
    df_clean = df_raw.iloc[1:].copy()
    df_clean.columns = make_unique_columns(df_raw.iloc[0].values)
    return df_clean.reset_index(drop=True)

def read_uploaded_files(uploaded_files):
    """支援單檔或多檔自動合併讀取（含 Excel、CSV 與 PDF 表格），並自動定位真實標題列"""
    dfs = []
    for f in uploaded_files:
        f.seek(0)
        fname = f.name.lower()
        if fname.endswith('.csv'):
            df_raw = pd.read_csv(f, header=None)
            dfs.append(detect_header_row(df_raw))
        elif fname.endswith('.pdf'):
            all_rows = []
            with pdfplumber.open(f) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for tbl in tables:
                        for row in tbl:
                            cleaned_row = [str(cell).replace("\n", "").strip() if cell is not None else "" for cell in row]
                            if any(cleaned_row):
                                all_rows.append(cleaned_row)
            if all_rows:
                max_len = max(len(r) for r in all_rows)
                padded_rows = [r + [""] * (max_len - len(r)) for r in all_rows]
                df_raw = pd.DataFrame(padded_rows)
                dfs.append(detect_header_row(df_raw))
        else:
            df_raw = pd.read_excel(f, header=None)
            dfs.append(detect_header_row(df_raw))
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def auto_match_col(columns, keywords):
    """智慧自動匹配欄位名稱"""
    for kw in keywords:
        for col in columns:
            if kw in str(col):
                return col
    return columns[0] if len(columns) > 0 else None

def extract_code_prefix(val):
    """從字串中擷取標準代碼（如 'BA01 基本身體清潔' -> 'BA01'）"""
    s = str(val).strip().upper()
    m = re.match(r"^([A-Z]{2}\d{2}(?:-\d+)?)", s)
    return m.group(1) if m else s

def build_formatted_workbook(final_payroll, a_split_df, bg_df, org_name, month_label, target_workers=None):
    """100% 還原機構專屬版面：上方 A碼服務明細 (A~I) ＋ 左下薪資明細表 (A~D) ＋ 右下個案服務拆帳統計細項 (E~I)"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wb = writer.book
        if target_workers is None:
            final_payroll.to_excel(writer, sheet_name="0_全機構薪資總表", index=False)
            workers = [w for w in final_payroll["服務人員"].dropna().unique() if not str(w).startswith("⚠️")]
        else:
            workers = target_workers

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

            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1
            ws.sheet_properties.pageSetUpPr.fitToPage = True

            # =========================================================
            # 上半部：A碼服務明細 (Columns A ~ I)
            # =========================================================
            ws.merge_cells("A1:I1")
            ws["A1"] = org_name
            ws["A1"].font = f_org
            ws["A1"].alignment = al_c

            ws.merge_cells("A2:I2")
            ws["A2"] = "A碼服務明細"
            ws["A2"].font = f_sub
            ws["A2"].alignment = al_c

            ws.merge_cells("A3:B3")
            ws["A3"] = f"居服員: {worker}"
            ws["A3"].font = f_worker
            ws["A3"].alignment = al_l

            r_idx = 4
            a_cols_hdr = [(1, "服務代碼"), (2, "服務項目類別"), (3, "個案姓名"), (4, "給付價格"), (5, "數量"), (6, "小計")]
            for ci, txt in a_cols_hdr:
                c = ws.cell(row=r_idx, column=ci, value=txt)
                c.font = f_th
                c.alignment = al_c
                c.border = Border(left=thin, right=thin, top=med, bottom=med)
            ws.merge_cells(start_row=r_idx, start_column=7, end_row=r_idx, end_column=9)
            ws.cell(row=r_idx, column=7, value="服務日期").font = f_th
            ws.cell(row=r_idx, column=7).alignment = al_c
            for ci in range(7, 10):
                ws.cell(row=r_idx, column=ci).border = Border(left=thin, right=thin, top=med, bottom=med)
            ws.row_dimensions[r_idx].height = 22

            w_a_df = a_split_df[a_split_df["服務人員"] == worker].copy() if not a_split_df.empty else pd.DataFrame()
            
            if not w_a_df.empty:
                df_a_hol = w_a_df[w_a_df["服務代碼"].str.startswith(("AA09", "GA09"))].copy()
                df_a_11 = w_a_df[w_a_df["服務代碼"].str.startswith("AA11")].copy()
                df_a_reg = w_a_df[~w_a_df["服務代碼"].str.startswith(("AA09", "GA09", "AA11"))].copy()
            else:
                df_a_hol = pd.DataFrame()
                df_a_11 = pd.DataFrame()
                df_a_reg = pd.DataFrame()

            # 1) 一般 A 碼與 AA11
            df_top_block = pd.concat([df_a_reg, df_a_11], ignore_index=True)
            reg_start_row = r_idx + 1
            if not df_top_block.empty:
                for _, r in df_top_block.iterrows():
                    r_idx += 1
                    vals = [
                        r["服務代碼"],
                        r["服務項目類別"],
                        r["個案姓名"],
                        round(float(r["給付價格"]), 0),
                        round(float(r["數量"]), 2),
                        round(float(r["小計"]), 0)
                    ]
                    for ci, v in enumerate(vals, start=1):
                        c = ws.cell(row=r_idx, column=ci, value=v)
                        c.font = f_td
                        c.border = b_grid
                        c.alignment = al_l if ci == 2 else al_c
                        if ci in [4, 6]:
                            c.number_format = "#,##0"
                        elif ci == 5:
                            c.number_format = "#,##0.##"
                    ws.merge_cells(start_row=r_idx, start_column=7, end_row=r_idx, end_column=9)
                    dc = ws.cell(row=r_idx, column=7, value=r["服務日期"])
                    dc.font = f_td
                    dc.alignment = al_c
                    for ci in range(7, 10):
                        ws.cell(row=r_idx, column=ci).border = b_grid
                    ws.row_dimensions[r_idx].height = 24
            else:
                r_idx += 1
                for ci in range(1, 7):
                    c = ws.cell(row=r_idx, column=ci, value="-" if ci == 1 else "")
                    c.font = f_td
                    c.border = b_grid
                    c.alignment = al_c
                ws.merge_cells(start_row=r_idx, start_column=7, end_row=r_idx, end_column=9)
                for ci in range(7, 10):
                    ws.cell(row=r_idx, column=ci).border = b_grid

            reg_end_row = r_idx
            r_idx += 1
            reg_tot_row = r_idx
            for ci in range(1, 7):
                c = ws.cell(row=r_idx, column=ci)
                c.border = b_grid
                if ci == 6:
                    c.value = f'=SUMIFS(F{reg_start_row}:F{reg_end_row}, A{reg_start_row}:A{reg_end_row}, "<>AA11*")'
                    c.font = f_red_tot
                    c.fill = fill_pink
                    c.alignment = al_c
                    c.number_format = "#,##0"
            ws.merge_cells(start_row=r_idx, start_column=7, end_row=r_idx, end_column=9)
            for ci in range(7, 10):
                c = ws.cell(row=r_idx, column=ci)
                c.fill = fill_pink
                c.border = b_grid
            ws.row_dimensions[r_idx].height = 22

            # 2) 假日碼 (AA09 / GA09) 明細列
            hol_start_row = r_idx + 1
            if not df_a_hol.empty:
                for _, r in df_a_hol.iterrows():
                    r_idx += 1
                    vals = [
                        r["服務代碼"],
                        r["服務項目類別"],
                        r["個案姓名"],
                        round(float(r["給付價格"]), 0),
                        round(float(r["數量"]), 2),
                        round(float(r["小計"]), 0)
                    ]
                    for ci, v in enumerate(vals, start=1):
                        c = ws.cell(row=r_idx, column=ci, value=v)
                        c.font = f_td
                        c.border = b_grid
                        c.alignment = al_l if ci == 2 else al_c
                        if ci in [4, 6]:
                            c.number_format = "#,##0"
                        elif ci == 5:
                            c.number_format = "#,##0.##"
                    ws.merge_cells(start_row=r_idx, start_column=7, end_row=r_idx, end_column=9)
                    dc = ws.cell(row=r_idx, column=7, value=r["服務日期"])
                    dc.font = f_td
                    dc.alignment = al_c
                    for ci in range(7, 10):
                        ws.cell(row=r_idx, column=ci).border = b_grid
                    ws.row_dimensions[r_idx].height = 24
            else:
                r_idx += 1
                for ci in range(1, 7):
                    c = ws.cell(row=r_idx, column=ci, value="-" if ci == 1 else "")
                    c.font = f_td
                    c.border = b_grid
                    c.alignment = al_c
                ws.merge_cells(start_row=r_idx, start_column=7, end_row=r_idx, end_column=9)
                for ci in range(7, 10):
                    ws.cell(row=r_idx, column=ci).border = b_grid

            hol_end_row = r_idx
            r_idx += 1
            hol_tot_row = r_idx
            for ci in range(1, 10):
                c = ws.cell(row=r_idx, column=ci)
                c.border = Border(left=thin, right=thin, top=thin, bottom=med)
                if ci == 6:
                    c.value = f"=SUM(F{hol_start_row}:F{hol_end_row})"
                    c.font = f_red_tot
                    c.fill = fill_pink
                    c.alignment = al_c
                    c.number_format = "#,##0"
                elif ci in [7, 8]:
                    c.fill = fill_pink
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
            m_display = f"{month_label[:3]}年{month_label[3:]}月薪資明細表" if len(month_label) == 5 and month_label.isdigit() else f"{month_label} 薪資明細表"
            ws.cell(row=r_idx, column=1, value=m_display).font = f_th
            ws.cell(row=r_idx, column=1).alignment = al_c

            r_idx += 1
            name_row = r_idx
            ws.cell(row=name_row, column=1, value="姓名:").font = f_td_bold
            ws.cell(row=name_row, column=1).alignment = al_l
            ws.cell(row=name_row, column=2, value=worker).font = f_sub
            ws.cell(row=name_row, column=2).alignment = al_c

            # =========================================================
            # 下半部右側：個案服務拆帳統計細項 (Columns E ~ I)
            # =========================================================
            right_start_row = name_row
            ws.merge_cells(start_row=right_start_row, start_column=5, end_row=right_start_row, end_column=9)
            rc_title = ws.cell(row=right_start_row, column=5, value="個案服務拆帳統計細項")
            rc_title.font = f_right_title
            rc_title.fill = fill_gray_hdr
            rc_title.alignment = al_c
            for ci in range(5, 10):
                ws.cell(row=right_start_row, column=ci).border = b_right_grid

            left_hdr_row = name_row + 1
            for ci, txt in [(1, ""), (2, "項目"), (3, "金額"), (4, "備註")]:
                c = ws.cell(row=left_hdr_row, column=ci, value=txt)
                c.font = f_th
                c.alignment = al_c
                c.border = Border(left=thin, right=thin, top=med, bottom=thin)

            right_headers = [(5, ""), (6, "個案姓名"), (7, "服務代碼"), (8, "服務次數"), (9, "金額")]
            for ci, txt in right_headers:
                c = ws.cell(row=left_hdr_row, column=ci, value=txt)
                c.font = f_th
                c.fill = fill_gray_sub
                c.alignment = al_c
                c.border = b_right_grid

            w_bg_df = bg_df[bg_df["服務人員"] == worker].copy() if not bg_df.empty else pd.DataFrame()
            if not w_bg_df.empty:
                client_bg_summary = (
                    w_bg_df.groupby("個案姓名", as_index=False, sort=False)
                    .agg({
                        "服務代碼": lambda s: ",".join(dict.fromkeys([str(x) for x in s if str(x)])),
                        "服務次數": "sum",
                        "居服員實領": "sum"
                    })
                )
            else:
                client_bg_summary = pd.DataFrame(columns=["個案姓名", "服務代碼", "服務次數", "居服員實領"])

            r_data_start = left_hdr_row + 1
            max_rows = max(20, len(client_bg_summary))
            for i in range(max_rows):
                curr_r = r_data_start + i
                c_idx_cell = ws.cell(row=curr_r, column=5, value=i + 1)
                c_idx_cell.font = f_td
                c_idx_cell.alignment = al_c
                c_idx_cell.border = b_right_grid

                if i < len(client_bg_summary):
                    c_name_val = client_bg_summary.iloc[i]["個案姓名"]
                    c_code_val = client_bg_summary.iloc[i]["服務代碼"]
                    c_cnt_val = round(float(client_bg_summary.iloc[i]["服務次數"]), 1)
                    c_amt_val = round(float(client_bg_summary.iloc[i]["居服員實領"]), 0)
                else:
                    c_name_val = ""
                    c_code_val = ""
                    c_cnt_val = ""
                    c_amt_val = ""

                c1 = ws.cell(row=curr_r, column=6, value=c_name_val)
                c1.font = f_td
                c1.alignment = al_c
                c1.border = b_right_grid

                c2 = ws.cell(row=curr_r, column=7, value=c_code_val)
                c2.font = f_td
                c2.alignment = al_c
                c2.border = b_right_grid

                c3 = ws.cell(row=curr_r, column=8, value=c_cnt_val)
                c3.font = f_td
                c3.alignment = al_c
                if c_cnt_val != "":
                    c3.number_format = "#,##0.#"
                c3.border = b_right_grid

                c4 = ws.cell(row=curr_r, column=9, value=c_amt_val)
                c4.font = f_td
                c4.alignment = al_r
                c4.number_format = '#,##0" 元"' if c_amt_val != "" else '@" 元"'
                c4.border = b_right_grid

            r_data_end = r_data_start + max_rows - 1
            r_sub_row = r_data_end + 1
            ws.merge_cells(start_row=r_sub_row, start_column=5, end_row=r_sub_row, end_column=6)
            ws.cell(row=r_sub_row, column=5, value="薪資小計 (A) :").font = f_th
            ws.cell(row=r_sub_row, column=5).alignment = al_r
            sub_val_cell = ws.cell(row=r_sub_row, column=9, value=f"=SUM(I{r_data_start}:I{r_data_end})")
            sub_val_cell.font = f_td_bold
            sub_val_cell.alignment = al_r
            sub_val_cell.number_format = '#,##0" 元"'
            for ci in range(5, 10):
                ws.cell(row=r_sub_row, column=ci).fill = fill_gray_hdr
                ws.cell(row=r_sub_row, column=ci).border = b_right_grid

            # 左下角「應領金額」與「應扣金額」明細表
            earn_start = left_hdr_row + 1
            earn_items = [
                ("工作薪資", f"=I{r_sub_row}"),
                ("偏鄉交通補助", 0),
                ("偏鄉獎勵津貼", 0),
                ("交通費", 0),
                ("轉場津貼", 0),
                ("AA05.06.07", f"=F{reg_tot_row}"),
                ("A碼假日(AA09)", f"=F{hol_tot_row}"),
                ("AA11 (身障/失智)", f'=SUMIFS(F{reg_start_row}:F{reg_end_row}, A{reg_start_row}:A{reg_end_row}, "AA11*")'),
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
                vc.number_format = '#,##0;-#,##0;"-"'
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

            ws.merge_cells(start_row=earn_start, start_column=1, end_row=earn_tot_row, end_column=1)
            ws.cell(row=earn_start, column=1, value="應領金額").font = f_th
            ws.cell(row=earn_start, column=1).alignment = al_c

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

            col_widths = {1: 11, 2: 24, 3: 13, 4: 11, 5: 7, 6: 13, 7: 14, 8: 11, 9: 15}
            for col_i, width in col_widths.items():
                ws.column_dimensions[get_column_letter(col_i)].width = width

        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
            del wb["Sheet"]

    return output.getvalue()

def render_home_care_payroll():
    st.title("💰 居家服務薪資拆帳與核銷對帳引擎")
    st.caption("支援【1.A碼核銷表 ＋ 2.服務紀錄明細表(BA/G碼)】雙檔精準運算，100% 還原機構專屬雙欄薪資明細表。")

    with st.expander("⚙️ 機構抬頭、薪資年月與拆帳內規設定", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("**1️⃣ 機構名稱與薪資年月**")
            org_name = st.text_input("機構全銜", value="南門長照社團法人附設屏東縣私立南門綜合長照機構")
            month_label = st.text_input("薪資年月代碼（如 11506）", value="11506")
        with col2:
            st.markdown("**2️⃣ A碼給薪與定額津貼**")
            a_share_pct = st.slider("一般 A碼 居服員分潤 (%)", 0, 100, 50, step=5, help="預設 55 拆帳 (50%)；AA04、AA11 自動全額 100% 給薪")
            custom_aa09 = st.number_input("AA09 / GA09 定額實領 (元/日)", value=500, step=50)
        with col3:
            st.markdown("**3️⃣ B碼 / G碼 分潤與進位**")
            bg_share_pct = st.slider("BA碼 / G碼 居服員分潤 (%)", 0, 100, 60, step=5, help="預設採平地單價 64 拆帳 (60%)")
            round_to_int = st.checkbox("金額四捨五入至整數", value=True)

    price_map = DEFAULT_PRICE_TABLE.copy()
    price_map["AA09"] = {"plain": custom_aa09, "indigenous": custom_aa09}
    price_map["GA09"] = {"plain": custom_aa09, "indigenous": custom_aa09}

    st.divider()

    # ==========================================
    # 雙檔上傳區（已移除多餘的橫向班表欄位）
    # ==========================================
    u_col1, u_col2 = st.columns(2)
    with u_col1:
        a_claim_files = st.file_uploader(
            "1️⃣ 上傳【A碼核銷表】(僅讀代碼/類別/個案/數量/日期/人員)",
            type=["xlsx", "xls", "csv", "pdf"],
            accept_multiple_files=True
        )
    with u_col2:
        service_detail_files = st.file_uploader(
            "2️⃣ 上傳【服務紀錄明細表】(算BA/G碼平地薪資 ＋ 自動統計A碼次數比例)",
            type=["xlsx", "xls", "csv", "pdf"],
            accept_multiple_files=True
        )

    if not a_claim_files or not service_detail_files:
        st.info("👆 請上傳上方【1.A碼核銷表】與【2.服務紀錄明細表】兩個檔案（支援 Excel、CSV、PDF）即可開始運算。")
        return

    df_a_raw = read_uploaded_files(a_claim_files)
    df_svc_raw = read_uploaded_files(service_detail_files)

    # ==========================================
    # 雙表欄位對齊確認
    # ==========================================
    st.subheader("🔗 雙表欄位對齊確認")
    with st.expander("展開/收合欄位對應細節", expanded=True):
        st.markdown("**1️⃣ 【A碼核銷表】欄位（強制僅讀取指定 6 欄位）：**")
        ac_cols = list(df_a_raw.columns)
        a1, a2, a3, a4, a5, a6 = st.columns(6)
        with a1:
            ac_code = st.selectbox("A碼-服務代碼", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["服務代碼", "代碼"])))
        with a2:
            ac_name = st.selectbox("A碼-服務項目類別", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["服務項目類別", "項目類別", "項目名稱", "類別"])))
        with a3:
            ac_client = st.selectbox("A碼-個案姓名", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["個案姓名", "個案", "姓名"])))
        with a4:
            ac_qty = st.selectbox("A碼-數量", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["數量", "次數", "天數"])))
        with a5:
            ac_date = st.selectbox("A碼-服務日期", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["服務日期", "日期"])))
        with a6:
            ac_worker = st.selectbox("A碼-服務人員", ac_cols, index=ac_cols.index(auto_match_col(ac_cols, ["服務人員", "居服員", "人員"])))

        st.markdown("**2️⃣ 【服務紀錄明細表】欄位（只讀取 BA碼 與 G碼 採平地計薪，並同步統計次數比例）：**")
        sv_cols = list(df_svc_raw.columns)
        v1, v2, v3, v4, v5 = st.columns(5)
        with v1:
            sv_client = st.selectbox("明細-個案姓名", sv_cols, index=sv_cols.index(auto_match_col(sv_cols, ["個案姓名", "個案", "姓名"])))
        with v2:
            sv_worker = st.selectbox("明細-服務人員", sv_cols, index=sv_cols.index(auto_match_col(sv_cols, ["服務人員姓名", "服務人員", "居服員", "照服員", "人員"])))
        with v3:
            sv_code = st.selectbox("明細-服務代碼", sv_cols, index=sv_cols.index(auto_match_col(sv_cols, ["服務項目代碼", "服務代碼", "照顧組合", "項目代碼", "代碼"])))
        with v4:
            sv_cnt_opts = ["(每列算 1 次)"] + sv_cols
            def_sv_cnt = auto_match_col(sv_cols, ["數量(僅整數)", "數量", "次數", "總次"])
            def_sv_cnt_idx = sv_cnt_opts.index(def_sv_cnt) if def_sv_cnt in sv_cols and any(k in str(def_sv_cnt) for k in ["次數", "數量"]) else 0
            sv_count = st.selectbox("明細-次數/數量", sv_cnt_opts, index=def_sv_cnt_idx)
        with v5:
            sv_fee_opts = ["(依代碼查平地單價表)"] + sv_cols
            exact_plain = next((c for c in sv_cols if str(c).strip() == "申請(補助)費用"), None)
            def_sv_fee = exact_plain if exact_plain else auto_match_col(sv_cols, ["申請(補助)費用", "平地", "金額", "小計"])
            def_sv_fee_idx = sv_fee_opts.index(def_sv_fee) if def_sv_fee in sv_cols and any(k in str(def_sv_fee) for k in ["申請(補助)費用", "平地", "金額", "小計"]) else 0
            sv_fee = st.selectbox("明細-平地費用欄(選填)", sv_fee_opts, index=def_sv_fee_idx)

    # ==========================================
    # 核心計算引擎
    # ==========================================
    if st.button("🚀 執行精準拆帳並產生薪資明細表", type="primary", use_container_width=True):
        # ------------------------------------------
        # Step 1：整理【服務紀錄明細表】
        # ------------------------------------------
        svc_clean = pd.DataFrame({
            "個案姓名": df_svc_raw[sv_client],
            "服務人員": df_svc_raw[sv_worker],
            "服務代碼": df_svc_raw[sv_code],
            "服務次數": 1.0 if sv_count == "(每列算 1 次)" else pd.to_numeric(df_svc_raw[sv_count], errors='coerce').fillna(1.0),
        })
        if sv_fee != "(依代碼查平地單價表)":
            svc_clean["平地總額"] = pd.to_numeric(df_svc_raw[sv_fee], errors='coerce').fillna(0.0)
        else:
            svc_clean["平地總額"] = np.nan

        svc_clean = svc_clean[svc_clean["個案姓名"].notna() & svc_clean["服務人員"].notna() & svc_clean["服務代碼"].notna()].copy()
        svc_clean = svc_clean[~svc_clean["個案姓名"].astype(str).str.contains("總計|合計|小計", na=False)].copy()
        svc_clean["個案姓名"] = svc_clean["個案姓名"].astype(str).str.strip()
        svc_clean["服務人員"] = svc_clean["服務人員"].astype(str).str.strip()
        svc_clean["服務代碼"] = svc_clean["服務代碼"].apply(extract_code_prefix)

        svc_clean["_w_list"] = svc_clean["服務人員"].apply(lambda x: [w.strip() for w in re.split(r"[、,，/]+", str(x)) if w.strip()])
        svc_clean["_n"] = svc_clean["_w_list"].apply(lambda l: max(len(l), 1))
        svc_clean["服務次數"] = svc_clean["服務次數"] / svc_clean["_n"]
        if sv_fee != "(依代碼查平地單價表)":
            svc_clean["平地總額"] = svc_clean["平地總額"] / svc_clean["_n"]
        svc_clean = svc_clean.explode("_w_list").reset_index(drop=True)
        svc_clean["服務人員"] = svc_clean["_w_list"]

        worker_client_cnt = svc_clean.groupby(["個案姓名", "服務人員"])["服務次數"].sum().reset_index().rename(columns={"服務次數": "個人服務次數"})
        client_tot_cnt = worker_client_cnt.groupby("個案姓名")["個人服務次數"].sum().reset_index().rename(columns={"個人服務次數": "該案總服務次數"})
        ratio_df = pd.merge(worker_client_cnt, client_tot_cnt, on="個案姓名")
        ratio_df["次數比例"] = ratio_df["個人服務次數"] / ratio_df["該案總服務次數"]

        bg_df = svc_clean[svc_clean["服務代碼"].str.startswith(("BA", "G"))].copy()
        if sv_fee == "(依代碼查平地單價表)":
            def get_plain_price(c):
                if c in price_map:
                    return price_map[c]["plain"]
                base_c = c.split("-")[0]
                return price_map.get(base_c, {"plain": 0})["plain"]

            bg_df["平地單價"] = bg_df["服務代碼"].apply(get_plain_price)
            bg_df["平地總額"] = bg_df["平地單價"] * bg_df["服務次數"]

        bg_df["居服員實領"] = bg_df["平地總額"] * (bg_share_pct / 100.0)

        # ------------------------------------------
        # Step 2：計算【A碼核銷表】
        # ------------------------------------------
        a_df = pd.DataFrame({
            "服務代碼": df_a_raw[ac_code],
            "服務項目類別": df_a_raw[ac_name],
            "個案姓名": df_a_raw[ac_client],
            "數量": df_a_raw[ac_qty],
            "服務日期": df_a_raw[ac_date],
            "服務人員": df_a_raw[ac_worker],
        })
        a_df = a_df[a_df["個案姓名"].notna() & a_df["服務代碼"].notna()].copy()
        a_df = a_df[~a_df["個案姓名"].astype(str).str.contains("總計|合計|小計", na=False)].copy()
        a_df["服務代碼"] = a_df["服務代碼"].apply(extract_code_prefix)
        a_df["服務項目類別"] = a_df["服務項目類別"].fillna("").astype(str).str.strip()
        a_df["個案姓名"] = a_df["個案姓名"].astype(str).str.strip()
        a_df["數量"] = pd.to_numeric(a_df["數量"], errors='coerce').fillna(1.0)
        a_df["服務日期"] = a_df["服務日期"].fillna("").astype(str).str.strip()
        a_df["服務人員"] = a_df["服務人員"].fillna("").astype(str).str.strip()
        a_df.loc[a_df["服務人員"].isin(["nan", "None", ""]), "服務人員"] = ""

        expanded_a_rows = []
        for idx_row, r in a_df.iterrows():
            code = r["服務代碼"]
            cname = r["個案姓名"]
            cat_name = r["服務項目類別"]
            raw_qty = float(r["數量"])
            raw_dates = [d.strip() for d in re.split(r"[,，、\s]+", r["服務日期"]) if d.strip() and d.strip() != "nan"]
            raw_workers = [w.strip() for w in re.split(r"[、,，/]+", r["服務人員"]) if w.strip()]

            unit_p = price_map.get(code, {"plain": 0})["plain"]
            if code.startswith(("AA09", "GA09")):
                unit_p = custom_aa09
                pay_rate = 1.0
            elif code.startswith(("AA04", "AA11")):
                pay_rate = 1.0
            else:
                pay_rate = a_share_pct / 100.0

            if len(raw_workers) > 1:
                n_w = len(raw_workers)
                for w in raw_workers:
                    expanded_a_rows.append({
                        "服務代碼": code,
                        "服務項目類別": cat_name,
                        "個案姓名": cname,
                        "給付價格": unit_p,
                        "數量": raw_qty / n_w,
                        "小計": (unit_p * raw_qty * pay_rate) / n_w,
                        "服務日期": r["服務日期"],
                        "服務人員": w,
                        "_date_key": r["服務日期"] if r["服務日期"] else f"row_{idx_row}"
                    })
            elif len(raw_workers) == 1:
                w = raw_workers[0]
                if len(raw_dates) > 1 and abs(raw_qty - len(raw_dates)) < 0.01:
                    for d in raw_dates:
                        expanded_a_rows.append({
                            "服務代碼": code,
                            "服務項目類別": cat_name,
                            "個案姓名": cname,
                            "給付價格": unit_p,
                            "數量": 1.0,
                            "小計": unit_p * 1.0 * pay_rate,
                            "服務日期": d,
                            "服務人員": w,
                            "_date_key": d
                        })
                else:
                    expanded_a_rows.append({
                        "服務代碼": code,
                        "服務項目類別": cat_name,
                        "個案姓名": cname,
                        "給付價格": unit_p,
                        "數量": raw_qty,
                        "小計": unit_p * raw_qty * pay_rate,
                        "服務日期": r["服務日期"],
                        "服務人員": w,
                        "_date_key": r["服務日期"] if r["服務日期"] else f"row_{idx_row}"
                    })
            else:
                matched_workers = ratio_df[ratio_df["個案姓名"] == cname]
                if not matched_workers.empty:
                    for _, mw in matched_workers.iterrows():
                        prop = float(mw["次數比例"])
                        expanded_a_rows.append({
                            "服務代碼": code,
                            "服務項目類別": cat_name,
                            "個案姓名": cname,
                            "給付價格": unit_p,
                            "數量": raw_qty * prop,
                            "小計": unit_p * raw_qty * pay_rate * prop,
                            "服務日期": r["服務日期"],
                            "服務人員": mw["服務人員"],
                            "_date_key": f"monthly_{idx_row}"
                        })
                else:
                    expanded_a_rows.append({
                        "服務代碼": code,
                        "服務項目類別": cat_name,
                        "個案姓名": cname,
                        "給付價格": unit_p,
                        "數量": raw_qty,
                        "小計": unit_p * raw_qty * pay_rate,
                        "服務日期": r["服務日期"],
                        "服務人員": "⚠️ 無明細紀錄(需手動指定)",
                        "_date_key": f"unmatched_{idx_row}"
                    })

        exp_a_df = pd.DataFrame(expanded_a_rows)

        if not exp_a_df.empty:
            valid_date_mask = ~exp_a_df["_date_key"].astype(str).str.startswith(("row_", "monthly_", "unmatched_"))
            same_day_counts = (
                exp_a_df[valid_date_mask]
                .groupby(["個案姓名", "服務代碼", "_date_key"])["服務人員"]
                .transform("nunique")
            )
            exp_a_df.loc[valid_date_mask, "_same_day_heads"] = same_day_counts
            exp_a_df["_same_day_heads"] = exp_a_df["_same_day_heads"].fillna(1.0)
            exp_a_df["數量"] = exp_a_df["數量"] / exp_a_df["_same_day_heads"]
            exp_a_df["小計"] = exp_a_df["小計"] / exp_a_df["_same_day_heads"]

            def join_dates(s):
                vals = [str(x).strip() for x in s if str(x).strip() and str(x).strip() != "nan"]
                return ",".join(dict.fromkeys(vals))

            a_split_df = (
                exp_a_df.groupby(["服務人員", "服務代碼", "服務項目類別", "個案姓名", "給付價格"], as_index=False)
                .agg({
                    "數量": "sum",
                    "小計": "sum",
                    "服務日期": join_dates
                })
            )
            if round_to_int:
                a_split_df["小計"] = a_split_df["小計"].round(0)
        else:
            a_split_df = pd.DataFrame(columns=["服務人員", "服務代碼", "服務項目類別", "個案姓名", "給付價格", "數量", "小計", "服務日期"])

        # ------------------------------------------
        # Step 3：彙總全機構總表（零尾差）
        # ------------------------------------------
        bg_sum = bg_df.groupby("服務人員")[["平地總額", "居服員實領"]].sum().reset_index().rename(
            columns={"平地總額": "BA/G碼平地總額", "居服員實領": "工作薪資(BA+G碼)"}
        ) if not bg_df.empty else pd.DataFrame(columns=["服務人員", "BA/G碼平地總額", "工作薪資(BA+G碼)"])
        if round_to_int and not bg_sum.empty:
            bg_sum["工作薪資(BA+G碼)"] = bg_sum["工作薪資(BA+G碼)"].round(0)

        a_sum = a_split_df.groupby("服務人員")["小計"].sum().reset_index().rename(
            columns={"小計": "A碼實領總額"}
        ) if not a_split_df.empty else pd.DataFrame(columns=["服務人員", "A碼實領總額"])

        final_payroll = pd.merge(bg_sum, a_sum, on="服務人員", how="outer").fillna(0)
        final_payroll["本月應發小計(A+BA/G)"] = final_payroll["工作薪資(BA+G碼)"] + final_payroll["A碼實領總額"]

        all_excel_bytes = build_formatted_workbook(
            final_payroll, a_split_df, bg_df, org_name, month_label, target_workers=None
        )

        st.session_state["payroll_result"] = {
            "final_payroll": final_payroll,
            "a_split_df": a_split_df,
            "bg_df": bg_df,
            "all_excel_bytes": all_excel_bytes,
            "org_name": org_name,
            "month_label": month_label
        }

    # ==========================================
    # 結果展示與單人/整批下載
    # ==========================================
    if "payroll_result" in st.session_state:
        res = st.session_state["payroll_result"]
        final_payroll = res["final_payroll"]
        a_split_df = res["a_split_df"]
        bg_df = res["bg_df"]
        m_label = res["month_label"]

        valid_workers = [w for w in final_payroll["服務人員"].tolist() if not str(w).startswith("⚠️")]
        st.success(f"✅ 雙表精準核算完成！共為 {len(valid_workers)} 位居服員產出專屬薪資明細表：")

        m1, m2, m3 = st.columns(3)
        m1.metric("全機構應發小計 (A+BA/G)", f"${final_payroll['本月應發小計(A+BA/G)'].sum():,.0f}")
        m2.metric("全機構 A碼實領總額", f"${final_payroll['A碼實領總額'].sum():,.0f}")
        m3.metric("全機構 BA/G碼工作薪資總額", f"${final_payroll['工作薪資(BA+G碼)'].sum():,.0f}")

        st.divider()

        preview_list = valid_workers + [w for w in final_payroll["服務人員"].tolist() if str(w).startswith("⚠️")]
        preview_worker = st.selectbox("🔍 選擇居服員檢視個人明細與下載單人薪資表：", preview_list)

        w_row = final_payroll[final_payroll["服務人員"] == preview_worker].iloc[0]
        wm1, wm2, wm3 = st.columns(3)
        wm1.metric(f"{preview_worker} 應發薪資小計", f"${w_row['本月應發小計(A+BA/G)']:,.0f}")
        wm2.metric("A碼實領總額", f"${w_row['A碼實領總額']:,.0f}")
        wm3.metric("工作薪資小計 (BA+G碼)", f"${w_row['工作薪資(BA+G碼)']:,.0f}")

        p_col1, p_col2 = st.columns(2)
        with p_col1:
            st.markdown(f"**{preview_worker} — A碼服務明細**")
            w_a_view = a_split_df[a_split_df["服務人員"] == preview_worker][["服務代碼", "服務項目類別", "個案姓名", "給付價格", "數量", "小計", "服務日期"]]
            st.dataframe(w_a_view, use_container_width=True, hide_index=True)
        with p_col2:
            st.markdown(f"**{preview_worker} — 個案服務拆帳統計細項 (BA/G碼)**")
            if not bg_df.empty:
                w_bg_view = bg_df[bg_df["服務人員"] == preview_worker].groupby("個案姓名", as_index=False, sort=False).agg({
                    "服務代碼": lambda s: ",".join(dict.fromkeys([str(x) for x in s if str(x)])),
                    "服務次數": "sum",
                    "居服員實領": lambda x: round(sum(x), 0)
                }).rename(columns={"居服員實領": "金額(元)"})
                st.dataframe(w_bg_view, use_container_width=True, hide_index=True)

        single_excel_bytes = build_formatted_workbook(
            final_payroll, a_split_df, bg_df, res["org_name"], m_label, target_workers=[preview_worker]
        )

        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                label=f"🖨️ 下載【{m_label} {preview_worker} 薪資明細表.xlsx】（完全還原機構版面）",
                data=single_excel_bytes,
                file_name=f"{m_label}_{preview_worker}_薪資明細表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with dl_col2:
            st.download_button(
                label=f"📥 下載【{m_label} 全機構居服員薪資明細表.xlsx】（一人一分頁）",
                data=res["all_excel_bytes"],
                file_name=f"{m_label}_全機構居服員薪資明細表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )