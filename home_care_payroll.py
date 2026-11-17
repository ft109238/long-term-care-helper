import streamlit as st
import pandas as pd
import numpy as np
import io
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

def build_formatted_workbook(final_payroll, a_split_df, b_df, col_client, col_worker, col_code, col_count, month_label, a_share_pct, b_share_pct):
    """產生含【全機構總表】＋【每位居服員專屬個人薪資單分頁】的精美格式 Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1. 先寫入總表
        final_payroll.to_excel(writer, sheet_name="0_全機構薪資總表", index=False)
        wb = writer.book

        # 樣式定義
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

        # 2. 依每位居服員建立專屬薪資明細分頁
        workers = [w for w in final_payroll[col_worker].dropna().unique() if not str(w).startswith("⚠️")]

        for worker in workers:
            # Excel 分頁名稱最長 31 字且不能含特殊字元
            safe_sheet_name = str(worker)[:30].replace("/", "_").replace("\\", "_").replace("*", "").replace("?", "").replace("[", "").replace("]", "")
            ws = wb.create_sheet(title=safe_sheet_name)

            # 第 1 列：主標題（如：11506 張敏 薪資服務明細）
            ws.merge_cells("A1:G1")
            ws["A1"] = f"{month_label} {worker} 薪資服務明細"
            ws["A1"].font = font_title
            ws["A1"].alignment = align_center
            ws.row_dimensions[1].height = 28

            # ------------------------------------------
            # 上半部：A 碼薪資明細表
            # ------------------------------------------
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
                        int(r["該案總服務次數"]) if pd.notna(r.get("該案總服務次數")) else 0,
                        int(r["個人服務次數"]) if pd.notna(r.get("個人服務次數")) else 0,
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
                            cell.number_format = "#,##0"
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

            # A碼總計列
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

            # ------------------------------------------
            # 下半部：B 碼服務薪資明細表（採平地計薪基準）
            # ------------------------------------------
            row_idx += 2
            ws.cell(row=row_idx, column=1, value="【二、B碼照顧服務費明細（依公司規定採平地標準拆帳）】").font = font_sec

            row_idx += 1
            b_headers = [
                "個案姓名",
                "服務項目類別",
                "服務次數",
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
                # 依「個案姓名 + 服務項目類別」彙整該居服員在每個個案上的明細，版面最乾淨
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
                        int(r[col_count]),
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
                            cell.number_format = "#,##0"
                        else:
                            cell.alignment = align_right
                            cell.number_format = "#,##0"
            else:
                row_idx += 1
                ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=5)
                ws.cell(row=row_idx, column=1, value="本月無 B 碼服務紀錄").alignment = align_center
                for c_i in range(1, 6):
                    ws.cell(row=row_idx, column=c_i).border = border_box

            # B碼總計列
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

            # 本月實領總薪資 (A + B)
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