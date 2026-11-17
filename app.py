import streamlit as st
from home_care_payroll import render_home_care_payroll
import io
import pandas as pd
import time
import re
from datetime import datetime
from openai import OpenAI

# ──────────────────────────────────────────────
# 1. 頁面設定與通用標題
# ──────────────────────────────────────────────
st.set_page_config(
    page_title=" AI 智慧行政工作台",
    page_icon="📋",
    layout="wide"
)

# ──────────────────────────────────────────────
# 2. 授權名單資料庫（角色權限嚴格控管）
# ──────────────────────────────────────────────
AUTHORIZED_USERS = {
    "CROCS-537": {
            "name": "管理者",
            "expire": "2035-12-31",
            "roles": ["機構社工", "居家督導", "個管師", "照顧專員", "居家行政(薪資拆帳)"]  # 全功能開通
    },
    "SHASHA-VIP": {
        "name": "種子顧問-莎莎",
        "expire": "2027-12-31",
        "roles": ["機構社工", "居家督導", "個管師", "照顧專員"]  # 全功能開通
    },
    "TEST-2026Q4": {
        "name": "2026 Q4 早鳥試用戶",
        "expire": "2026-12-31",
        "roles": ["機構社工"]
    },
    "TEST-202612": {
        "name": "202612 早鳥試用戶",
        "expire": "2026-12-31",
        "roles": ["居家督導", "居家行政(薪資拆帳)"]
    },
    "TEST-202600": {
        "name": "202600 早鳥試用戶",
        "expire": "2026-12-31",
        "roles": ["個管師"]
    },
    "TEST-202611": {
        "name": "202611 早鳥試用戶",
        "expire": "2026-12-31",
        "roles": ["照顧專員"]
    },
    "AMMAWU-VIP": {
        "name": "欣蓉居家督導",
        "expire": "2027-12-31",
        "roles": ["居家督導"]
    }
}

# 初始化 Session State
if "generated_df" not in st.session_state:
    st.session_state.generated_df = None
if "current_quarter" not in st.session_state:
    st.session_state.current_quarter = "2026年 第一季"

# ──────────────────────────────────────────────
# 3. 側邊欄：授權驗證與心理防禦
# ──────────────────────────────────────────────
st.sidebar.header("🔑 系統授權驗證")
user_code = st.sidebar.text_input("請輸入授權通行碼", type="password").strip()

if not user_code:
    st.sidebar.info("💡 請輸入通行碼解鎖專屬工作台。")
    st.stop()

if user_code not in AUTHORIZED_USERS:
    st.sidebar.error("❌ 無效的通行碼！請確認輸入或聯繫管理員。")
    st.stop()

# 取得客戶授權資料
user_info = AUTHORIZED_USERS[user_code]
expire_date = datetime.strptime(user_info["expire"], "%Y-%m-%d")
days_left = (expire_date - datetime.now()).days

# 4. 到期狀態阻斷與分級提醒
if days_left < 0:
    st.error(f"🚫 您的授權已於 **{user_info['expire']}** 到期，所有生成功能已停用。續約請洽客服。")
    st.sidebar.error(f"❌ 授權已於 {user_info['expire']} 到期")
    st.stop()

# 驗證成功心理防禦展示
st.sidebar.success(f"授權單位：{user_info['name']}")
st.sidebar.caption(f"📅 有效期限至：{user_info['expire']}（剩餘 {days_left} 天）")
st.sidebar.markdown(
    """
    <div style='background-color: #2b1d0c; padding: 8px; border-radius: 5px; border-left: 3px solid #ff9800; font-size: 11px; color: #ffb74d;'>
    ⚠️ <b>資安監控啟動</b><br>
    本帳戶已綁定專屬機構使用權限。後台具備即時異常流量偵測機制，嚴禁跨機構借調或轉售通行碼。
    </div>
    """,
    unsafe_allow_html=True
)

# 到期前 30 天彈窗提示
@st.dialog("⏳ 年度授權即將到期提醒")
def show_expiration_dialog(d_left, exp_date):
    st.warning(f"敬愛的客戶您好：\n\n您的年度工作台授權尚餘 **{d_left} 天**（到期日：`{exp_date}`）。")
    st.write("為避免季末評鑑資料處理中斷，建議提早聯繫您的專屬客服或發起續約申請。")
    if st.button("我知道了，進入工作台", use_container_width=True):
        st.session_state["seen_warning"] = True
        st.rerun()

if 0 <= days_left <= 30 and "seen_warning" not in st.session_state:
    show_expiration_dialog(days_left, user_info['expire'])
elif 30 < days_left <= 90:
    st.sidebar.warning(f"🔔 續約提醒：授權尚餘 {days_left} 天，請留意續約排程。")

# ──────────────────────────────────────────────
# 5. 側邊欄：角色權限切換
# ──────────────────────────────────────────────
st.sidebar.write("---")
st.sidebar.subheader("🎯 業務模組選擇")
allowed_roles = user_info.get("roles", ["機構社工"])

if len(allowed_roles) == 1:
    selected_role = allowed_roles[0]
    st.sidebar.info(f"當前模式：**{selected_role}**")
else:
    selected_role = st.sidebar.selectbox("切換工作模組", allowed_roles)

# ──────────────────────────────────────────────
# 6. 側邊欄常駐資訊（教學、方案、條款）
# ──────────────────────────────────────────────
st.sidebar.write("---")
st.sidebar.subheader("📚 支援與合規中心")

with st.sidebar.expander("📖 A | 操作教學手冊"):
    st.markdown(
        """
        **1. 匯出名冊**：自機構或現有護理系統匯出個案當季狀況檔。
        **2. 隱私清理**：務必移除真實姓名、身分證字號，僅保留「個案代碼/床號」。
        **3. 套用範本**：將病況複製至提供的標準 Excel 範本。
        **4. 測試運行**：先點選「先測試前 3 筆」，檢查產出口吻。
        **5. 一鍵完成**：確認無誤後點擊批量生成並下載。
        """
    )

with st.sidebar.expander("📋 B | 訂閱方案與定價"):
    st.markdown(
        """
        * **單角色專業版**：NT$ 2,000 / 年（社工／居督／個管師 擇一）。
        * **雙模組整合版**：NT$ 3,500 / 年。
        * **全功能機構版**：NT$ 6,000 / 年（四合一完整開通）。
        * **年度使用配額**：單一帳戶每年上限 800 次報告呼叫。
        """
    )

with st.sidebar.expander("⚖️ C | 服務條款與責任界定"):
    st.markdown(
        """
        * **個資責任**：本工具不儲存個案明細，使用者負有上傳前資料脫敏之全責。
        * **核對責任**：AI 生成之內容屬輔助性質，正式送件前須經專業人員人工覆核。
        * **終止與變更**：本服務保留維護終止權。若因重大不可抗力終止營運，將按未到期天數比例合理結算。
        """
    )

# ──────────────────────────────────────────────
# 7. 主畫面顯示
# ──────────────────────────────────────────────
st.title("📋  AI 智慧行政工作台")
st.caption(f"⚡ 當前運行模式：**【{selected_role}專區】** ｜ 登入身份：{user_info['name']}")

# 讀取 API Key（優先讀 Secrets）
api_key = None
try:
    if "OPENAI_API_KEY" in st.secrets:
        api_key = st.secrets["OPENAI_API_KEY"]
except Exception:
    pass

if not api_key:
    api_key = st.sidebar.text_input("OpenAI API Key (本機測試用)", type="password")
    if not api_key:
        st.info("尚待系統載入金鑰，請聯繫管理員。")
        st.stop()

client = OpenAI(api_key=api_key)

# ==============================================================================
# 8. 依據選取的角色載入對應工作台
# ==============================================================================
if selected_role == "機構社工":
    tab_work, tab_guide = st.tabs(["🚀 季度評估報告生成台", "📖 操作指南與常見問題"])

    with tab_guide:
        st.subheader("📚 系統操作與匯入指南")
        st.markdown(
            """
            本工作台支援各類長照資訊系統（不限單一廠商）之 Excel / CSV 匯出檔，具備**智慧欄位對照**，無需繁瑣剪貼調整格式：

            #### 🔹 步驟一：從機構既有系統匯出資料
            1. 登入您機構平時使用的管理系統。
            2. 找到「住民清冊」或「當季評估/護理紀錄」匯出功能，下載為 **Excel (.xlsx)** 或 **CSV** 檔。
            
            #### 🔹 步驟二：個人資料去識別化（個資保護原則）
            * 匯出後，請確認僅保留 **「個案編號 / 床號」** 作為唯一辨識碼，務必**移除真實姓名與身分證字號**。

            #### 🔹 步驟三：直接上傳並對應欄位
            * 上傳檔案後，在下拉選單確認「個案代碼」與「病況/前季紀錄」對應哪一欄即可一鍵批次生成。
            """
        )

    with tab_work:
        st.info("💡 提醒：匯入資料請僅保留「個案編號/床號」，切勿包含真實姓名與身分證字號。")

        with st.expander("📥 若無現成系統匯出檔，可點此下載【標準格式 Excel 範本】"):
            st.markdown(
                """
                **填寫說明：**
                1. **個案編號**：機構內住民唯一識別碼（如病歷號、流水號）。
                2. **床號**：住民現住床號。
                3. **初始參考範本(前季度)**：貼上上一季度的評估紀錄（供 AI 對齊機構慣用語氣）；新入住個案可留白。
                4. **本季重大事件與觀察**：若本季無特殊事件請保持空白。
                """
            )
            template_df = pd.DataFrame({
                "個案編號": ["N001", "N002", "N003"],
                "床號": ["1F-101-1", "1F-101-2", "2F-201-1"],
                "基準診斷與病況": [
                    "失智症、高血壓，可自行操作輪椅活動，口語溝通順暢",
                    "腦中風後遺症、長期臥床，留置鼻胃管，需照服員定期翻身拍背",
                    "帕金森氏症，步態稍不穩，需助行器輔助行走"
                ],
                "初始參考範本(前季度)": [
                    "個案情緒穩定，能配合機構日常活動，與同儕互動良好。",
                    "個案臥床休養，意識清楚但口語表達有限，家屬每週定期探視。",
                    "個案日常作息規律，能主動參與復健運動，社交互動佳。"
                ],
                "本季重大事件與觀察": [
                    "本季開始參加園藝手作活動，專注度與笑容明顯增加。",
                    "",
                    "月初因感冒住院三天，返院後體力已逐步恢復。"
                ]
            })
            tpl_out = io.BytesIO()
            with pd.ExcelWriter(tpl_out, engine="openpyxl") as writer:
                template_df.to_excel(writer, index=False, sheet_name="評估名冊")
                ws = writer.sheets["評估名冊"]
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
                header_fill = PatternFill(start_color="2B547E", end_color="2B547E", fill_type="solid")
                data_font = Font(name="微軟正黑體", size=10)
                thin_border = Border(
                    left=Side(style="thin", color="D9D9D9"),
                    right=Side(style="thin", color="D9D9D9"),
                    top=Side(style="thin", color="D9D9D9"),
                    bottom=Side(style="thin", color="D9D9D9"),
                )
                col_widths = {"A": 14, "B": 14, "C": 32, "D": 45, "E": 35}
                for col_letter, width in col_widths.items():
                    ws.column_dimensions[col_letter].width = width
                for cell in ws[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border
                ws.row_dimensions[1].height = 28
                for row in ws.iter_rows(min_row=2, max_row=len(template_df)+1):
                    for cell in row:
                        cell.font = data_font
                        cell.alignment = Alignment(vertical="top", wrap_text=True)
                        cell.border = thin_border

            st.download_button(
                label="📥 點此下載【長照機構評估_初次建立標準範本.xlsx】",
                data=tpl_out.getvalue(),
                file_name="長照機構評估_初次建立標準範本.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        uploaded_file = st.file_uploader("上傳住民名冊 Excel / CSV（支援任意系統匯出檔或標準範本）", type=["xlsx", "csv"])

        if uploaded_file:
            if uploaded_file.name.endswith(".xlsx"):
                df = pd.read_excel(uploaded_file, dtype=str)
            else:
                df = pd.read_csv(uploaded_file, dtype=str)

            st.success(f"✅ 成功讀取檔案，共 {len(df)} 筆資料！")

            # 滾動歷史自動接軌（若為標準範本二次上傳）
            if "生成之季度評估報告" in df.columns:
                st.info("🔄 系統偵測到這份檔案包含上一次生成的評估報告！")
                if st.checkbox("🔄 自動將上一次的生成結果，轉為本次的「初始參考範本(前季度)」"):
                    df["初始參考範本(前季度)"] = df["生成之季度評估報告"]
                    if "本季重大事件與觀察" in df.columns:
                        df["本季重大事件與觀察"] = ""

            # 智慧欄位對照區塊
            st.markdown("##### 🔍 智慧欄位對照（免改欄位名稱，系統已自動推測，可微調）：")
            cols = df.columns.tolist()

            def guess_col(keywords, default_idx=0):
                for idx, c in enumerate(cols):
                    if any(k in str(c) for k in keywords):
                        return idx
                return min(default_idx, len(cols) - 1)

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                col_id = st.selectbox("1. 個案編號/代碼", cols, index=guess_col(["編號", "代碼", "案號", "ID"], 0))
            with c2:
                col_bed = st.selectbox("2. 床號（選填）", ["(無)"] + cols, index=(guess_col(["床號", "房號"], 0) + 1) if any("床" in str(c) or "房" in str(c) for c in cols) else 0)
            with c3:
                col_base = st.selectbox("3. 病況/前季評估紀錄", cols, index=guess_col(["範本", "評估", "病況", "診斷", "紀錄", "摘要"], 1))
            with c4:
                col_event = st.selectbox("4. 本季異動/觀察（選填）", ["(無)"] + cols, index=next((i + 1 for i, col_name in enumerate(cols) if any(k in str(col_name) for k in ["事件", "觀察", "異動", "備註"])), 0))

            st.dataframe(df.head(5), use_container_width=True)

            col_opt1, col_opt2 = st.columns(2)
            with col_opt1:
                quarter_label = st.text_input("本次評估季度標籤（用於匯出檔名）", value="115年Q4")
            with col_opt2:
                run_mode = st.radio("執行模式", ["🧪 先測試前 3 筆（確認語氣）", "🚀 全部批次生成"], horizontal=True)

            if st.button("✨ 開始生成專業社工季度評估報告", type="primary", use_container_width=True):
                target_df = df.head(3).copy() if "前 3 筆" in run_mode else df.copy()
                results = []
                progress_bar = st.progress(0)
                status_text = st.empty()

                for i, (_, row) in enumerate(target_df.iterrows()):
                    cid = str(row.get(col_id, f"個案{i+1}"))
                    base_info = str(row.get(col_base, "")) if pd.notna(row.get(col_base)) else ""
                    event_info = str(row.get(col_event, "")) if (col_event != "(無)" and pd.notna(row.get(col_event))) else ""

                    status_text.text(f"正在撰寫第 {i+1}/{len(target_df)} 筆（{cid}）...")

                    sys_prompt = (
                        "你是一位資深的台灣長照機構社工師。請根據提供的住民過往評估紀錄/病況與本季觀察，"
                        "直接撰寫一段專業、客觀、流暢且符合長照評鑑規範的季度社會心理與適應評估內文（約120～180字）。"
                        "嚴格要求：\n"
                        "1. 【直接輸出正文】：絕對不要在開頭加上『季度社會心理與適應評估報告』等任何標題或前言。\n"
                        "2. 【不寫個案編號】：內文中一律以「住民」或「個案」稱呼，絕對不要寫出個案代號或數字編號（例如不可寫「個案134」）。\n"
                        "3. 【專業語氣】：使用台灣長照社工慣用語彙；若無特殊重大事件，請自然延續並微調句型，切勿生硬照抄，也絕不捏造未提供的醫療診斷。"
                    )
                    user_prompt = f"過往紀錄/基本病況：{base_info}\n本季觀察與事件：{event_info if event_info.strip() else '本季作息與情緒平穩，無特殊重大事件。'}"

                    try:
                        resp = client.chat.completions.create(
                            model="gpt-4o-mini",
                            messages=[
                                {"role": "system", "content": sys_prompt},
                                {"role": "user", "content": user_prompt}
                            ],
                            temperature=0.7
                        )
                        raw_text = resp.choices[0].message.content.strip()
                        # 雙重保險：自動濾除開頭可能殘留的標題行與個案編號
                        raw_text = re.sub(r"^.*?評估報告[：:\s]*\n+", "", raw_text)
                        raw_text = raw_text.replace(f"個案{cid}", "個案").replace(f"住民{cid}", "住民")
                        results.append(raw_text.strip())
                    except Exception as e:
                        results.append(f"[生成失敗: {str(e)}]")
                    progress_bar.progress((i + 1) / len(target_df))

                target_df["生成之季度評估報告"] = results
                status_text.success("🎉 全部生成完成！")
                st.session_state["social_worker_df"] = target_df
                st.session_state["social_worker_quarter"] = quarter_label

            # 顯示與精緻 Excel 匯出區（存入 session_state，錄影按下載畫面也不會消失）
            if "social_worker_df" in st.session_state and st.session_state["social_worker_df"] is not None:
                final_sw_df = st.session_state["social_worker_df"]
                saved_q = st.session_state.get("social_worker_quarter", quarter_label)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    final_sw_df.to_excel(writer, index=False, sheet_name="季度評估結果")
                    ws = writer.sheets["季度評估結果"]
                    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                    from openpyxl.utils import get_column_letter

                    header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
                    header_fill = PatternFill(start_color="2B547E", end_color="2B547E", fill_type="solid")
                    body_font = Font(name="微軟正黑體", size=10)
                    thin_border = Border(
                        left=Side(style="thin", color="CCCCCC"),
                        right=Side(style="thin", color="CCCCCC"),
                        top=Side(style="thin", color="CCCCCC"),
                        bottom=Side(style="thin", color="CCCCCC"),
                    )

                    # 標題列樣式
                    for cell in ws[1]:
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.border = thin_border
                    ws.row_dimensions[1].height = 28

                    # 依據欄位名稱自動設定合適欄寬
                    for col_idx, col_name in enumerate(final_sw_df.columns, start=1):
                        col_letter = get_column_letter(col_idx)
                        name_str = str(col_name)
                        if any(k in name_str for k in ["編號", "床號", "代碼", "ID"]):
                            ws.column_dimensions[col_letter].width = 14
                        elif any(k in name_str for k in ["生成", "範本", "評估報告"]):
                            ws.column_dimensions[col_letter].width = 55
                        else:
                            ws.column_dimensions[col_letter].width = 32

                    # 資料列：自動換行 ＋ 動態計算列高
                    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=len(final_sw_df) + 1), start=2):
                        max_lines = 1
                        for cell in row:
                            cell.font = body_font
                            cell.border = thin_border
                            cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
                            val_str = str(cell.value or "")
                            est_lines = val_str.count("\n") + (len(val_str) // 26) + 1
                            if est_lines > max_lines:
                                max_lines = est_lines
                        ws.row_dimensions[row_idx].height = max(36, max_lines * 16)

                st.download_button(
                    label="📥 【點此下載】排版完成之完整 Excel 報告檔 (.xlsx)",
                    data=output.getvalue(),
                    file_name=f"社工評估報告_{saved_q}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                st.dataframe(final_sw_df, use_container_width=True)

# ==============================================================================
# 居家服務督導員（B單位居督）專屬工作台
# ==============================================================================
elif selected_role == "居家督導":
    tab_plan, tab_visit, tab_guide = st.tabs([
        "📋 個別化服務計畫與訪視生成台（支援 .html）",
        "📞 單筆快速訪視防重複產生器",
        "📖 居督模組操作指南"
    ])

    def parse_care_html(uploaded_html_file):
        raw_bytes = uploaded_html_file.read()
        try:
            html_str = raw_bytes.decode("utf-8")
        except UnicodeDecodeError:
            html_str = raw_bytes.decode("big5", errors="ignore")
        html_str = re.sub(r"<(script|style)[^>]*>.*?</\1>", "", html_str, flags=re.DOTALL | re.IGNORECASE)
        html_str = re.sub(r"</(tr|p|div|li|h[1-6]|br)>", "\n", html_str, flags=re.IGNORECASE)
        html_str = re.sub(r"</(td|th)>", " | ", html_str, flags=re.IGNORECASE)
        text = re.sub(r"<[^>]+>", "", html_str)
        text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        clean_text = "\n".join(lines)

        # 用 Python 預先幫 AI 把「每月X次」精準換算成「每週約Y次」，杜絕 AI 心算秀逗！
        def convert_monthly_to_weekly(match):
            num = int(match.group(1))
            weekly = max(1, round(num / 4))
            return f"每月{num}次（系統已精準換算為：每週{weekly}次）"

        clean_text = re.sub(r"每月\s*(\d+)\s*次", convert_monthly_to_weekly, clean_text)
        return clean_text

    # --------------------------------------------------------------------------
    # 子分頁 1：個別化服務計畫、評值、多份家訪/電訪生成台
    # --------------------------------------------------------------------------
    with tab_plan:
        st.info("🔒 個資安全提醒：系統僅在記憶體即時解析照管計畫內容，不會儲存任何個案資料。產出後請確認無真實姓名再貼回系統。")

        input_source = st.radio(
            "選擇資料匯入方式：",
            ["📂 直接拖曳上傳衛福部照管平台 .html 檔（支援單檔或多檔批次）", "✍️ 手動貼上計畫文字 / BA・G 碼條件"],
            horizontal=True
        )

        parsed_cases = []

        if "拖曳上傳" in input_source:
            html_files = st.file_uploader(
                "請直接拖曳「照顧服務管理資訊平台」匯出的 .html 檔案至此（可一次選取多個檔案）",
                type=["html", "htm"],
                accept_multiple_files=True,
                key="plan_html_uploader"
            )
            if html_files:
                for f in html_files:
                    clean_txt = parse_care_html(f)
                    parsed_cases.append({"source_name": f.name, "content": clean_txt})
                st.success(f"✅ 已成功解析 {len(parsed_cases)} 份照管計畫 HTML 檔案！")
                with st.expander("👀 檢視系統自動萃取的計畫純文字內容（已預先換算週計次數）"):
                    for item in parsed_cases:
                        st.markdown(f"**📄 檔案：{item['source_name']}**")
                        st.text(item["content"][:1500] + ("\n...(以下省略)" if len(item["content"]) > 1500 else ""))
        else:
            manual_txt = st.text_area(
                "請貼上照顧計畫內容或輸入服務條件（例如：BA02一週1次修剪指甲、BA07一週6次、GA09喘息一週1次）：",
                height=140,
                placeholder="貼上照管系統的照顧問題、核定 BA 碼與 G 碼別與頻率..."
            )
            if manual_txt.strip():
                def conv_m(m):
                    n = int(m.group(1))
                    return f"每月{n}次（系統已精準換算為：每週{max(1, round(n / 4))}次）"
                parsed_cases.append({
                    "source_name": "手動輸入個案",
                    "content": re.sub(r"每月\s*(\d+)\s*次", conv_m, manual_txt.strip())
                })

        st.markdown("---")
        st.markdown("##### 🎯 選擇要生成的文書項目與份數（可單獨生成或同時生成）：")

        # 第一排：服務計畫 & 定期評值
        col_top1, col_top2 = st.columns(2)
        with col_top1:
            gen_plan = st.checkbox("📋 個別化服務計畫（僅讀取 BA 碼、G 碼／強制週計／統整格式）", value=True)
        with col_top2:
            gen_eval = st.checkbox("🔄 定期評值追蹤（不限次數／強制週計不秀逗）", value=False)

        eval_note = ""
        if gen_eval:
            eval_note = st.text_input(
                "📝 定期評值條件補充（選填，不限第幾次）：",
                placeholder="例如：本次評值無變化（維持原計畫）／ 或：原 BA13 每週 2 次改為每週 1 次"
            )

        # 第二排：家訪（三個月一次，下拉 1~4 份） & 電訪（每月一次，下拉 1~12 份）
        col_v1, col_v2 = st.columns(2)
        with col_v1:
            with st.container(border=True):
                gen_home = st.checkbox("🏠 家訪督導紀錄（三個月一次）", value=False)
                home_count = st.selectbox(
                    "選擇同時生成「家訪紀錄」份數（1～4 份）：",
                    options=[1, 2, 3, 4],
                    index=0,
                    disabled=not gen_home,
                    key="home_count_select"
                )
        with col_v2:
            with st.container(border=True):
                gen_phone = st.checkbox("📞 電訪督導紀錄（每月一次）", value=False)
                phone_count = st.selectbox(
                    "選擇同時生成「電訪紀錄」份數（1～12 份）：",
                    options=list(range(1, 13)),
                    index=0,
                    disabled=not gen_phone,
                    key="phone_count_select"
                )

        visit_note = ""
        if gen_home or gen_phone:
            visit_note = st.text_input(
                "📝 訪視實際狀況補充（選填，套用於本次家訪/電訪）：",
                placeholder="例如：案主身體穩定、定期回診、居服員服務態度良好無異狀"
            )

        # 分開獨立的生成按鈕區（可只按某一種，也可一鍵全部生成）
        btn_col1, btn_col2, btn_col3 = st.columns(3)
        with btn_col1:
            click_plan_only = st.button("📋 僅生成服務計畫／評值", use_container_width=True, key="btn_only_plan")
        with btn_col2:
            click_visit_only = st.button("🏠📞 僅生成家訪／電訪紀錄", use_container_width=True, key="btn_only_visit")
        with btn_col3:
            click_all = st.button("✨ 依勾選項目一鍵全部生成", type="primary", use_container_width=True, key="btn_gen_all")

        # 判斷使用者按了哪一顆按鈕
        run_plan = False
        run_eval = False
        run_home = False
        run_phone = False
        triggered = False

        if click_plan_only:
            triggered = True
            run_plan = True
            run_eval = gen_eval
        elif click_visit_only:
            triggered = True
            run_home = gen_home
            run_phone = gen_phone
            if not (run_home or run_phone):
                st.warning("⚠️ 請先勾選「家訪督導紀錄」或「電訪督導紀錄」並選擇份數！")
                triggered = False
        elif click_all:
            triggered = True
            run_plan = gen_plan
            run_eval = gen_eval
            run_home = gen_home
            run_phone = gen_phone
            if not any([run_plan, run_eval, run_home, run_phone]):
                st.warning("⚠️ 請至少勾選一項要生成的文書項目！")
                triggered = False

        if triggered:
            if not parsed_cases:
                st.warning("⚠️ 請先上傳 .html 檔案或輸入計畫內容！")
            else:
                selected_tasks = []
                if run_plan:
                    selected_tasks.append(
                        "【個別化服務計畫】：\n"
                        "  1. 【碼別過濾鐵律】：只讀取照專計畫中核定的 **BA 碼（居家服務）** 與 **G 碼（如 GA09 喘息服務等）**，並將所有核定的 BA 碼與 G 碼項目**完整陳列、一項不漏**；嚴格排除 OT 碼（營養餐飲）、交通接送等其他碼別。\n"
                        "  2. 【統整版面鐵律（嚴禁分開循環打）】：全篇只能出現一次「問題說明」、一次「目標」、一次「計畫」這三個大標題！假設有 N 項核定服務（BA碼+G碼），格式必須嚴格如下：\n"
                        "     問題說明\n"
                        "     1. 案主因[具體原因]，在[第1項服務細項]能力受限，需他人協助，以維持[目的]。\n"
                        "     2. 案主因[具體原因]，在[第2項服務細項]能力受限，需他人協助，以維持[目的]。\n"
                        "     （依序條列至第 N 點）\n\n"
                        "     目標\n"
                        "     1. 協助案主完成[第1項服務細項]，維持/提升[具體目標]。\n"
                        "     2. 協助案主完成[第2項服務細項]，維持/提升[具體目標]。\n"
                        "     （依序條列至第 N 點）\n\n"
                        "     計畫\n"
                        "     1. 每週提供X次[第1項服務名稱，不需硬寫代碼或可簡寫服務名稱]，協助案主[具體服務內容]，以維持[目的]。\n"
                        "     2. 每週提供X次[第2項服務名稱]，協助案主[具體服務內容]，以維持[目的]。\n"
                        "     （依序條列至第 N 點）\n"
                        "  3. 【強制週計鐵律】：『計畫』區塊中的所有服務次數，**一律只能以「每週提供X次」呈現**（請直接採用資料中系統已換算好的『每週X次』數字，若無則將每月次數除以 4 四捨五入），**嚴禁出現「每月」或「每年」字眼**！"
                    )
                if run_eval:
                    selected_tasks.append(
                        f"【定期評值追蹤】：根據補充條件（{eval_note if eval_note.strip() else '目前服務執行順利，無特殊異動'}），撰寫一段約 80～110 字的精簡追蹤評值紀錄。若提及服務頻率，**一律只能寫「每週X次」（直接採用系統已換算之週計數字，嚴禁把每月次數誤寫成每週，也禁止出現每月/每年）**。全篇禁止出現「居督」二字主語。"
                    )
                if run_home:
                    selected_tasks.append(
                        f"【家訪督導紀錄（三個月一次，共需生成 {home_count} 份獨立紀錄）】：請依序標示『第1次家訪紀錄』至『第{home_count}次家訪紀錄』，每份約 100～130 字。僅針對資料中實際核定的 BA 碼與 G 碼服務項目及近況（{visit_note if visit_note.strip() else '整體生活與身體狀況穩定'}）撰寫，涵蓋案主身心與居家環境、服務執行落實度、家屬回饋與後續追蹤。**{home_count} 份的開頭、切入角度與句型必須明顯不同，絕不重複，且全程禁止出現「居督」主語**。"
                    )
                if run_phone:
                    selected_tasks.append(
                        f"【電訪督導紀錄（每月一次，共需生成 {phone_count} 份獨立紀錄）】：請依序標示『第1次電訪紀錄』至『第{phone_count}次電訪紀錄』，每份約 80～110 字。精準扣合該個案實際核定的 BA 碼與 G 碼服務項目及近況（{visit_note if visit_note.strip() else '身體狀況尚穩定、服務執行正常'}）。**{phone_count} 份的開頭、關懷重點與句型結構必須逐月變化、絕不複製貼上，且全程禁止出現「居督」主語**。"
                    )

                new_outputs = []

                with st.status("⏳ AI 正在解析計畫並撰寫專屬文書，請稍候（份數較多時約需 10～25 秒）...", expanded=True) as status_box:
                    prog = st.progress(0)
                    for idx, case in enumerate(parsed_cases):
                        st.write(f"🔄 正在處理第 {idx + 1} / {len(parsed_cases)} 份：**{case['source_name']}** ...")

                        sys_prompt = (
                            "你是一位資深的台灣長照 B 單位行政專家，精通衛福部照顧服務管理資訊平台登打規格。\n"
                            "核心鐵律：\n"
                            "1. 【只看 BA 碼與 G 碼】：只陳列照專核定的 BA 碼（居家服務）與 G 碼（喘息服務等），一項不漏；自動剔除 OT 碼等其他項目。\n"
                            "2. 【統整版面不拆開】：問題說明、目標、計畫各自集中條列 `1.~N.`，絕不可把每項服務拆成獨立段落循環輸出。\n"
                            "3. 【強制週計、嚴禁算錯】：所有服務頻率一律以「每週X次」表示（請嚴格使用括號內系統已換算好的每週次數），絕不可把每月次數直接當成每週次數，也禁止出現「每月」或「每年」。\n"
                            "4. 【嚴格去主語】：所有紀錄一律省略「居督」二字主語。\n"
                            "5. 【多份防重複】：若要求一次產出多份家訪或電訪，每份都必須使用不同的句型開頭與觀察重點。\n"
                            "6. 【隱私保護】：一律以「案主」或「案屬」稱呼，絕不輸出真實姓名、身分證字號或地址。"
                        )

                        user_prompt = (
                            f"以下是從照管平台萃取的完整個案計畫資料（已自動預先換算週計次數）：\n"
                            f"--------------------\n{case['content'][:18000]}\n--------------------\n\n"
                            f"請嚴格遵守所有格式與週計鐵律，生成以下指定項目：\n" + "\n".join(selected_tasks)
                        )

                        try:
                            resp = client.chat.completions.create(
                                model="gpt-4o-mini",
                                messages=[
                                    {"role": "system", "content": sys_prompt},
                                    {"role": "user", "content": user_prompt}
                                ],
                                temperature=0.45
                            )
                            res_text = resp.choices[0].message.content.strip()
                        except Exception as e:
                            res_text = f"❌ 生成發生錯誤：{str(e)}"

                        new_outputs.append({"file": case["source_name"], "result": res_text})
                        prog.progress((idx + 1) / len(parsed_cases))

                    status_box.update(label=f"✅ 全部完成！共成功生成 {len(new_outputs)} 位個案的專屬文書", state="complete", expanded=False)

                st.session_state["supervisor_outputs"] = new_outputs

        # ──────────────────────────────────────────────────────────────────────
        # 顯示與精緻 Excel 匯出區（存入 session_state，下載後畫面不消失）
        # ──────────────────────────────────────────────────────────────────────
        if "supervisor_outputs" in st.session_state and st.session_state["supervisor_outputs"]:
            saved_outputs = st.session_state["supervisor_outputs"]
            st.success(f"🎉 目前已保留 {len(saved_outputs)} 位個案的生成結果（下載檔案不會清除畫面）：")

            export_rows = []
            for item in saved_outputs:
                with st.container(border=True):
                    st.subheader(f"📄 個案來源：{item['file']}")
                    st.markdown(item["result"])
                    with st.expander("📋 點此展開純文字框（方便直接全選複製）"):
                        st.code(item["result"], language=None)

                clean_excel_text = re.sub(r"\*\*", "", item["result"])
                clean_excel_text = re.sub(r"#{1,4}\s*", "", clean_excel_text).strip()
                export_rows.append({
                    "來源檔案": item["file"],
                    "生成之居督個別化文書內容": clean_excel_text
                })

            df_exp = pd.DataFrame(export_rows)
            out_buf = io.BytesIO()
            with pd.ExcelWriter(out_buf, engine="openpyxl") as writer:
                df_exp.to_excel(writer, index=False, sheet_name="居督個別化文書")
                ws = writer.sheets["居督個別化文書"]
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

                header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
                header_fill = PatternFill(start_color="2B547E", end_color="2B547E", fill_type="solid")
                body_font = Font(name="微軟正黑體", size=10)
                thin_border = Border(
                    left=Side(style="thin", color="CCCCCC"),
                    right=Side(style="thin", color="CCCCCC"),
                    top=Side(style="thin", color="CCCCCC"),
                    bottom=Side(style="thin", color="CCCCCC"),
                )

                ws.column_dimensions["A"].width = 28
                ws.column_dimensions["B"].width = 85

                for cell in ws[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border
                ws.row_dimensions[1].height = 28

                for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=len(df_exp) + 1), start=2):
                    text_val = str(row[1].value or "")
                    line_count = text_val.count("\n") + (len(text_val) // 42) + 3
                    ws.row_dimensions[row_idx].height = max(60, line_count * 16)
                    for cell in row:
                        cell.font = body_font
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

            st.download_button(
                label="📥 打包下載排版完成之正式 Excel 報告 (.xlsx)",
                data=out_buf.getvalue(),
                file_name="居督個別化計畫與訪視紀錄.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

    # --------------------------------------------------------------------------
    # 子分頁 2：單筆快速訪視防重複產生器
    # --------------------------------------------------------------------------
    with tab_visit:
        st.subheader("📞 單筆／多份訪視紀錄防重複產生器（免上傳 HTML）")
        st.caption("臨時需補打單一個案的多份電訪（每月一次）或家訪（三個月一次）時，可直接勾選服務項目並選擇份數。")

        v_col1, v_col2 = st.columns(2)
        with v_col1:
            v_type = st.radio("選擇訪視類型：", ["🏠 家訪督導紀錄（三個月一次）", "📞 電訪督導紀錄（每月一次）"], horizontal=False)
        with v_col2:
            if "家訪" in v_type:
                v_count = st.selectbox("選擇同時生成家訪份數（1～4 份）：", options=[1, 2, 3, 4], index=0, key="quick_home_cnt")
            else:
                v_count = st.selectbox("選擇同時生成電訪份數（1～12 份）：", options=list(range(1, 13)), index=0, key="quick_phone_cnt")

        ba_options = st.multiselect(
            "勾選此個案實際核定的 BA 碼與 G 碼服務內容（可多選）：",
            [
                "BA01基本身體清潔",
                "BA02基本日常照顧",
                "BA03測量生命徵象",
                "BA04協助進食或管灌餵食",
                "BA05-1餐食照顧(一般備餐)",
                "BA07協助沐浴及洗頭",
                "BA08足部照護",
                "BA10翻身拍背",
                "BA11肢體關節活動",
                "BA12協助上下樓梯",
                "BA13陪同外出",
                "BA14陪同就醫",
                "BA15-1家務協助(自用)",
                "BA15-1家務協助(自用)(自費)",
                "BA15-2家務協助(共用)",
                "BA16-1代購或代領或代送服務(自用)",
                "BA16-2代購或代領或代送服務(共用)",
                "BA17a人工氣道管內(非氣管內管)分泌物抽吸",
                "BA17b口腔內(懸壅垂之前)分泌物抽吸",
                "BA17c尿管及鼻胃管之清潔與固定",
                "BA17d1血糖機驗血糖",
                "BA17d2甘油球通便",
                "BA17e依指示置入藥盒",
                "BA18安全看視",
                "BA20陪伴服務",
                "BA22巡視服務",
                "BA23協助洗頭",
                "BA24協助排泄",
                "GA09居家喘息"
            ],
            default=["BA01基本身體清潔", "BA15-1家務協助(自用)"]
        )        
        v_status = st.text_area(
            "輸入本次訪視了解到的近況與回饋：",
            value="案主身體狀況尚穩定，有規律回診；家屬表示居服員服務態度良好、準時且細心，符合期待。",
            height=100
        )

        if st.button("🚀 生成個別化訪視紀錄", type="primary", key="btn_quick_visit"):
            sys_p = (
                "你是台灣長照 B 單位資深督導。請根據勾選的實際 BA/G 碼服務項目與訪視近況，"
                f"依序撰寫 {v_count} 份不同句型結構的【{v_type}】（每則約 80～110 字，精簡不冗長）。"
                "嚴格要求：1. 全程禁止出現「居督」二字主語（直接寫「致電案家...」、「本次家訪觀察...」）。2. 只能提及勾選的服務項目。3. 每一份的開頭與句型必須明顯不同，絕不重複。"
            )
            user_p = f"核定服務項目：{', '.join(ba_options)}\n本次訪視近況：{v_status}"
            with st.spinner(f"正在撰寫 {v_count} 份高個別化訪視紀錄..."):
                resp = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "system", "content": sys_p}, {"role": "user", "content": user_p}],
                    temperature=0.8
                )
                st.markdown(resp.choices[0].message.content)

    # --------------------------------------------------------------------------
    # 子分頁 3：居督模組操作指南
    # --------------------------------------------------------------------------
    with tab_guide:
        st.subheader("📖 居家服務督導員模組使用說明")
        st.markdown("#### 🔹 1. 如何從衛福部照管平台匯出 `.html` 檔？\n1. 登入衛福部「照顧服務管理資訊平台」，進入個案的**照顧計畫**頁面。\n2. 在網頁空白處按右鍵選擇 **「另存新檔」**（或按快捷鍵 `Ctrl + S`），存檔類型選擇 **「網頁，僅 HTML」** 或 **「網頁，完整」**。\n3. 若透過通訊軟體（如 LINE）傳檔，檔案可能會被自動打包成 `.zip` 壓縮檔，上傳前請先在電腦上**解壓縮**，再將裡面的 `.html` 檔拖曳進本系統即可。\n\n#### 🔹 2. 本次升級核心規則說明\n* **碼別與週計自動換算**：系統僅擷取照專核定的 **BA 碼與 G 碼**，並自動在底層將每月次數除以 4 換算成「每週幾次」，徹底避免「每週 23 次」這類計算誤差。\n* **統整格式與多份訪視生成**：服務計畫統一採「問題說明 → 目標 → 計畫」三區塊集中條列；家訪（三個月一次，1～4 份）與電訪（每月一次，1～12 份）可獨立按需生成，且每份句型自動輪換不重複。")

elif selected_role == "居家行政(薪資拆帳)":
    render_home_care_payroll()

elif selected_role == "個管師":
    st.info("🚧 【A 單位個案管理師模組】建置中。")
    st.stop()
elif selected_role == "照顧專員":
    st.info("🚧 【照專綜合評估模組】建置中。")
    st.stop()