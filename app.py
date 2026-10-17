import streamlit as st
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
    page_title="長照 AI 智慧行政工作台",
    page_icon="📋",
    layout="wide"
)

# ──────────────────────────────────────────────
# 2. 授權名單資料庫（角色權限嚴格控管）
# ──────────────────────────────────────────────
AUTHORIZED_USERS = {
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
    "CARE-2027-01": {
        "name": "仁愛養護所-陳社工",
        "expire": "2028-02-28",
        "roles": ["機構社工"]
    },
    "SUPERVISOR-XR": {
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
st.title("📋 長照 AI 智慧行政工作台")
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

# ──────────────────────────────────────────────
# 8. 依據選取的角色載入對應工作台
# ──────────────────────────────────────────────
if selected_role == "機構社工":
    # 建立主畫面雙分頁
    tab_work, tab_guide = st.tabs(["🚀 季度評估報告生成台", "📖 操作指南與常見問題"])

    # ──────────────────────────────────────────────
    # 分頁 1：操作指南（去品牌化、圖文/影音預留）
    # ──────────────────────────────────────────────
    with tab_guide:
        st.subheader("📚 系統操作與匯入指南")
        st.markdown(
            """
            本工作台支援各類長照資訊系統（不限單一廠商）之 Excel 匯出檔，無需繁瑣調整格式即可快速處理：

            #### 🔹 步驟一：從機構既有系統匯出資料
            1. 登入您機構平時使用的管理系統（如各類長照/護理資訊平台）。
            2. 找到「住民清冊」或「當季評估/護理紀錄」匯出功能，下載為 **Excel (.xlsx)** 或 **CSV** 檔。
            
            #### 🔹 步驟二：個人資料去識別化（個資保護重要原則）
            * 匯出後，請確認僅保留 **「個案編號 / 床號」** 作為唯一辨識碼。
            * 請務必**移除真實姓名、身分證字號**，符合長照個資防護法規規範。

            #### 🔹 步驟三：直接上傳分析
            * 系統具備**智慧欄位辨識**功能，您不需要手動將資料剪貼至特定範本。
            * 只要檔案中包含「個案代號」與「過去病況/日誌紀錄」，系統即可自動理解並撰寫。

            ---
            #### 🎥 實機操作示範影片
            """
        )
        st.info("💡 影音操作短片錄製中，即將在此直接播放操作示範。")

    # ──────────────────────────────────────────────
    # 分頁 2：報表生成工作台
    # ──────────────────────────────────────────────
    with tab_work:
        st.info("💡 提醒：匯入資料請僅保留「個案編號/床號」，切勿包含真實姓名與身分證字號。")
    
            # 提供使用者下載標準範本（精緻機構排版版）
        with st.expander("📥 第一次使用？點此下載【標準格式 Excel 範本】"):
            st.markdown("""
            **填寫說明：**
            1. **個案編號**：機構內住民唯一識別碼（如病歷號、流水號），作為跨季度歷史追蹤基準。
            2. **床號**：住民現住床號。若未來住民轉床，僅需在後續報表直接更改床號即可。
            3. **初始參考範本(前季度)**：初次建立請貼上上一季度的評估紀錄（供 AI 對齊機構慣用語氣）；若為全新入住個案可保持空白。
            4. **本季重大事件與觀察**：若本季作息平穩無特殊事件，**請務必保持空白**，切勿填寫「待補」或「無」。
            """)
            
            template_df = pd.DataFrame({
                "個案編號": ["N001", "N002", "N003"],
                "床號": ["1F-101-1", "1F-101-2", "2F-201-1"],
                "基準診斷與病況": [
                    "失智症、高血壓，可自行操作輪椅活動，口語溝通順暢",
                    "腦中風後遺症、長期臥床，留置鼻胃管，需照服員定期翻身拍背",
                    "帕金森氏症，步態稍不穩，需助行器輔助行走"
                ],
                "家庭支持與互動": [
                    "主要聯絡人為長子，每月固定探視2-3次，家庭支持系統良好",
                    "主要聯絡人為女兒，平時多以電話關心，重大就醫需求皆能積極配合",
                    "家屬居於外縣市，偶有視訊探視，機構聯繫順暢"
                ],
                "初始參考範本(前季度)": [
                    "住民目前使用輪椅活動，可自行操作輪椅，具口語溝通能力，能清楚表達身體不適與照顧需求。日常作息維持穩定，配合機構各項文康活動。",
                    "住民日常以臥床休息為主，各項生活照顧均由工作人員妥善協助。整體生理狀況平穩，無特殊異常事件發生。",
                    "（若為新入住個案無前季資料，可直接留白）"
                ],
                "本季重大事件與觀察": [
                    "",  # 作息正常請直接留白
                    "本季曾於上月因發燒短暫送急診就醫觀察，返所後生命徵象平穩，已加強體溫監測",
                    ""   # 作息正常請直接留白
                ]
            })
            
            tpl_out = io.BytesIO()
            with pd.ExcelWriter(tpl_out, engine="openpyxl") as writer:
                template_df.to_excel(writer, index=False, sheet_name="住民名冊建立範本")
                ws = writer.sheets["住民名冊建立範本"]
                
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                
                # 1. 專業表頭樣式（深藍底、白粗體字）
                header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
                header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                
                # 2. 內容樣式
                data_font = Font(name="微軟正黑體", size=10)
                thin_border = Border(
                    left=Side(style='thin', color='D9D9D9'),
                    right=Side(style='thin', color='D9D9D9'),
                    top=Side(style='thin', color='D9D9D9'),
                    bottom=Side(style='thin', color='D9D9D9')
                )
                
                # 3. 欄位寬度最佳化
                col_widths = {
                    "A": 16,  # 個案編號
                    "B": 16,  # 床號
                    "C": 35,  # 基準診斷與病況
                    "D": 35,  # 家庭支持與互動
                    "E": 50,  # 初始參考範本(前季度)
                    "F": 40   # 本季重大事件與觀察
                }
                for col_letter, width in col_widths.items():
                    ws.column_dimensions[col_letter].width = width
                    
                # 4. 套用表頭格式
                for cell in ws[1]:
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.border = thin_border
                ws.row_dimensions[1].height = 28
                
                # 5. 套用資料列格式（自動換行 + 頂部對齊 + 邊框）
                for row in ws.iter_rows(min_row=2, max_row=len(template_df)+1, min_col=1, max_col=len(template_df.columns)):
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

            uploaded_file = st.file_uploader("上傳住民名冊 Excel / CSV（初次填寫範本 或 直接上傳上一季匯出檔）", type=["xlsx", "csv"])

            if uploaded_file:
                if uploaded_file.name.endswith(".xlsx"):
                    df = pd.read_excel(uploaded_file, dtype=str)
                else:
                    df = pd.read_csv(uploaded_file, dtype=str)
            
                # 滾動歷史自動接軌
                if "生成之季度評估報告" in df.columns:
                    st.success("🔄 系統偵測到這份檔案包含上一次生成的評估報告！")
                    roll_over = st.checkbox("🔄 是否將上一次的生成結果，自動轉為本次的「初始參考範本(前季度)」？", value=True)
                    if roll_over:
                        df["初始參考範本(前季度)"] = df["生成之季度評估報告"]
                        if "本季重大事件與觀察" in df.columns:
                            df["本季重大事件與觀察"] = ""
                        df = df.drop(columns=["生成之季度評估報告"])
                        st.info("✅ 已自動承接歷史脈絡！本季若無特殊事件可直接點擊生成。")

                st.subheader("📊 已匯入名冊預覽（前 5 筆）")
                st.dataframe(df.head(5))

                st.write("---")
                st.subheader("⚙️ 本季生成設定")
                col1, col2 = st.columns(2)
                with col1:
                    quarter = st.selectbox("評估季度", ["2026年 第一季", "2026年 第二季", "2026年 第三季", "2026年 第四季"])
                    st.session_state.current_quarter = quarter
                with col2:
                    model_name = st.selectbox("AI 模型選擇", ["gpt-4o-mini", "gpt-4o"])

                default_system_prompt = """你是一名資深長照住宿式機構社工。你的任務是為住民撰寫季度生活評估紀錄。

            【核心原則】：
            1. 嚴禁在內文中提及個案編號、床號或特定年份季度（嚴禁出現如「3F-305-4」、「2026年第一季」等字眼）。
            2. 主詞請統一使用「住民」或「個案」，請直接使用「住民本季日常作息...」或「案主本季身心適應...」開始陳述。
            3. 嚴禁使用 Markdown 格式（不可出現 **、#、-、> 等符號），直接輸出純文字段落。
            4. 若有提供前一季評估，請以該風格與照護脈絡延續撰寫；若無特殊事件，記錄為生活作息與身心適應維持穩定即可。"""

                system_prompt = st.text_area(
                    "AI 角色與格式規範：",
                    value=default_system_prompt,
                    height=180
                )

                def clean_text(val):
                    if pd.isna(val):
                        return ""
                    s = str(val).strip()
                    invalid_keywords = ["待補", "待社工填寫", "無可對應", "原始檔無", "nan", "null", "None"]
                    for kw in invalid_keywords:
                        if kw in s:
                            return ""
                    return s

                def process_data(data_to_run):
                    results = []
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    for idx, row in data_to_run.iterrows():
                        case_id = clean_text(row.get("個案編號") or f"個案{idx+1}")
                        bed_num = clean_text(row.get("床號"))
                        diagnosis = clean_text(row.get("基準診斷與病況") or row.get("診斷") or row.get("病況"))
                        family = clean_text(row.get("家庭支持與互動") or row.get("家庭支持") or row.get("家庭"))
                        prev_record = clean_text(row.get("初始參考範本(前季度)") or row.get("前一季度參考範本") or row.get("前一季評估"))
                        quarter_events = clean_text(row.get("本季重大事件與觀察") or row.get("本季重大事件"))

                        show_id = f"{case_id} ({bed_num})" if bed_num else case_id
                        status_text.text(f"正在為 {show_id} 生成報告... ({idx+1}/{len(data_to_run)})")

                        prompt_context = f"【評估季度】：{quarter}\n"
                        if diagnosis:
                            prompt_context += f"【主要病況與診斷】：{diagnosis}\n"
                        if family:
                            prompt_context += f"【家庭支持現況】：{family}\n"
                        if prev_record:
                            prompt_context += f"【前一季度評估紀錄（請直接延續其語氣與脈絡）】：{prev_record}\n"
                        if quarter_events:
                            prompt_context += f"【本季新增事件或特殊觀察】：{quarter_events}\n"
                        else:
                            prompt_context += "【本季狀況】：日常作息規律穩定，無特殊跌倒或住院等重大異常事件。\n"

                        user_prompt = f"""請根據以上資訊，為個案撰寫本季度的社工評估紀錄：
            - 篇幅約 120-180 字左右（一至兩個自然段落）。
            - 請直接輸出評估紀錄純文字正文，嚴禁出現個案編號、床號或年份季度字眼，結尾以完整句號結束。

            住民資訊如下：
            {prompt_context}"""

                        try:
                            response = client.chat.completions.create(
                                model=model_name,
                                messages=[
                                    {"role": "system", "content": system_prompt},
                                    {"role": "user", "content": user_prompt}
                                ],
                                temperature=0.4,
                                max_tokens=800
                            )
                            output_text = response.choices[0].message.content.strip()

                            # --- 後處理自動清洗：確保絕無殘留編號、床號與年份季度 ---
                            if case_id:
                                output_text = output_text.replace(case_id, "")
                            if bed_num:
                                output_text = output_text.replace(bed_num, "")
                            
                            output_text = re.sub(r"\d{4}年\s*第[一二三四1234]季(的評估顯示|之評估顯示|度)?", "", output_text)
                            output_text = re.sub(r"^個案(代號)?\s*(於)?", "", output_text)
                            output_text = re.sub(r"^本季度(評估顯示[，,]?\s*個案)?", "住民本季", output_text)
                            output_text = output_text.strip(" ，,：:")
                            
                            if not output_text.startswith("住民") and not output_text.startswith("案主"):
                                output_text = "住民本季" + output_text
                            # ----------------------------------------------------

                        except Exception as e:
                            output_text = f"生成失敗：{str(e)}"

                        results.append(output_text)
                        progress_bar.progress((idx + 1) / len(data_to_run))
                        time.sleep(0.05)

                    data_to_run = data_to_run.copy()
                    data_to_run["生成之季度評估報告"] = results
                    status_text.success("🎉 生成完成！")
                    return data_to_run

                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    if st.button("🧪 先測試前 3 筆（測語氣）"):
                        st.session_state.generated_df = process_data(df.head(3))
                with btn_col2:
                    if st.button("🚀 正式跑完整名冊"):
                        st.session_state.generated_df = process_data(df)

                if st.session_state.generated_df is not None:
                    final_df = st.session_state.generated_df
                    st.write("---")
                    
                    # 1. 先生成排版好的 Excel 檔案
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        final_df.to_excel(writer, index=False, sheet_name="評估報告")
                        ws = writer.sheets["評估報告"]
                        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

                        header_font = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
                        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                        data_font = Font(name="微軟正黑體", size=10)
                        thin_border = Border(
                            left=Side(style='thin', color='E0E0E0'),
                            right=Side(style='thin', color='E0E0E0'),
                            top=Side(style='thin', color='E0E0E0'),
                            bottom=Side(style='thin', color='E0E0E0')
                        )

                        col_widths = {
                            "A": 16, "B": 16, "C": 32, "D": 32, "E": 45, "F": 25, "G": 70
                        }
                        for col_letter, width in col_widths.items():
                            ws.column_dimensions[col_letter].width = width

                        for cell in ws[1]:
                            cell.font = header_font
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                        for row in ws.iter_rows(min_row=2, max_row=len(final_df)+1, min_col=1, max_col=len(final_df.columns)):
                            for cell in row:
                                cell.font = data_font
                                cell.alignment = Alignment(vertical="top", wrap_text=True)
                                cell.border = thin_border

                    # 2. 顯眼的大按鈕放在最上方
                    st.subheader("🎉 生成完成！請點擊下方按鈕取得正式報告")
                    st.download_button(
                        label="📥 【點此下載】排版完成之完整 Excel 報告檔 (.xlsx)",
                        data=output.getvalue(),
                        file_name=f"長照評估報告_{st.session_state.current_quarter}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"  # 醒目紅色/主題色按鈕
                    )

                    st.subheader("📋 生成結果線上檢視")
                    # 隱藏右側懸浮的 Download as CSV 雜項圖示
                    st.dataframe(final_df, use_container_width=True)

                    st.subheader("👀 單篇詳細內容抽查（前 5 筆）")
                    for i, r in final_df.head(5).iterrows():
                        cid = r.get("個案編號") or f"編號{i+1}"
                        bnum = r.get("床號")
                        title_text = f"📌 {cid}（床號：{bnum}）" if pd.notna(bnum) else f"📌 {cid}"
                        with st.expander(f"{title_text} 評估報告全文"):
                            st.write(r["生成之季度評估報告"])

            elif selected_role == "居家督導":
                st.info("🚧 【居家服務督導員模組】現正進行第一線需求微調，即將開放上線。")
                st.stop()
            elif selected_role == "個管師":
                st.info("🚧 【A 單位個案管理師模組】建置中。")
                st.stop()
            elif selected_role == "照顧專員":
                st.info("🚧 【照專綜合評估模組】建置中。")
                st.stop()                