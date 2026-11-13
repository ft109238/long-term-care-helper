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
    page_title=" AI 智慧行政工作台",
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
                        "撰寫一段專業、客觀、流暢且符合長照評鑑規範的『季度社會心理與適應評估報告』（約120～180字）。"
                        "要求：1. 語氣溫暖專業，使用台灣長照社工慣用語彙。2. 若無特殊重大事件，請自然延續並微調句型，切勿生硬照抄。3. 絕不捏造未提供的醫療診斷。"
                    )
                    user_prompt = f"個案代號：{cid}\n過往紀錄/基本病況：{base_info}\n本季觀察與事件：{event_info if event_info.strip() else '本季作息與情緒平穩，無特殊重大事件。'}"

                    try:
                        resp = client.chat.completions.create(
                            model="gpt-4o-mini",
                            messages=[
                                {"role": "system", "content": sys_prompt},
                                {"role": "user", "content": user_prompt}
                            ],
                            temperature=0.7
                        )
                        results.append(resp.choices[0].message.content.strip())
                    except Exception as e:
                        results.append(f"[生成失敗: {str(e)}]")

                    progress_bar.progress((i + 1) / len(target_df))

                target_df["生成之季度評估報告"] = results
                status_text.success("🎉 全部生成完成！")

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    target_df.to_excel(writer, index=False, sheet_name="季度評估結果")

                st.download_button(
                    label="📥 【點此下載】排版完成之完整 Excel 報告檔 (.xlsx)",
                    data=output.getvalue(),
                    file_name=f"社工評估報告_{quarter_label}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                st.dataframe(target_df, use_container_width=True)


# ==============================================================================
# 居家服務督導員（B單位居督）專屬工作台
# ==============================================================================
elif selected_role == "居家督導":
    tab_plan, tab_visit, tab_guide = st.tabs([
        "📋 個別化服務計畫與評值（支援 .html 直接上傳）",
        "📞 客製化家訪／電訪督導紀錄",
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
        return "\n".join(lines)

    # --------------------------------------------------------------------------
    # 子分頁 1：個別化服務計畫與訪視紀錄
    # --------------------------------------------------------------------------
    with tab_plan:
        st.info("🔒 個資安全提醒：系統僅在記憶體即時解析照管計畫內容，不會儲存任何個案資料。產出後請確認無真實姓名再貼回系統。")

        input_source = st.radio(
            "選擇資料匯入方式：",
            ["📂 直接拖曳上傳衛福部照管平台 .html 檔（支援單檔或多檔批次）", "✍️ 手動貼上計畫文字 / BA 碼條件"],
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
                with st.expander("👀 檢視系統自動萃取的計畫純文字內容（確認無誤）"):
                    for item in parsed_cases:
                        st.markdown(f"**📄 檔案：{item['source_name']}**")
                        st.text(item["content"][:1500] + ("\n...(以下省略)" if len(item["content"]) > 1500 else ""))
        else:
            manual_txt = st.text_area(
                "請貼上照顧計畫內容或輸入服務條件（例如：BA01、BA11一週五次；BA13一週兩次用來外出復健，一次4-5組）：",
                height=140,
                placeholder="貼上照管系統的照顧問題、核定 BA 碼別與頻率..."
            )
            if manual_txt.strip():
                parsed_cases.append({"source_name": "手動輸入個案", "content": manual_txt.strip()})

        st.markdown("---")
        st.markdown("##### 🎯 勾選本次需要生成的文書項目（可複選）：")
        chk_col1, chk_col2, chk_col3, chk_col4 = st.columns(4)
        with chk_col1:
            gen_plan = st.checkbox("📋 個別化服務計畫\n（問題・目標・計畫）", value=True)
        with chk_col2:
            gen_eval = st.checkbox("🔄 定期評值追蹤\n（不限次數/異動調整）", value=False)
        with chk_col3:
            gen_home = st.checkbox("🏠 家訪督導紀錄\n（實地訪視評估）", value=False)
        with chk_col4:
            gen_phone = st.checkbox("📞 電訪督導紀錄\n（電話追蹤關懷）", value=False)

        eval_note = ""
        visit_note = ""
        if gen_eval:
            eval_note = st.text_input(
                "📝 定期評值條件補充（選填，不限第幾次）：",
                placeholder="例如：本次評值無變化（維持原計畫）／ 或：原 BA13 一週 2 次改為一週 1 次"
            )
        if gen_home or gen_phone:
            visit_note = st.text_input(
                "📝 本次訪視實際狀況補充（選填）：",
                placeholder="例如：案主身體穩定、定期回診、有至據點上課、居服員服務態度良好無異狀"
            )

    if st.button("✨ 開始生成個別化居督文書", type="primary", use_container_width=True, key="btn_gen_plan"):
                if not parsed_cases:
                    st.warning("⚠️ 請先上傳 .html 檔案或輸入計畫內容！")
                elif not any([gen_plan, gen_eval, gen_home, gen_phone]):
                    st.warning("⚠️ 請至少勾選一項要生成的文書項目！")
                else:
                    selected_tasks = []
                    if gen_plan:
                        selected_tasks.append(
                            "【個別化服務計畫】：\n"
                            "  ※ 重要過濾規則：請只抓取『居家照顧服務項目（BA碼別，例如 BA07、BA15-1、BA20 等）』，**自動忽略「喘息服務（GA碼/年度額度）」及交通接送**。\n"
                            "  ※ 點數絕對一致規則：假設篩選出 N 項 BA 服務，則『問題說明』、『目標』、『計畫』三個區塊**各自都必須嚴格列出 1.～N. 點**（1項服務=1點，三者點數必須完全相同）。\n"
                            "  ※ 頻率換算規則：在『計畫』區塊中，**一律統一以「每週」計算服務次數**（若原始資料寫每月次數，請自動除以 4 換算為每週幾次，例如每月8次→每週提供2次；每月27次→每週提供6至7次），**嚴禁出現「每月」或「每年」字眼**。\n"
                            "  輸出格式範例：\n"
                            "  問題說明\n"
                            "  1. 案主因[原因]，在[具體服務細項]能力受限，需他人協助，以維持[目的]。\n"
                            "  目標\n"
                            "  1. 協助案主完成[具體服務細項]，維持/提升[具體目標]。\n"
                            "  計畫\n"
                            "  1. 每週提供X次[服務名稱/碼別]，協助案主[具體服務細節]，以維持[目的]。"
                        )
                    if gen_eval:
                        selected_tasks.append(
                            f"【定期評值追蹤】：根據補充條件（{eval_note if eval_note.strip() else '目前服務執行順利，無特殊異動'}），撰寫一段約 80～110 字的精簡追蹤評值紀錄。若有調整頻率/碼別，須交代考量原因、調整前後對比與後續追蹤；若無變化，簡潔說明各項服務符合案主需求、維持原計畫即可。**全篇禁止出現「居督」二字主語**。"
                        )
                    if gen_home:
                        selected_tasks.append(
                            f"【家訪督導紀錄】：僅根據資料中『實際核定的 BA 服務項目』與近況（{visit_note if visit_note.strip() else '整體生活與身體狀況穩定'}），撰寫一段約 100～130 字的精簡實地訪視紀錄，涵蓋案主身心及居家環境、服務執行落實度、家屬回饋與後續追蹤。**全程使用客觀記錄口吻，禁止出現「居督」作為主語**（直接寫「本次家訪觀察案主...」）。"
                        )
                    if gen_phone:
                        selected_tasks.append(
                            f"【電訪督導紀錄】：僅根據資料中『實際核定的 BA 服務項目』與近況（{visit_note if visit_note.strip() else '身體狀況尚穩定、服務執行正常'}），撰寫一段約 80～110 字的精簡電訪紀錄。**全程省略「居督」主語**（直接寫「致電案屬關心近況...」或「電訪了解案主近期生活情形...」）。"
                        )

                    new_outputs = []

                    with st.status("⏳ AI 正在解析計畫並撰寫專屬文書，請稍候（約 5～15 秒）...", expanded=True) as status_box:
                        prog = st.progress(0)
                        for idx, case in enumerate(parsed_cases):
                            st.write(f"🔄 正在處理第 {idx + 1} / {len(parsed_cases)} 份：**{case['source_name']}** ...")

                            style_hints = [
                                "訪視紀錄開頭請從案主近期生活與身體狀況切入，用詞精煉。",
                                "訪視紀錄開頭請從服務執行情形與案家回饋切入，條理分明。",
                                "訪視紀錄開頭請以定期追蹤關懷視角切入，避免制式客套話。",
                                "語句簡潔俐落、貼近台灣長照系統實務登打風格，不寫冗長廢話。"
                            ]
                            variety_rule = style_hints[idx % len(style_hints)]

                            sys_prompt = (
                                "你是一位資深的台灣長照 B 單位行政專家，精通衛福部照顧服務管理資訊平台登打規格。\n"
                                "核心要求：\n"
                                "1. 【只看 BA 居家服務、一律以每週計算】：只針對 BA 碼居家服務撰寫，自動排除喘息服務（GA碼）。計畫頻率一律換算為「每週提供X次」，絕不寫每月或每年。\n"
                                "2. 【1對1點數完全一致】：有幾項 BA 服務，問題說明、目標、計畫就各自列幾點，不可多也不可少。\n"
                                "3. 【嚴格去主語】：所有紀錄一律省略「居督」二字主語。\n"
                                "4. 【防重複與精簡】：" + variety_rule + " 每點控制在 25～45 字內，精準不冗長。\n"
                                "5. 【隱私保護】：一律以「案主」或「案屬」稱呼，絕不輸出真實姓名、身分證字號或地址。"
                            )

                            user_prompt = (
                                f"以下是從照管平台萃取的完整個案計畫資料（請閱讀至後半段的『核定服務項目與BA碼別明細』）：\n"
                                f"--------------------\n{case['content'][:18000]}\n--------------------\n\n"
                                f"請嚴格遵守『只抓 BA 服務、計畫一律以每週計算、三區塊點數 1 對 1 一致』，生成以下勾選的項目：\n" + "\n".join(selected_tasks)
                            )

                            try:
                                resp = client.chat.completions.create(
                                    model="gpt-4o-mini",
                                    messages=[
                                        {"role": "system", "content": sys_prompt},
                                        {"role": "user", "content": user_prompt}
                                    ],
                                    temperature=0.4
                                )
                                res_text = resp.choices[0].message.content.strip()
                            except Exception as e:
                                res_text = f"❌ 生成發生錯誤：{str(e)}"

                            new_outputs.append({"file": case["source_name"], "result": res_text})
                            prog.progress((idx + 1) / len(parsed_cases))

                        status_box.update(label=f"✅ 全部完成！共成功生成 {len(new_outputs)} 位個案的專屬文書", state="complete", expanded=False)

                    # 將結果存入 session_state，按下載按鈕時畫面絕不會消失！
                    st.session_state["supervisor_outputs"] = new_outputs

            # ──────────────────────────────────────────────────────────────────────
            # 顯示區（讀取 session_state，下載後依然穩固保留在畫面上，不重跑 API）
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

            # 清理 Markdown 星號(**)與多餘井字號(###)，讓匯出到 Excel 的文字乾淨美觀
            clean_excel_text = re.sub(r"\*\*", "", item["result"])
            clean_excel_text = re.sub(r"#{1,4}\s*", "", clean_excel_text).strip()
            export_rows.append({
                "來源檔案": item["file"],
                "生成之居督個別化文書內容": clean_excel_text
            })

        # 產生精緻排版之 Excel（自動換行、適當欄寬、自動調整列高）
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
                for col_i, cell in enumerate(row):
                    cell.font = body_font
                    cell.border = thin_border
                    cell.alignment = Alignment(
                        horizontal="left",
                        vertical="top",
                        wrap_text=True
                    )

        st.download_button(
            label="📥 打包下載排版完成之正式 Excel 報告 (.xlsx)",
            data=out_buf.getvalue(),
            file_name="居督個別化計畫與訪視紀錄.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    # --------------------------------------------------------------------------
    # 子分頁 2：純訪視紀錄快速批次生成
    # --------------------------------------------------------------------------
    with tab_visit:
        st.subheader("📞 快速單筆／多筆訪視紀錄防重複產生器")
        st.caption("若臨時只需補打單筆電訪或家訪，可直接在此勾選該個案實際使用的服務項目，系統自動避開重複句型。")

        v_type = st.radio("訪視方式", ["📞 電話訪視（電訪）", "🏠 實地訪視（家訪）"], horizontal=True)
        ba_options = st.multiselect(
            "勾選此個案實際核定的服務內容（可多選）：",
            [
                "BA01 基本身體清潔", "BA02 基本日常照顧", "BA04 協助進食或管灌",
                "BA05 餐食照顧", "BA07 協助沐浴及洗頭", "BA10 翻身拍背",
                "BA11 肢體關節活動", "BA12 協助上(下)樓梯", "BA13 陪同外出（如復健/採買）",
                "BA14 陪同就醫", "BA15 家務協助", "BA16 代購或代領或代送服務",
                "BA17 協助執行輔助性醫療", "BA20 陪伴服務", "BA23 協助洗頭", "BA24 協助排泄"
            ],
            default=["BA01 基本身體清潔", "BA15 家務協助"]
        )
        v_status = st.text_area(
            "輸入本次訪視了解到的近況與回饋：",
            value="案主身體狀況尚穩定，有規律回診，目前平日會至社區據點上課；家屬表示居服員服務態度良好、準時且細心，符合期待。",
            height=100
        )
        v_count = st.slider("一次產生幾種不同表述版本供挑選？（防重複引擎）", min_value=1, max_value=5, value=2)

        if st.button("🚀 生成個別化訪視紀錄", type="primary", key="btn_quick_visit"):
            sys_p = (
                "你是台灣長照 B 單位資深督導。請根據勾選的實際 BA 服務項目與訪視近況，"
                f"撰寫 {v_count} 個不同句型結構的【{v_type}】紀錄（每則約 80～110 字，精簡不冗長）。"
                "嚴格要求：1. 全程禁止出現「居督」二字主語（直接寫「致電案家...」、「本次家訪觀察...」）。2. 只能提及勾選的服務項目。3. 每個版本的開頭與句型必須明顯不同。"
            )
            user_p = f"核定服務項目：{', '.join(ba_options)}\n本次訪視近況：{v_status}"
            with st.spinner("正在撰寫高個別化訪視紀錄..."):
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
        st.markdown(
            """
            #### 🔹 1. 如何從衛福部照管平台匯出 `.html` 檔？
            1. 登入衛福部「照顧服務管理資訊平台」，進入個案的**照顧計畫**頁面。
            2. 在網頁空白處按右鍵選擇 **「另存新檔」**（或按快捷鍵 `Ctrl + S`），存檔類型選擇 **「網頁，僅 HTML」** 或 **「網頁，完整」**。
            3. 若透過通訊軟體（如 LINE）傳檔，檔案可能會被自動打包成 `.zip` 壓縮檔，上傳前請先在電腦上**解壓縮**，再將裡面的 `.html` 檔拖曳進本系統即可。

            #### 🔹 2. 多項目勾選與防重複機制
            * 系統會自動讀取 `.html` 內的照顧問題與 BA 服務碼別，**完全排除個案未核定的項目**。
            * 支援一次拖曳多份 `.html` 檔批次處理，系統會自動輪換句型開頭與描述視角，確保每位個案的計畫、評值與訪視紀錄皆具高度個別化。
            """
        )


elif selected_role == "個管師":
    st.info("🚧 【A 單位個案管理師模組】建置中。")
    st.stop()
elif selected_role == "照顧專員":
    st.info("🚧 【照專綜合評估模組】建置中。")
    st.stop()