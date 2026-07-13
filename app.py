import streamlit as st
import pandas as pd
import datetime
import os
import smtplib
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import gspread
from google.oauth2.service_account import Credentials
import openai

# 1. 基礎網頁配置與美化
st.set_page_config(page_title="Geann B2B 業務戰情監控版 (完全體)", layout="wide")

st.markdown("""
    <style>
    .main-title { font-size:32px; font-weight:bold; color:#1E3A8A; margin-bottom:10px; }
    .stButton>button { width: 100%; border-radius: 8px; background-color: #1E3A8A; color: white; }
    .stButton>button:hover { background-color: #2563EB; color: white; }
    </style>
    <div class="main-title">🚀 Geann B2B 業務戰情監控版 (完全體)</div>
""", unsafe_allow_html=True)

# 2. 安全讀取保險箱後台金鑰 (Secrets)
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", "")
GMAIL_ACCOUNT = st.secrets.get("GMAIL_ACCOUNT", "")
GMAIL_APP_PASSWORD = st.secrets.get("GMAIL_APP_PASSWORD", "")
SHEET_URL = "https://docs.google.com/spreadsheets/d/1j10PqYMrg09T-N21UCrnZs7UVqXb1ZmglRgjZmxcbgs/edit?gid=0#gid=0"

# 業務名單與 Email 對照
SALES_MAP = {
    "Frances": "francesma@geann.com.tw",
    "Amanda": "amandalee@geann.com.tw",
    "Alex": "alexwang@geann.com.tw",
    "Jay": "Jayhuang@geann.com.tw"
}
BOSS_EMAIL = "fionawang@geann.com.tw"

# 3. 初始化 Google Sheets 連線
@st.cache_resource
def init_gspread():
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        gcp_sa = dict(st.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(gcp_sa, scopes=scope)
        return gspread.authorize(credentials)
    except Exception as e:
        st.error(f"Google 試算表連線失敗，請檢查金鑰設定。錯誤訊息: {e}")
        return None

gc = init_gspread()

def load_data_from_sheet():
    if gc:
        try:
            sh = gc.open_by_url(SHEET_URL)
            worksheet = sh.get_worksheet(0)
            data = worksheet.get_all_records()
            if not data:
                return pd.DataFrame(columns=["進度狀態", "已回覆", "業務備註", "最後聯絡日期", "負責業務", "公司名稱", "網址", "First Name", "Last Name", "Job Title", "Email"])
            df = pd.DataFrame(data)
            df["已回覆"] = df["已回覆"].astype(bool)
            return df
        except Exception as e:
            st.error(f"讀取雲端資料失敗: {e}")
    return pd.DataFrame()

def save_data_to_sheet(df):
    if gc:
        try:
            sh = gc.open_by_url(SHEET_URL)
            worksheet = sh.get_worksheet(0)
            worksheet.clear()
            # 確保欄位順序一字不差
            cols = ["進度狀態", "已回覆", "業務備註", "最後聯絡日期", "負責業務", "公司名稱", "網址", "First Name", "Last Name", "Job Title", "Email"]
            df_to_save = df[cols].copy()
            df_to_save["已回覆"] = df_to_save["已回覆"].astype(str).str.upper() # 轉成 TRUE/FALSE 方便試算表讀取
            worksheet.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())
            return True
        except Exception as e:
            st.error(f"寫入雲端資料失敗: {e}")
    return False

# 讀取當前資料庫
df_crm = load_data_from_sheet()

# 4. 建立雙頁籤系統
tab1, tab2 = st.tabs(["🚀 名單開火與跟催區", "🗄️ CRM 客戶池總覽"])

# ==================== 頁籤 1：開火與跟催區 ====================
with tab1:
    st.subheader("📥 匯入 Apollo 尋人名單 (新彈藥)")
    uploaded_file = st.file_uploader("請上傳清洗後的 Keyman 名單 (Excel)", type=["xlsx"])
    
    if uploaded_file:
        try:
            input_df = pd.read_excel(uploaded_file)
            # 自動補足戰情所需控制欄位
            if "進度狀態" not in input_df.columns:
                input_df.insert(0, "進度狀態", "準備開發")
            if "已回覆" not in input_df.columns:
                input_df.insert(1, "已回覆", False)
            if "業務備註" not in input_df.columns:
                input_df.insert(2, "業務備註", "")
            if "最後聯絡日期" not in input_df.columns:
                input_df.insert(3, "最後聯絡日期", "")
            if "負責業務" not in input_df.columns:
                input_df.insert(4, "負責業務", "未指定")
            
            # 對齊欄位
            cols = ["進度狀態", "已回覆", "業務備註", "最後聯絡日期", "負責業務", "公司名稱", "網址", "First Name", "Last Name", "Job Title", "Email"]
            input_df = input_df[cols]
            
            if st.button("🔥 確認將新名單匯入雲端資料庫"):
                if not df_crm.empty:
                    df_combined = pd.concat([df_crm, input_df]).drop_duplicates(subset=["Email"], keep="last")
                else:
                    df_combined = input_df
                if save_data_to_sheet(df_combined):
                    st.success("🎉 名單成功同步至 Google Sheet 雲端資料庫！請重新整理網頁檢視。")
                    df_crm = df_combined
        except Exception as e:
            st.error(f"Excel 解析失敗，請確認欄位名稱。錯誤: {e}")

    st.write("---")
    st.subheader("🎯 當日待開火與跟催名單")
    
    if not df_crm.empty:
        # 只顯示未回覆的客戶在打擊區，避免畫面過大
        df_active = df_crm[df_crm["已回覆"] == False].copy()
        
        # 建立互動式精簡表格
        edited_df = st.data_editor(
            df_active,
            column_config={
                "進度狀態": st.column_config.SelectboxColumn("進度狀態", options=["準備開發", "已發首封", "已發跟催信", "暫停開發"], required=True),
                "負責業務": st.column_config.SelectboxColumn("負責業務", options=["未指定", "Frances", "Amanda", "Alex", "Jay"], required=True),
                "已回覆": st.column_config.CheckboxColumn("已回覆?", default=False),
                "業務備註": st.column_config.TextColumn("📝 業務備註", width="medium")
            },
            disabled=["公司名稱", "網址", "First Name", "Last Name", "Job Title", "Email", "最後聯絡日期"],
            key="active_editor"
        )
        
        if st.button("💾 儲存今日看板變更至雲端"):
            # 將變更合併回總表
            df_crm.update(edited_df)
            if save_data_to_sheet(df_crm):
                st.success("雲端資料庫已即時更新！")
        
        # 5. 信件發送與 AI 潤稿核心引擎
        def create_draft_via_gmail(to_email, subject, html_content):
            try:
                # 建立加密 HTML 內嵌圖片格式
                msg = MIMEMultipart('related')
                msg['Subject'] = subject
                msg['From'] = GMAIL_ACCOUNT
                msg['To'] = to_email
                
                msgAlternative = MIMEMultipart('alternative')
                msg.attach(msgAlternative)
                msgAlternative.attach(MIMEText(html_content, 'html'))
                
                # 內嵌圖片處理 (報價單)
                if os.path.exists("產品報價.jpg"):
                    with open("產品報價.jpg", 'rb') as f:
                        msgImage = MIMEImage(f.read())
                        msgImage.add_header('Content-ID', '<quote_img>')
                        msg.attach(msgImage)
                
                # 內嵌圖片處理 (產品圖)
                if os.path.exists("真空破除閥.png"):
                    with open("真空破除閥.png", 'rb') as f:
                        msgImage2 = MIMEImage(f.read())
                        msgImage2.add_header('Content-ID', '<product_img>')
                        msg.attach(msgImage2)
                
                # 利用 Gmail IMAP 將郵件寫入草稿匣，而不直接寄出
                import imaplib
                obj = imaplib.IMAP4_SSL('imap.gmail.com', 993)
                obj.login(GMAIL_ACCOUNT, GMAIL_APP_PASSWORD)
                obj.select('[Gmail]/Drafts')
                # 格式轉換
                import time
                obj.append('[Gmail]/Drafts', '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
                obj.logout()
                return True
            except Exception as e:
                st.error(f"草稿塞入公用 Gmail 失敗: {e}")
                return False

        def ask_ai_to_write(company, name, title, stage="first"):
            if not OPENAI_API_KEY:
                return "Subject: Business Cooperation Inquiry", "OpenAI Key 未設定，無法生成客製化內文。"
            
            client = openai.OpenAI(api_key=OPENAI_API_KEY)
            
            # AI 痛點戰略切換指令
            system_prompt = """You are a senior B2B sales director at Geann, a trusted manufacturer in the plumbing industry for decades.
Your goal is to write a highly compelling, personalized introductory hook for an email to a potential client.
You MUST analyze the prospect's Job Title to tailor the pain point:
- If the title relates to Sourcing, Purchasing, Procurement, or Supply Chain: Focus heavily on COST REDUCTION and EFFICIENCY. Mention that switching to Geann's Stainless Steel Vacuum Breaker (US$ 1.08/pc) instead of traditional Lead-Free Brass (US$ 1.38/pc) will save them nearly 20% in material costs without sacrificing quality.
- If the title relates to Engineering, R&D, Quality, or Product Development: Focus heavily on TECHNICAL ADVANTAGES, MATERIAL COMPLIANCE, and ZERO MAINTENANCE. Emphasize that Geann's stainless steel series requires NO ELECTROPLATING, eliminating concerns about coating durability or peeling, and fully complies with ASSE 1011 / CSA B64.2 and NSF61-9 + NSF372 standards.
- If the title is generic, provide a balanced hook covering both cost savings and regulatory compliance.

Output format:
Your output must follow this strict format:
[SUBJECT]
Your personalized, catchy email subject line here
[HOOK]
Dear {First Name},
Your customized opening and hook paragraph here. (Keep it within 3-4 professional and punchy sentences).
"""
            user_content = f"Company: {company}\nContact Name: {name}\nJob Title: {title}\nEmail Stage: {'First Cold Email' if stage=='first' else '7-Day Follow-Up seeking Free Samples'}"
            
            try:
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_content}
                    ],
                    temperature=0.7
                )
                res_text = response.choices[0].message.content
                
                # 解析主旨與 Hook
                subject = "Business Inquiry from Geann"
                hook = f"Dear {name},\nI hope this message finds you well."
                if "[SUBJECT]" in res_text and "[HOOK]" in res_text:
                    parts = res_text.split("[HOOK]")
                    subject = parts[0].replace("[SUBJECT]", "").strip()
                    hook = parts[1].strip()
                return subject, hook
            except Exception as e:
                return "Collaboration Opportunity with Geann", f"Dear {name},\n\nError generating AI hook: {e}"

        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("🔥 啟動第一階段：發動首波開發信 (自動塞入公用 Gmail 草稿)"):
                df_targets = df_crm[(df_crm["進度狀態"] == "準備開發") & (df_crm["已回覆"] == False)]
                if df_targets.empty:
                    st.info("目前沒有「準備開發」的目標客戶。")
                else:
                    success_count = 0
                    for idx, row in df_targets.iterrows():
                        sub, ai_hook = ask_ai_to_write(row["公司名稱"], row["First Name"], row["Job Title"], "first")
                        
                        # 完美拼接 HTML 內文與內嵌圖片
                        html_body = f"""
                        <html>
                        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333333;">
                            <p>{ai_hook.replace('\n', '<br>')}</p>
                            
                            <p>After reviewing your website, we believe that our <b>Vacuum Breakers / Hose Bibbs</b> could be an excellent fit for your product line.<br>
                            Geann has been a trusted manufacturer in the plumbing industry for decades, and we are confident that our products can support your business with both quality and efficiency.</p>
                            
                            <p><b>Why Choose Geann:</b><br>
                            1. <b>Fast & Reliable Supply</b> – We ensure quick response and on-time delivery based on your production schedule.<br>
                            2. <b>Fully Certified</b> – Our products meet <b>ASSE 1011 / CSA B64.2</b> and <b>NSF61-9 + NSF372</b> requirements.<br>
                            3. <b>Guaranteed Compatibility</b> – Designed to match and integrate seamlessly with your existing products.<br>
                            4. <b>Durable Construction</b> – Made with heavy-duty stainless steel / brass and equipped with a durable vacuum breaker mechanism for long-lasting protection and performance.<br>
                            5. <span style="color: blue; font-weight: bold;">Stainless Steel Vacuum Breaker Advantage – Lead-Free & No Plating Required</span><br>
                            &nbsp;&nbsp;&nbsp;&nbsp;We would especially like to highlight our Stainless Steel Vacuum Breaker series, which is a strong alternative to traditional brass designs.<br>
                            &nbsp;&nbsp;&nbsp;&nbsp;• <b>Naturally lead-free material</b>, fully compliant with modern safety and drinking water standards<br>
                            &nbsp;&nbsp;&nbsp;&nbsp;• <b>No electroplating required</b>, eliminating concerns about coating durability or peeling<br>
                            &nbsp;&nbsp;&nbsp;&nbsp;• Excellent <b>corrosion resistance and long service life</b><br>
                            &nbsp;&nbsp;&nbsp;&nbsp;• Ideal for customers seeking a more sustainable, high-end, and maintenance-free solution<br>
                            &nbsp;&nbsp;&nbsp;&nbsp;This makes our stainless steel vacuum breaker a highly competitive option in today’s market.</p>
                            
                            <p>You can find more details on our Vacuum Breaker series here:<br>
                            <a href="https://www.geann.com.tw/en/category/Vacuum-Breaker/CAT-Vacuum-Breaker_VB.html">https://www.geann.com.tw/en/category/Vacuum-Breaker/CAT-Vacuum-Breaker_VB.html</a></p>
                            
                            <p>We would be glad to provide samples, technical files, or further information upon request.<br>
                            Please feel free to contact us anytime—we look forward to the opportunity to work with you.</p>
                            
                            <p><img src="cid:quote_img" alt="Product Quotation" style="max-width:100%; height:auto; margin: 15px 0;"></p>
                            <p><img src="cid:product_img" alt="Geann Vacuum Breaker" style="max-width:100%; height:auto; margin: 15px 0;"></p>
                        </body>
                        </html>
                        """
                        if create_draft_via_gmail(row["Email"], sub, html_body):
                            df_crm.at[idx, "進度狀態"] = "已發首封"
                            df_crm.at[idx, "最後聯絡日期"] = datetime.date.today().strftime("%Y-%m-%d")
                            success_count += 1
                    save_data_to_sheet(df_crm)
                    st.success(f"成功為 {success_count} 位客戶生成客製化首封信草稿，已全部塞入公用 Gmail 草稿匣！")

        with col_btn2:
            if st.button("🦷 啟動第二階段：死咬跟催系統 (AI 自動追蹤續推樣品)"):
                df_follow = df_crm[(df_crm["進度狀態"] == "已發首封") & (df_crm["已回覆"] == False)]
                if df_follow.empty:
                    st.info("目前沒有符合跟催條件的客戶。")
                else:
                    f_count = 0
                    for idx, row in df_follow.iterrows():
                        sub, ai_hook = ask_ai_to_write(row["公司名稱"], row["First Name"], row["Job Title"], "follow")
                        
                        html_follow_body = f"""
                        <html>
                        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333333;">
                            <p>{ai_hook.replace('\n', '<br>')}</p>
                            
                            <p>I wanted to briefly follow up regarding our Vacuum Breakers and Hose Bibbs.<br>
                            As mentioned, Geann has been supplying high-quality plumbing products for decades, and our Vacuum Breakers are fully certified to ASSE 1011 / CSA B64.2 and NSF61-9 + NSF372 standards. We believe they could be a valuable addition to your product line.</p>
                            
                            <p><b>To help you evaluate our products more efficiently, we would be pleased to provide free samples for your review and testing.</b></p>
                            
                            <p><span style="color: blue; font-weight: bold;">Stainless Steel Vacuum Breaker Advantage – Lead-Free & No Plating Required</span><br>
                            We would especially like to highlight our Stainless Steel Vacuum Breaker series, which is a strong alternative to traditional brass designs.<br>
                            • <b>Naturally lead-free material</b>, fully compliant with modern safety and drinking water standards<br>
                            • <b>No electroplating required</b>, eliminating concerns about coating durability or peeling<br>
                            • Excellent <b>corrosion resistance and long service life</b><br>
                            • Ideal for customers seeking a more sustainable, high-end, and maintenance-free solution<br>
                            This makes our stainless steel vacuum breaker a highly competitive option in today’s market.</p>
                            
                            <p>If you are interested, please simply provide:<br>
                            • <u>Your shipping address/contact information</u><br>
                            • <u>Your courier account number (FedEx, DHL, or UPS)</u><br>
                            We will then arrange the sample shipment accordingly.<br>
                            Should the samples meet your expectations, we would be happy to discuss further cooperation opportunities, technical requirements, and pricing details.</p>
                            
                            <p><img src="cid:quote_img" alt="Product Quotation" style="max-width:100%; height:auto; margin: 15px 0;"></p>
                            <p><img src="cid:product_img" alt="Geann Vacuum Breaker" style="max-width:100%; height:auto; margin: 15px 0;"></p>
                            <p>Thank you for your time and consideration.<br>I look forward to hearing from you.</p>
                        </body>
                        </html>
                        """
                        if create_draft_via_gmail(row["Email"], "Follow up: Free Samples available from Geann", html_follow_body):
                            df_crm.at[idx, "進度狀態"] = "已發跟催信"
                            df_crm.at[idx, "最後聯絡日期"] = datetime.date.today().strftime("%Y-%m-%d")
                            f_count += 1
                    save_data_to_sheet(df_crm)
                    st.success(f"成功生成 {f_count} 封「免費樣品跟催信」草稿至公用 Gmail 草稿匣！")

    else:
        st.info("雲端資料庫目前空空如也，請先上傳名單彈藥！")

# ==================== 頁籤 2：CRM 客戶池總覽 ====================
with tab2:
    st.subheader("🗄️ Geann 既有客戶中央雲端記憶池")
    st.write("此處會顯示 Google Sheet 內的所有既有客戶資產（包含開發中與已回覆）。目前第一版保持簡潔總覽表格，方便隨時翻看歷史戰果。")
    
    if not df_crm.empty:
        # 提供簡單的篩選器功能
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filter_sales = st.selectbox("🔍 依業務篩選客戶", options=["全部"] + list(SALES_MAP.keys()))
        with col_f2:
            search_company = st.text_input("🔍 輸入公司名稱搜尋")
        
        df_display = df_crm.copy()
        if filter_sales != "全部":
            df_display = df_display[df_display["負責業務"] == filter_sales]
        if search_company:
            df_display = df_display[df_display["公司名稱"].str.contains(search_company, case=False, na=False)]
            
        st.dataframe(df_display, use_container_width=True)
    else:
        st.info("雲端資料庫無資料。")
