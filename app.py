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

# 🔐 新增密碼登入驗證機制
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
        
    if st.session_state["password_correct"]:
        return True

    st.markdown("<h2 style='text-align: center; color: #1E3A8A;'>🔒 Geann B2B 戰情室安全登入</h2>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        password = st.text_input("請輸入團隊通關密碼：", type="password")
        if st.button("確認登入"):
            correct_password = st.secrets.get("LOGIN_PASSWORD", "GeannAdmin")
            if password == correct_password:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("❌ 密碼錯誤，請重新輸入！")
    return False

if not check_password():
    st.stop()

st.markdown("""
    <style>
    .main-title { font-size:32px; font-weight:bold; color:#1E3A8A; margin-bottom:10px; }
    .stButton>button { width: 100%; border-radius: 8px; background-color: #1E3A8A; color: white; }
    .stButton>button:hover { background-color: #2563EB; color: white; }
    </style>
    <div class="main-title">🚀 Geann B2B 業務戰情監控版 (完全體)</div>
""", unsafe_allow_html=True)

# 2. 安全讀取保險箱後台金鑰
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", "")
GMAIL_ACCOUNT = st.secrets.get("GMAIL_ACCOUNT", "")
GMAIL_APP_PASSWORD = st.secrets.get("GMAIL_APP_PASSWORD", "")
SHEET_URL = "https://docs.google.com/spreadsheets/d/1j10PqYMrg09T-N21UCrnZs7UVqXb1ZmglRgjZmxcbgs/edit?gid=0#gid=0"

SALES_MAP = {
    "Frances": "francesma@geann.com.tw",
    "Amanda": "amandalee@geann.com.tw",
    "Alex": "alexwang@geann.com.tw",
    "Jay": "Jayhuang@geann.com.tw"
}

@st.cache_resource
def init_gspread():
    try:
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        gcp_sa = dict(st.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(gcp_sa, scopes=scope)
        return gspread.authorize(credentials)
    except Exception as e:
        st.error(f"Google 試算表連線失敗: {e}")
        return None

gc = init_gspread()

def load_data_from_sheet():
    if gc:
        try:
            sh = gc.open_by_url(SHEET_URL)
            worksheet = sh.get_worksheet(0)
            data = worksheet.get_all_records()
            if not data:
                return pd.DataFrame(columns=["進度狀態", "已回覆", "業務備註", "最後聯絡日期", "負責業務", "公司名稱", "網址", "First Name", "Last Name", "Job Title", "Email", "標記刪除"])
            df = pd.DataFrame(data)
            # 🐛 關鍵修復：嚴格判定布林值，確保未回覆的客戶是 False，不會被隱藏
            df["已回覆"] = df["已回覆"].astype(str).str.strip().str.upper() == "TRUE"
            
            # 新增刪除標記欄位（若雲端沒有的話）
            if "標記刪除" not in df.columns:
                df["標記刪除"] = False
            else:
                df["標記刪除"] = df["標記刪除"].astype(str).str.strip().str.upper() == "TRUE"
                
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
            cols = ["進度狀態", "已回覆", "業務備註", "最後聯絡日期", "負責業務", "公司名稱", "網址", "First Name", "Last Name", "Job Title", "Email", "標記刪除"]
            
            # 確保缺少欄位時補齊
            for col in cols:
                if col not in df.columns:
                    df[col] = ""
                    
            df_to_save = df[cols].copy()
            df_to_save["已回覆"] = df_to_save["已回覆"].astype(bool).astype(str).str.upper()
            df_to_save["標記刪除"] = df_to_save["標記刪除"].astype(bool).astype(str).str.upper()
            worksheet.update([df_to_save.columns.values.tolist()] + df_to_save.values.tolist())
            return True
        except Exception as e:
            st.error(f"寫入雲端資料失敗: {e}")
    return False

df_crm = load_data_from_sheet()

tab1, tab2 = st.tabs(["🚀 名單開火與跟催區", "🗄️ CRM 客戶池總覽"])

# ==================== 頁籤 1：開火與跟催區 ====================
with tab1:
    st.subheader("📥 匯入 Apollo 尋人名單 (新彈藥)")
    uploaded_file = st.file_uploader("請上傳清洗後的 Keyman 名單 (Excel)", type=["xlsx"])
    
    if uploaded_file:
        try:
            input_df = pd.read_excel(uploaded_file)
            input_df = input_df.fillna("") # 清洗空白
            
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
            if "標記刪除" not in input_df.columns:
                input_df["標記刪除"] = False
                
            cols = ["進度狀態", "已回覆", "業務備註", "最後聯絡日期", "負責業務", "公司名稱", "網址", "First Name", "Last Name", "Job Title", "Email", "標記刪除"]
            input_df = input_df[cols]
            
            if st.button("🔥 確認將新名單匯入雲端資料庫"):
                if not df_crm.empty:
                    df_combined = pd.concat([df_crm, input_df]).drop_duplicates(subset=["Email"], keep="last")
                else:
                    df_combined = input_df
                if save_data_to_sheet(df_combined):
                    st.success("🎉 匯入成功！請重新整理網頁。")
                    st.rerun()
        except Exception as e:
            st.error(f"Excel 解析失敗: {e}")

    st.write("---")
    st.subheader("🎯 當日待開火與跟催名單 (✅ 可在此處直接下拉指定業務)")
    
    if not df_crm.empty:
        # 只顯示未回覆且未標記刪除的客戶
        df_active = df_crm[(df_crm["已回覆"] == False) & (df_crm["標記刪除"] == False)].copy()
        
        edited_df = st.data_editor(
            df_active,
            column_config={
                "進度狀態": st.column_config.SelectboxColumn("進度狀態", options=["準備開發", "已發首封", "已發跟催信", "暫停開發"]),
                "負責業務": st.column_config.SelectboxColumn("負責業務", options=["未指定", "Frances", "Amanda", "Alex", "Jay"]),
                "已回覆": st.column_config.CheckboxColumn("已回覆?", default=False),
                "業務備註": st.column_config.TextColumn("📝 業務備註", width="medium"),
                "標記刪除": None # 在打擊區隱藏刪除按鈕，保持版面乾淨
            },
            disabled=["公司名稱", "網址", "First Name", "Last Name", "Job Title", "Email", "最後聯絡日期"],
            key="active_editor"
        )
        
        if st.button("💾 儲存今日看板變更至雲端"):
            df_crm.update(edited_df)
            if save_data_to_sheet(df_crm):
                st.success("雲端資料庫已即時更新！")
                st.rerun()
        
        # --- 信件發送模組 ---
        # --- 信件發送模組 (含雷達偵測與 Debug 功能) ---
        def create_draft_via_gmail(to_email, subject, html_content):
            try:
                msg = MIMEMultipart('related')
                msg['Subject'] = subject
                msg['From'] = GMAIL_ACCOUNT
                msg['To'] = to_email
                
                msgAlternative = MIMEMultipart('alternative')
                msg.attach(msgAlternative)
                msgAlternative.attach(MIMEText(html_content, 'html'))
                
                if os.path.exists("產品報價.jpg"):
                    with open("產品報價.jpg", 'rb') as f:
                        msgImage = MIMEImage(f.read())
                        msgImage.add_header('Content-ID', '<quote_img>')
                        msg.attach(msgImage)
                
                if os.path.exists("真空破除閥.png"):
                    with open("真空破除閥.png", 'rb') as f:
                        msgImage2 = MIMEImage(f.read())
                        msgImage2.add_header('Content-ID', '<product_img>')
                        msg.attach(msgImage2)
                
                import imaplib, time
                # 連線伺服器
                obj = imaplib.IMAP4_SSL('imap.gmail.com', 993)
                obj.login(GMAIL_ACCOUNT, GMAIL_APP_PASSWORD)
                
                # 🔍 雷達偵測：找出真正對應「草稿」的系統資料夾名稱
                typ, data = obj.list()
                draft_folder = '"[Gmail]/Drafts"' # 預設值
                for box in data:
                    box_str = box.decode('utf-8')
                    # 尋找帶有 Drafts 標籤的資料夾
                    if '\\Drafts' in box_str:
                        # 擷取 Gmail 回傳格式中的資料夾名稱
                        parts = box_str.split(' "/" ')
                        if len(parts) >= 2:
                            draft_folder = parts[1].strip()
                        break
                
                # 💡 在畫面上印出系統找到了哪個資料夾，讓你安心
                st.info(f"🔍 [Debug 追蹤] 系統偵測到你的真實草稿匣路徑為: {draft_folder}")
                
                # 切換到該資料夾並塞入信件
                obj.select(draft_folder)
                status, response = obj.append(draft_folder, '(\\Draft)', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
                
                obj.logout()
                
                # 嚴格確認是否真的寫入成功
                if status == 'OK':
                    return True
                else:
                    st.error(f"❌ 寫入失敗，伺服器回傳: {response}")
                    return False
                    
            except Exception as e:
                st.error(f"❌ 草稿塞入發生嚴重錯誤: {e}")
                return False

        def ask_ai_to_write(company, name, title, stage="first"):
            if not OPENAI_API_KEY:
                return "Subject: Business Cooperation Inquiry", "OpenAI Key 未設定。"
            
            client = openai.OpenAI(api_key=OPENAI_API_KEY)
            system_prompt = """You are a senior B2B sales director at Geann.
Your output must follow this strict format:
[SUBJECT]
Your personalized, catchy email subject line here
[HOOK]
Dear {First Name},
Your customized opening and hook paragraph here.
"""
            user_content = f"Company: {company}\nContact Name: {name}\nJob Title: {title}\nEmail Stage: {'First Cold Email' if stage=='first' else '7-Day Follow-Up'}"
            
            try:
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "system", "content": system_prompt}, {"role": "user", "content": user_content}],
                    temperature=0.7
                )
                res_text = response.choices[0].message.content
                subject = "Business Inquiry from Geann"
                hook = f"Dear {name},\nI hope this message finds you well."
                if "[SUBJECT]" in res_text and "[HOOK]" in res_text:
                    parts = res_text.split("[HOOK]")
                    subject = parts[0].replace("[SUBJECT]", "").strip()
                    hook = parts[1].strip()
                return subject, hook
            except Exception as e:
                return "Collaboration Opportunity with Geann", f"Dear {name},\n\nError generating AI hook: {e}"

        st.write("---")
        # 🧪 新增：單筆測試與全體發送按鈕並列
        col_test, col_batch = st.columns(2)
        
        with col_test:
            if st.button("🧪 測試發射 (僅針對列表第一位客戶產生一封草稿)"):
                df_targets = df_crm[(df_crm["進度狀態"] == "準備開發") & (df_crm["已回覆"] == False) & (df_crm["標記刪除"] == False)]
                if df_targets.empty:
                    st.info("目前沒有「準備開發」的客戶可供測試。")
                else:
                    row = df_targets.iloc[0] # 只抓第一筆
                    idx = df_targets.index[0]
                    sub, ai_hook = ask_ai_to_write(row["公司名稱"], row["First Name"], row["Job Title"], "first")
                    
                    html_body = f"""
                    <html>
                    <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333333;">
                        <p>{ai_hook.replace('\n', '<br>')}</p>
                        <p>After reviewing your website, we believe that our <b>Vacuum Breakers / Hose Bibbs</b> could be an excellent fit for your product line.<br>
                        Geann has been a trusted manufacturer in the plumbing industry for decades, and we are confident that our products can support your business with both quality and efficiency.</p>
                        <p><img src="cid:quote_img" alt="Product Quotation" style="max-width:100%; height:auto; margin: 15px 0;"></p>
                        <p><img src="cid:product_img" alt="Geann Vacuum Breaker" style="max-width:100%; height:auto; margin: 15px 0;"></p>
                    </body>
                    </html>
                    """
                    if create_draft_via_gmail(row["Email"], sub, html_body):
                        df_crm.at[idx, "進度狀態"] = "已發首封"
                        df_crm.at[idx, "最後聯絡日期"] = datetime.date.today().strftime("%Y-%m-%d")
                        save_data_to_sheet(df_crm)
                        st.success(f"✅ 測試成功！已為 {row['公司名稱']} 生成一封草稿，並將狀態改為已發首封。請至公用 Gmail 草稿匣查看！")
                        st.rerun()

        with col_batch:
            if st.button("🔥 正式啟動：為所有「準備開發」客戶產生草稿"):
                df_targets = df_crm[(df_crm["進度狀態"] == "準備開發") & (df_crm["已回覆"] == False) & (df_crm["標記刪除"] == False)]
                if df_targets.empty:
                    st.info("目前沒有「準備開發」的目標客戶。")
                else:
                    success_count = 0
                    for idx, row in df_targets.iterrows():
                        sub, ai_hook = ask_ai_to_write(row["公司名稱"], row["First Name"], row["Job Title"], "first")
                        html_body = f"<html><body><p>{ai_hook.replace('\n', '<br>')}</p><p>Geann Vacuum Breakers.</p><p><img src='cid:quote_img'></p><p><img src='cid:product_img'></p></body></html>"
                        if create_draft_via_gmail(row["Email"], sub, html_body):
                            df_crm.at[idx, "進度狀態"] = "已發首封"
                            df_crm.at[idx, "最後聯絡日期"] = datetime.date.today().strftime("%Y-%m-%d")
                            success_count += 1
                    save_data_to_sheet(df_crm)
                    st.success(f"成功生成 {success_count} 封客製化首封信草稿！")
                    st.rerun()

# ==================== 頁籤 2：CRM 客戶池總覽 ====================
with tab2:
    st.subheader("🗄️ Geann 既有客戶中央雲端記憶池 (含刪除管理)")
    
    if not df_crm.empty:
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            filter_sales = st.selectbox("🔍 依業務篩選客戶", options=["全部"] + list(SALES_MAP.keys()))
        with col_f2:
            search_company = st.text_input("🔍 輸入公司名稱搜尋")
        
        # 預設隱藏已經被真正刪除的資料，但保留打勾標記的讓他們編輯
        df_display = df_crm.copy()
        if filter_sales != "全部":
            df_display = df_display[df_display["負責業務"] == filter_sales]
        if search_company:
            df_display = df_display[df_display["公司名稱"].str.contains(search_company, case=False, na=False)]
            
        st.write("💡 **剔除客戶教學**：在最右側的「🗑️ 標記刪除」打勾，然後點擊下方按鈕，該客戶就會從系統與 Google Sheet 中被永久移除。")
        
        # 允許在此處編輯標記刪除
        edited_crm = st.data_editor(
            df_display,
            column_config={
                "標記刪除": st.column_config.CheckboxColumn("🗑️ 標記刪除", default=False),
                "進度狀態": st.column_config.SelectboxColumn("進度狀態", options=["準備開發", "已發首封", "已發跟催信", "暫停開發"]),
                "負責業務": st.column_config.SelectboxColumn("負責業務", options=["未指定", "Frances", "Amanda", "Alex", "Jay"])
            },
            key="crm_total_editor"
        )
        
        if st.button("🗑️ 執行刪除與儲存 CRM 變更"):
            # 將變更寫回主資料庫
            df_crm.update(edited_crm)
            
            # 過濾掉被標記刪除的行
            df_crm = df_crm[df_crm["標記刪除"] == False]
            
            if save_data_to_sheet(df_crm):
                st.success("變更與刪除已同步至雲端資料庫！")
                st.rerun()
    else:
        st.info("雲端資料庫無資料。")
   
