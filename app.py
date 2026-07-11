import streamlit as st
import pandas as pd
import imaplib
import email.message
import time
import io
from openai import OpenAI

# --- 🔒 系統金鑰與環境設定 ---
st.set_page_config(page_title="Geann 業務戰情監控版", layout="wide")
GMAIL_ACCOUNT = st.secrets.get("GMAIL_ACCOUNT", "")
GMAIL_APP_PASSWORD = st.secrets.get("GMAIL_APP_PASSWORD", "")
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY", "")

# --- 📧 核心功能：將信件無聲投遞至 Gmail 草稿匣 ---
def save_to_gmail_drafts(subject, body, to_email):
    try:
        # 連線至 Gmail IMAP 伺服器
        conn = imaplib.IMAP4_SSL('imap.gmail.com')
        conn.login(GMAIL_ACCOUNT, GMAIL_APP_PASSWORD)
        
        # 建立信件結構
        msg = email.message.EmailMessage()
        msg['Subject'] = subject
        msg['From'] = GMAIL_ACCOUNT
        msg['To'] = to_email
        msg.set_content(body)
        
        # 附加到草稿匣 (Drafts)
        # 備註：若 Gmail 語言為中文，資料夾名稱可能是 '[Gmail]/草稿'
        conn.append('[Gmail]/Drafts', '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        conn.logout()
        return True
    except Exception as e:
        st.error(f"投遞草稿失敗: {e}")
        return False

# --- 🤖 核心功能：AI 撰寫信件 (等待老大提供 Template 後解鎖) ---
def generate_ai_email(name, title, company, email_type="first"):
    # 這裡將接上 OpenAI API，目前先輸出測試用字串
    if email_type == "first":
        subject = f"Cost Reduction Opportunity on Valve Components for {company}"
        body = f"Hi {name},\n\n[這裡是 AI 根據 {title} 變形的『首封開發信』]\n\nBest regards,\nGeann Team"
    else:
        subject = f"Follow-up: Valve Components for {company}"
        body = f"Hi {name},\n\n[這裡是 AI 根據 {title} 變形的『7天跟催信』]\n\nBest regards,\nGeann Team"
    return subject, body

# --- 📊 網頁主視覺與操作介面 ---
st.title("🚀 Geann B2B 業務戰情監控版 (MVP)")
st.markdown("匯入名單 ➡️ AI 自動寫信存草稿 ➡️ 業務勾選追蹤 ➡️ 自動跟催")

# 1. 檔案上傳區
uploaded_file = st.file_uploader("📥 請上傳清洗後的 Keyman 名單 (Excel)", type=['xlsx'])

if uploaded_file:
    # 讀取並初始化狀態
    if 'df' not in st.session_state:
        df = pd.read_excel(uploaded_file)
        # 初始化業務追蹤欄位
        if "狀態" not in df.columns:
            df.insert(0, "狀態", "準備開發")
        if "已回覆" not in df.columns:
            df.insert(1, "已回覆", False)
        if "業務備註" not in df.columns:
            df.insert(2, "業務備註", "")
        st.session_state.df = df
    
    st.success("✅ 名單匯入成功！請在下方看板直接修改狀態或備註。")

    # 2. 業務動態互動看板 (可直接打勾、改字)
    edited_df = st.data_editor(
        st.session_state.df,
        column_config={
            "狀態": st.column_config.SelectboxColumn(
                "進度狀態",
                options=["準備開發", "草稿已建立", "已發首封", "已發跟催信", "暫停跟進"],
                required=True
            ),
            "已回覆": st.column_config.CheckboxColumn("客戶已回覆?", default=False),
            "業務備註": st.column_config.TextColumn("📝 業務備註 (如：不買/太貴)")
        },
        use_container_width=True,
        num_rows="dynamic"
    )
    st.session_state.df = edited_df

    st.divider()

    # 3. 戰術執行區
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 🎯 第一階段：發動首波開發")
        if st.button("⚡ 自動生成『首封開發信』並投遞至 Gmail 草稿"):
            with st.spinner("AI 正在為『準備開發』的客戶量身撰寫信件..."):
                count = 0
                for idx, row in edited_df.iterrows():
                    if row['狀態'] == '準備開發' and pd.notna(row['Email']):
                        # AI 生成信件
                        subj, body = generate_ai_email(row['First Name'], row['Job Title'], row['公司名稱'], "first")
                        # 投遞草稿
                        if save_to_gmail_drafts(subj, body, row['Email']):
                            st.session_state.df.at[idx, '狀態'] = '草稿已建立'
                            count += 1
                st.success(f"🎉 成功建立 {count} 封客製化開發信草稿！請至 Gmail 查看並發送。")
                st.rerun()

    with col2:
        st.markdown("### 🔄 第二階段：死咬跟催系統")
        if st.button("🔥 自動生成『跟催信 (Follow-up)』並投遞草稿"):
            with st.spinner("AI 正在為未回覆的客戶撰寫跟催信..."):
                count = 0
                for idx, row in edited_df.iterrows():
                    # 邏輯：狀態是已發首封，且業務「沒有」勾選已回覆
                    if row['狀態'] == '已發首封' and not row['已回覆']:
                        subj, body = generate_ai_email(row['First Name'], row['Job Title'], row['公司名稱'], "followup")
                        if save_to_gmail_drafts(subj, body, row['Email']):
                            st.session_state.df.at[idx, '狀態'] = '草稿已建立' # 業務確認後再手動改為'已發跟催信'
                            count += 1
                st.success(f"🎉 成功建立 {count} 封跟催信草稿！請至 Gmail 查看並發送。")
                st.rerun()

    # 下載最新進度
    st.divider()
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.df.to_excel(writer, index=False)
    output.seek(0)
    st.download_button(label="💾 下載今日最新戰情進度 (Excel)", data=output, file_name="Geann_戰情追蹤表.xlsx")
