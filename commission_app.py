import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime

# --- 1. 系統設定與手機版樣式 ---
st.set_page_config(page_title="佣金計算系統", layout="wide")

st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3.5em; margin-bottom: 10px; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3.5em; background-color: #4CAF50; color: white; }
    div[data-testid="stForm"] button[kind="secondary"] { color: white; background-color: #ff4b4b; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 授權登入 ---
if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔐 產險佣金管理系統")
    with st.form("login_gate"):
        pwd = st.text_input("請輸入授權密碼", type="password")
        if st.form_submit_button("確認登入", use_container_width=True):
            if pwd == "085799": 
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("密碼錯誤")
    st.stop()

# --- 3. 佣金利率定義 (依指令設定) ---
# 車險任意險: 7%, 機車強制: 4%, 汽車強制: 60元, 住火: 3%, 旅平: 10%, 責任險: 10%
RATES = {
    "車險任意險": 0.07,
    "機車強制險": 0.04,
    "汽車強制險": 60, 
    "住火險": 0.03,
    "旅平險": 0.10,
    "產品責任險": 0.10
}

# --- 4. 主介面 ---
st.title("💰 佣金對帳與計算中心")

with st.sidebar:
    if st.button("🔓 安全登出系統"):
        st.session_state.auth = False
        st.rerun()

# --- 5. 檔案上傳與清洗 ---
st.subheader("第一步：上傳原始資料")
col1, col2 = st.columns(2)

with col1:
    prog_file = st.file_uploader("📤 上傳『出單進度表』(Excel)", type="xlsx")
with col2:
    comm_file = st.file_uploader("📤 上傳『佣金表』(Excel)", type="xlsx")

if prog_file and comm_file:
    # 讀取並清洗欄位，避免空格與換行導致抓不到「應領佣金」
    df_prog = pd.read_excel(prog_file).fillna("")
    df_comm = pd.read_excel(comm_file).fillna("")
    
    df_prog.columns = [str(c).replace('\n', '').replace(' ', '').strip() for c in df_prog.columns]
    df_comm.columns = [str(c).replace('\n', '').replace(' ', '').strip() for c in df_comm.columns]

    # 檢查關鍵欄位是否存在
    required_comm = ["被保險人姓名", "被保險人身分證字號/統一編號", "保單號碼", "實收保費", "應領佣金"]
    missing = [c for c in required_comm if c not in df_comm.columns]
    
    if missing:
        st.error(f"❌ 佣金表缺少必要欄位：{missing}")
        st.info(f"目前偵測到的欄位有：{list(df_comm.columns)}")
        st.stop()

    # --- 6. 核心運算邏輯 ---
    results = []
    total_raw_comm_sum = 0  # 佣金表中的總「應領佣金」

    # 計算總所得稅金 (佣金表總「應領佣金」 * 5%)
    total_raw_comm_sum = pd.to_numeric(df_comm["應領佣金"], errors='coerce').sum()
    total_income_tax = total_raw_comm_sum * 0.05

    # 逐筆比對
    for i, c_row in df_comm.iterrows():
        name = str(c_row["被保險人姓名"]).strip()
        uid = str(c_row["被保險人身分證字號/統一編號"]).strip()
        p_no = str(c_row["保單號碼"]).strip()
        
        # 比對進度表 (姓名 + ID + 保單號碼)
        match = df_prog[
            (df_prog["被保險人姓名"].astype(str).str.strip() == name) &
            (df_prog["被保險人身份證字號/統一編號"].astype(str).str.strip() == uid) &
            (df_prog["新年度保單號碼"].astype(str).str.strip() == p_no)
        ]
        
        if not match.empty:
            p_row = match.iloc[0]
            servicer = str(p_row.get("實際服務人員", "")).strip()
            
            # 條件：必須有「實際服務人員」才記錄
            if servicer:
                plate = p_row.get("牌照號碼", "")
                premium = pd.to_numeric(c_row["實收保費"], errors='coerce')
                
                # 判定險種 (優先抓佣金表的「險種」或「保件種類」欄位內容)
                # 這裡會掃描所有可能含有險種名稱的字串
                ins_desc = str(c_row.get("險種", "")) + str(c_row.get("保件種類", "")) + p_no
                
                calc_comm = 0
                if "機車" in ins_desc and "強制" in ins_desc:
                    calc_comm = premium * RATES["機車強制險"]
                elif "汽車" in ins_desc and "強制" in ins_desc:
                    calc_comm = RATES["汽車強制險"]
                elif "車" in ins_desc and ("任意" in ins_desc or "乙" in ins_desc or "丙" in ins_desc):
                    calc_comm = premium * RATES["車險任意險"]
                elif "火" in ins_desc:
                    calc_comm = premium * RATES["住火險"]
                elif "旅" in ins_desc:
                    calc_comm = premium * RATES["旅平險"]
                elif "責任" in ins_desc:
                    calc_comm = premium * RATES["產品責任險"]
                else:
                    # 預設：若無法辨識，嘗試根據進度表的「保險種類」
                    p_type = str(p_row.get("保險種類", ""))
                    if "車" in p_type: calc_comm = premium * RATES["車險任意險"]
                    else: calc_comm = 0

                results.append({
                    "製表日期": datetime.now().strftime("%Y-%m-%d"),
                    "被保險人姓名": name,
                    "保單號碼": p_no,
                    "車牌號碼": plate,
                    "實收保費": premium,
                    "應付佣金": round(calc_comm, 0),
                    "實際服務人員": servicer
                })

    # --- 7. 製表與統計分析 ---
    if results:
        res_df = pd.DataFrame(results)
        total_calculated_comm = res_df["應付佣金"].sum() # 留凱基 (應付佣金總額)
        remit_cathay = total_raw_comm_sum - total_calculated_comm # 匯國泰
        
        st.divider()
        st.subheader("📋 結算結果")
        
        # 顯示三大金流指標
        c1, c2, c3 = st.columns(3)
        c1.metric("📌 總所得稅金 (5%)", f"${total_income_tax:,.0f}")
        c2.metric("🏦