import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime

# --- 1. 系統設定與樣式 ---
st.set_page_config(page_title="產險佣金計算系統", layout="wide")

st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3em; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3em; background-color: #4CAF50; color: white; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 授權登入 (密碼驗證) ---
if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔐 產險佣金管理系統")
    with st.form("login_gate"):
        pwd = st.text_input("請輸入授權密碼", type="password")
        if st.form_submit_button("確認登入"):
            if pwd == "085799": # 沿用你原本的密碼
                st.session_state.auth = True
                st.rerun()
            else:
                st.error("密碼錯誤")
    st.stop()

# --- 3. 佣金率定義 ---
RATES = {
    "車險任意險": 0.07,
    "機車強制險": 0.04,
    "汽車強制險": 60, # 這是固定金額
    "住火險": 0.03,
    "旅平險": 0.10,
    "產品責任險": 0.10
}

# --- 4. 主介面 ---
st.title("💰 佣金對帳與計算中心")

# 側邊欄：一鍵登出
with st.sidebar:
    if st.button("🔓 安全登出"):
        st.session_state.auth = False
        st.rerun()

# --- 5. 檔案上傳區 ---
st.subheader("第一步：上傳原始資料")
col1, col2 = st.columns(2)

with col1:
    prog_file = st.file_uploader("📤 上傳『出單進度表』(Excel)", type="xlsx")
with col2:
    comm_file = st.file_uploader("📤 上傳『佣金表』(Excel)", type="xlsx")

if prog_file and comm_file:
    # 讀取資料
    df_prog = pd.read_excel(prog_file).fillna("")
    df_comm = pd.read_excel(comm_file).fillna("")
    
    # 確保欄位名稱乾淨
    df_prog.columns = [str(c).strip() for c in df_prog.columns]
    df_comm.columns = [str(c).strip() for c in df_comm.columns]

    st.success("檔案上傳成功，開始進行資料比對...")

    # --- 6. 核心比對與計算邏輯 ---
    results = []
    total_raw_comm = 0 # 佣金表中的總應領佣金
    total_calculated_comm = 0 # 我們自己算的總應付佣金 (留凱基)

    # 預先計算總所得稅金 (佣金表 應領佣金 * 5%)
    if "應領佣金" in df_comm.columns:
        total_income_tax = pd.to_numeric(df_comm["應領佣金"], errors='coerce').sum() * 0.05
    else:
        total_income_tax = 0
        st.warning("佣金表中未找到『應領佣金』欄位，稅金計算為 0")

    # 開始逐筆比對
    for _, c_row in df_comm.iterrows():
        name = str(c_row.get("被保險人姓名", "")).strip()
        uid = str(c_row.get("被保險人身分證字號/統一編號", "")).strip()
        policy_no = str(c_row.get("保單號碼", "")).strip()
        
        # 比對進度表 (比對姓名、身分證、保單號碼)
        match = df_prog[
            (df_prog["被保險人姓名"].astype(str).str.strip() == name) &
            (df_prog["被保險人身份證字號/統一編號"].astype(str).str.strip() == uid) &
            (df_prog["新年度保單號碼"].astype(str).str.strip() == policy_no)
        ]
        
        if not match.empty:
            p_row = match.iloc[0]
            servicer = str(p_row.get("實際服務人員", "")).strip()
            
            # 如果沒有實際服務人員則忽略 (依指令要求)
            if not servicer:
                continue
                
            plate = p_row.get("牌照號碼", "")
            
            # 判斷險種與計算佣金
            premium = pd.to_numeric(c_row.get("實收保費", 0), errors='coerce')
            ins_type = str(c_row.get("險種", "")) # 假設佣金表有寫，或從進度表判定
            
            calc_comm = 0
            if "車險" in ins_type and "任意" in ins_type:
                calc_comm = premium * RATES["車險任意險"]
            elif "機車" in ins_type and "強制" in ins_type:
                calc_comm = premium * RATES["機車強制險"]
            elif "汽車" in ins_type and "強制" in ins_type:
                calc_comm = RATES["汽車強制險"]
            elif "火險" in ins_type:
                calc_comm = premium * RATES["住火險"]
            elif "旅平" in ins_type:
                calc_comm = premium * RATES["旅平險"]
            elif "責任" in ins_type:
                calc_comm = premium * RATES["產品責任險"]
            else:
                # 預設邏輯：若無法判定，嘗試從進度表的保險種類判定
                calc_comm = 0 

            results.append({
                "製表日期": datetime.now().strftime("%Y-%m-%d"),
                "被保險人姓名": name,
                "保單號碼": policy_no,
                "車牌號碼": plate,
                "實收保費": premium,
                "應付佣金": round(calc_comm, 0),
                "實際服務人員": servicer
            })
            
            total_calculated_comm += calc_comm
            total_raw_comm += pd.to_numeric(c_row.get("應領佣金", 0), errors='coerce')

    # --- 7. 顯示統計數據 ---
    st.divider()
    st.subheader("📊 結算統計報告")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("總所得稅金 (5%)", f"${total_income_tax:,.0f}")
    m2.metric("留凱基 (應付佣金總計)", f"${total_calculated_comm:,.0f}")
    
    # 匯國泰 = 佣金表總應領 - 留凱基
    remit_cathay = total_raw_comm - total_calculated_comm
    m3.metric("匯國泰", f"${remit_cathay:,.0f}")
    
    # --- 8. 顯示結果表格與下載 ---
    if results:
        res_df = pd.DataFrame(results)
        st.dataframe(res_df, use_container_width=True)
        
        # 產出 Excel 下載
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            res_df.to_excel(writer, index=False, sheet_name='佣金紀錄明細')
            # 在旁邊記錄統計數據
            summary_df = pd.DataFrame([{
                "總所得稅金": total_income_tax,
                "留凱基": total_calculated_comm,
                "匯國泰": remit_cathay
            }])
            summary_df.to_excel(writer, index=False, sheet_name='匯總摘要')
        
        st.download_button(
            label="📥 下載佣金結算 Excel 表格",
            data=output.getvalue(),
            file_name=f"佣金對帳單_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("查無符合比對條件（姓名+ID+保單號碼且有服務人員）的資訊。")