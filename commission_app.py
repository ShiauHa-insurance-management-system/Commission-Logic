import streamlit as st
import pandas as pd
import io
import math  # 用於處理無條件進位邏輯
from datetime import datetime

# --- 1. 系統設定 ---
st.set_page_config(page_title="產險佣金計算系統", layout="wide")

# 套用自定義 CSS 樣式
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3.5em; margin-bottom: 10px; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3.5em; background-color: #4CAF50; color: white; border: none; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. 授權登入 ---
if "auth" not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    st.title("🔐 產險佣金管理系統")
    with st.form("login_gate"):
        pwd = st.text_input("請輸入授權密碼", type="password")
        if st.form_submit_button("確認登入"):
            if pwd == "085799": 
                st.session_state.auth = True
                st.rerun()
            else: st.error("密碼錯誤")
    st.stop()

# --- 3. 佣金利率定義 ---
RATES = {
    "車險任意險": 0.07, "機車強制險": 30, "汽車強制險": 60, 
    "住火險": 0.03, "旅平險": 0.10, "產品責任險": 0.10
}

st.title("💰 佣金對帳與計算中心")

# --- 4. 智慧讀取 Excel 函式 ---
def clean_cols(df):
    df.columns = [str(c).replace('\n', '').replace(' ', '').strip() for c in df.columns]
    return df

st.subheader("第一步：上傳原始資料")
col1, col2 = st.columns(2)
with col1:
    prog_file = st.file_uploader("📤 上傳『出單進度表』", type="xlsx")
with col2:
    comm_file = st.file_uploader("📤 上傳『佣金表』", type="xlsx")

if prog_file and comm_file:
    # 讀取進度表
    df_prog = pd.read_excel(prog_file).fillna("")
    df_prog = clean_cols(df_prog)
    
    # 讀取佣金表（標題在第 3 行，對應 header=2）
    df_comm = pd.read_excel(comm_file, header=2).fillna("")
    df_comm = clean_cols(df_comm)

    if "應領佣金" not in df_comm.columns:
        st.error(f"❌ 佣金表定位失敗。看到的標題有：{list(df_comm.columns)}")
        st.stop()

    # --- 5. 核心運算邏輯 ---
    results = []
    
    # 計算總所得稅金
    # 1. 先加總佣金表中的「應領佣金」欄位
    raw_comm_total = pd.to_numeric(df_comm["應領佣金"], errors='coerce').sum()
    
    # 2. 【核心修改】：計算 5% 稅金並執行「小數點第一位無條件進位」取整數
    # math.ceil(87.96) -> 88
    total_income_tax = math.ceil(raw_comm_total * 0.05)

    for i, c_row in df_comm.iterrows():
        # 抓取姓名與保單號碼進行比對
        c_name = str(c_row.get("被保險人", c_row.get("被保險人姓名", ""))).strip()
        c_p_no = str(c_row.get("保單號碼", c_row.get("新年度保單號碼", ""))).strip()
        
        if not c_name or c_name == "nan" or "備註" in c_name: continue

        # 執行雙重欄位比對
        match = df_prog[
            ((df_prog.get("被保險人姓名", pd.Series()).astype(str).str.strip() == c_name) |
             (df_prog.get("被保險人", pd.Series()).astype(str).str.strip() == c_name)) &
            (df_prog.get("新年度保單號碼", pd.Series()).astype(str).str.strip() == c_p_no)
        ]
        
        if not match.empty:
            p_row = match.iloc[0]
            servicer = str(p_row.get("實際服務人員", "")).strip()
            
            if servicer:
                plate = p_row.get("牌照號碼", "")
                premium = pd.to_numeric(c_row.get("實收保費", 0), errors='coerce')
                ins_desc = str(c_row.get("險種", "")) + str(p_row.get("保險種類", ""))
                
                # 計算應付佣金邏輯
                calc_comm = 0
                if "機車" in ins_desc and "強制" in ins_desc:
                    calc_comm = premium * RATES["機車強制險"]
                elif "汽車" in ins_desc and "強制" in ins_desc:
                    calc_comm = RATES["汽車強制險"]
                elif any(k in ins_desc for k in ["車", "任意", "CTA", "QTHO"]):
                    calc_comm = premium * RATES["車險任意險"]
                elif "火" in ins_desc:
                    calc_comm = premium * RATES["住火險"]
                elif "旅" in ins_desc:
                    calc_comm = premium * RATES["旅平險"]
                elif "責任" in ins_desc:
                    calc_comm = premium * RATES["產品責任險"]
                
                results.append({
                    "製表日期": datetime.now().strftime("%Y-%m-%d"),
                    "被保險人姓名": c_name,
                    "保單號碼": c_p_no,
                    "車牌號碼": plate,
                    "實收保費": premium,
                    "應付佣金": round(calc_comm, 0),
                    "實際服務人員": servicer
                })

    # --- 6. 數據匯總與顯示 ---
    if results:
        res_df = pd.DataFrame(results)
        total_calc_comm = res_df["應付佣金"].sum() # 留凱基
        remit_cathay = raw_comm_total - total_calc_comm # 匯國泰
        
        st.divider()
        st.subheader("📊 結算數據統計")
        m1, m2, m3 = st.columns(3)
        # 顯示整數格式
        m1.metric("📌 總所得稅金 (5% 無條件進位)", f"${int(total_income_tax):,}")
        m2.metric("🏦 留凱基 (應付佣金總額)", f"${int(total_calc_comm):,}")
        m3.metric("🏦 匯國泰 (剩餘金額)", f"${int(remit_cathay):,}")
        
        st.dataframe(res_df, use_container_width=True)
        
        # 匯出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            res_df.to_excel(writer, index=False, sheet_name='佣金明細')
            pd.DataFrame([{
                "總所得稅金(無條件進位)": total_income_tax, 
                "應付佣金總計(留凱基)": total_calc_comm, 
                "應匯回金額(匯國泰)": remit_cathay
            }]).to_excel(writer, index=False, sheet_name='統計摘要')
        
        st.download_button(
            label="📥 下載結算報表", 
            data=output.getvalue(), 
            file_name=f"佣金對帳單_{datetime.now().strftime('%m%d')}.xlsx"
        )
    else:
        st.warning("⚠️ 比對完成，但查無符合條件的資料。")