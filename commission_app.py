import streamlit as st
import pandas as pd
import io
import math  # 用於處理進位(ceil)與捨去(floor)
from datetime import datetime

# --- 1. 系統設定 ---
st.set_page_config(page_title="產險佣金計算系統", layout="wide")

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
    "車險任意險": 0.07, "機車強制險": 0.04, "汽車強制險": 60, 
    "住火險": 0.03, "旅平險": 0.10, "產品責任險": 0.10
}

st.title("💰 佣金對帳與計算中心")

# --- 4. 智慧讀取 Excel ---
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
    df_prog = pd.read_excel(prog_file).fillna("")
    df_prog = clean_cols(df_prog)
    
    df_comm = pd.read_excel(comm_file, header=2).fillna("")
    df_comm = clean_cols(df_comm)

    if "應領佣金" not in df_comm.columns:
        st.error(f"❌ 佣金表標題定位失敗。")
        st.stop()

    # --- 5. 核心運算邏輯 ---
    results = []
    # 總應領佣金加總
    raw_comm_total = pd.to_numeric(df_comm["應領佣金"], errors='coerce').sum()
    
    # 總所得稅金：5% 且無條件進位 (math.ceil)
    total_income_tax = math.ceil(raw_comm_total * 0.05)

    for i, c_row in df_comm.iterrows():
        c_name = str(c_row.get("被保險人", c_row.get("被保險人姓名", ""))).strip()
        c_p_no = str(c_row.get("保單號碼", c_row.get("新年度保單號碼", ""))).strip()
        
        if not c_name or c_name == "nan" or "備註" in c_name: continue

        # 比對進度表
        match = df_prog[
            ((df_prog.get("被保險人姓名", pd.Series()).astype(str).str.strip() == c_name) |
             (df_prog.get("被保險人", pd.Series()).astype(str).str.strip() == c_name)) &
            (df_prog.get("新年度保單號碼", pd.Series()).astype(str).str.strip() == c_p_no)
        ]
        
        if not match.empty:
            p_row = match.iloc[0]
            servicer = str(p_row.get("實際服務人員", "")).strip()
            
            # 過濾「謝騏鴻」
            if servicer == "謝騏鴻":
                continue
            
            if servicer:
                plate = p_row.get("牌照號碼", "")
                premium = pd.to_numeric(c_row.get("實收保費", 0), errors='coerce')
                ins_desc = str(c_row.get("險種", "")) + str(p_row.get("保險種類", ""))
                
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

                # 【關鍵修正】：應付佣金採「無條件捨去」(math.floor)
                # 例如 21.6 元會變成 21 元
                final_comm = math.floor(calc_comm)

                results.append({
                    "製表日期": datetime.now().strftime("%Y-%m-%d"),
                    "被保險人姓名": c_name,
                    "保單號碼": c_p_no,
                    "車牌號碼": plate,
                    "實收保費": premium,
                    "應付佣金": final_comm,
                    "實際服務人員": servicer
                })

    # --- 6. 結果顯示與下載 ---
    if results:
        res_df = pd.DataFrame(results)
        total_calc_comm = res_df["應付佣金"].sum()
        remit_cathay = raw_comm_total - total_calc_comm
        
        st.divider()
        st.subheader("📊 結算數據統計")
        m1, m2, m3 = st.columns(3)
        m1.metric("📌 總所得稅金 (5% 進位)", f"${int(total_income_tax):,}")
        m2.metric("🏦 留凱基 (捨去後佣金總額)", f"${int(total_calc_comm):,}")
        m3.metric("🏦 匯國泰", f"${int(remit_cathay):,}")
        
        st.dataframe(res_df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            res_df.to_excel(writer, index=False, sheet_name='佣金明細')
            pd.DataFrame([{
                "總所得稅金(進位)": total_income_tax, 
                "留凱基合計(捨去)": total_calc_comm, 
                "匯國泰合計": remit_cathay
            }]).to_excel(writer, index=False, sheet_name='統計摘要')
        st.download_button(label="📥 下載結算報表", data=output.getvalue(), file_name="佣金結算單.xlsx")
    else:
        st.warning("⚠️ 查無符合條件的資料。")