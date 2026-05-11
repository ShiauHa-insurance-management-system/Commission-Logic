import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- 1. 系統設定與手機版優化 ---
st.set_page_config(page_title="產險佣金計算系統", layout="wide")

st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; height: 3.5em; margin-bottom: 10px; }
    .stDownloadButton>button { width: 100%; border-radius: 5px; height: 3.5em; background-color: #4CAF50; color: white; border: none; }
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

# --- 3. 佣金利率定義 ---
RATES = {
    "車險任意險": 0.07,
    "機車強制險": 0.04,
    "汽車強制險": 60, # 固定金額
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

# --- 5. 檔案上傳與智慧欄位清洗 ---
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
    
    # 清洗欄位名稱：去掉換行與空格
    df_prog.columns = [str(c).replace('\n', '').replace(' ', '').strip() for c in df_prog.columns]
    df_comm.columns = [str(c).replace('\n', '').replace(' ', '').strip() for c in df_comm.columns]

    # 智慧偵測身分證欄位 (解決 份/分 不一的問題)
    id_col_comm = next((c for c in df_comm.columns if "身分證" in c or "身份證" in c), "被保險人身分證字號/統一編號")
    id_col_prog = next((c for c in df_prog.columns if "身分證" in c or "身份證" in c), "被保險人身份證字號/統一編號")

    # 檢查佣金表關鍵欄位
    if "應領佣金" not in df_comm.columns:
        st.error(f"❌ 佣金表找不到『應領佣金』欄位。目前的欄位有：{list(df_comm.columns)}")
        st.stop()

    # --- 6. 核心計算邏輯 ---
    results = []
    # 總所得稅金 = 佣金表總「應領佣金」 * 5%
    raw_comm_total = pd.to_numeric(df_comm["應領佣金"], errors='coerce').sum()
    total_income_tax = raw_comm_total * 0.05

    for i, c_row in df_comm.iterrows():
        # 抓取佣金表資訊並去空格
        name = str(c_row.get("被保險人姓名", "")).strip()
        uid = str(c_row.get(id_col_comm, "")).strip()
        p_no = str(c_row.get("保單號碼", "")).strip()
        
        # 比對進度表 (強制轉字串比對，避免數字格式導致報錯)
        match = df_prog[
            (df_prog["被保險人姓名"].astype(str).str.strip() == name) &
            (df_prog[id_col_prog].astype(str).str.strip() == uid) &
            (df_prog["新年度保單號碼"].astype(str).str.strip() == p_no)
        ]
        
        if not match.empty:
            p_row = match.iloc[0]
            servicer = str(p_row.get("實際服務人員", "")).strip()
            
            # 指令：沒有服務人員則忽略
            if servicer:
                plate = p_row.get("牌照號碼", "")
                premium = pd.to_numeric(c_row.get("實收保費", 0), errors='coerce')
                
                # 判定險種關鍵字 (佣金表險種欄位或進度表種類)
                ins_type = str(c_row.get("險種", "")) + str(c_row.get("保件種類", "")) + str(p_row.get("保險種類", ""))
                
                calc_comm = 0
                if "機車" in ins_type and "強制" in ins_type:
                    calc_comm = premium * RATES["機車強制險"]
                elif "汽車" in ins_type and "強制" in ins_type:
                    calc_comm = RATES["汽車強制險"] # 固定 60 元
                elif "車" in ins_type:
                    calc_comm = premium * RATES["車險任意險"]
                elif "火" in ins_type:
                    calc_comm = premium * RATES["住火險"]
                elif "旅" in ins_type:
                    calc_comm = premium * RATES["旅平險"]
                elif "責任" in ins_type:
                    calc_comm = premium * RATES["產品責任險"]
                else:
                    calc_comm = 0

                results.append({
                    "製表日期": datetime.now().strftime("%Y-%m-%d"),
                    "被保險人姓名": name,
                    "保單號碼": p_no,
                    "車牌號碼": plate,
                    "實收保費": premium,
                    "應付佣金": round(calc_comm, 0),
                    "實際服務人員": servicer
                })

    # --- 7. 顯示結果與下載 ---
    if results:
        res_df = pd.DataFrame(results)
        total_calculated_comm = res_df["應付佣金"].sum() # 留凱基
        remit_cathay = raw_comm_total - total_calculated_comm # 匯國泰
        
        st.divider()
        st.subheader("📊 結算數據統計")
        m1, m2, m3 = st.columns(3)
        m1.metric("📌 總所得稅金 (5%)", f"${total_income_tax:,.0f}")
        m2.metric("🏦 留凱基 (應付佣金)", f"${total_calculated_comm:,.0f}")
        m3.metric("🏦 匯國泰", f"${remit_cathay:,.0f}")
        
        st.dataframe(res_df, use_container_width=True)

        # 產出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            res_df.to_excel(writer, index=False, sheet_name='佣金明細')
            pd.DataFrame([{
                "總所得稅金": total_income_tax,
                "留凱基": total_calculated_comm,
                "匯國泰": remit_cathay,
                "製表時間": datetime.now().strftime("%Y-%m-%d %H:%M")
            }]).to_excel(writer, index=False, sheet_name='匯總摘要')
        
        st.download_button(
            label="📥 下載佣金結算 Excel 報表",
            data=output.getvalue(),
            file_name=f"佣金結算單_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("⚠️ 比對完成，但沒有符合條件的資料。請確認兩份表格的「姓名」、「身分證/統一編號」及「保單號碼」是否完全一致，且進度表需有「實際服務人員」。")