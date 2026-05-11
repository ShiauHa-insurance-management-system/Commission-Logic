import streamlit as st
import pandas as pd
import io
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
    # 進度表：根據截圖，標題通常在第1或2行。我們用 header=0 讀取後清洗。
    # 如果進度表第一行是空白，pandas 會自動處理，這裡我們確保讀到標題。
    df_prog = pd.read_excel(prog_file).fillna("")
    df_prog = clean_cols(df_prog)
    
    # 佣金表：標題在第 3 行 (header=2)
    df_comm = pd.read_excel(comm_file, header=2).fillna("")
    df_comm = clean_cols(df_comm)

    # 檢查佣金表標頭
    if "應領佣金" not in df_comm.columns:
        st.error(f"❌ 佣金表標題定位失敗。看到的標題有：{list(df_comm.columns)}")
        st.stop()

    # --- 5. 核心運算邏輯 ---
    results = []
    # 總所得稅金 = 佣金表「應領佣金」總額 * 5%
    raw_comm_total = pd.to_numeric(df_comm["應領佣金"], errors='coerce').sum()
    total_income_tax = raw_comm_total * 0.05

    for i, c_row in df_comm.iterrows():
        # 佣金表欄位名稱
        c_name = str(c_row.get("被保險人姓名", c_row.get("被保險人", ""))).strip()
        c_p_no = str(c_row.get("保單號碼", c_row.get("新年度保單號碼", ""))).strip()
        
        if not c_name or c_name == "nan" or "備註" in c_name: continue

        # 比對進度表
        # 我們同時支援「被保險人姓名」或「被保險人」
        match = df_prog[
            (
                (df_prog.get("被保險人姓名", pd.Series()).astype(str).str.strip() == c_name) |
                (df_prog.get("被保險人", pd.Series()).astype(str).str.strip() == c_name)
            ) &
            (df_prog.get("新年度保單號碼", pd.Series()).astype(str).str.strip() == c_p_no)
        ]
        
        if not match.empty:
            p_row = match.iloc[0]
            servicer = str(p_row.get("實際服務人員", "")).strip()
            
            if servicer:
                plate = p_row.get("牌照號碼", "")
                premium = pd.to_numeric(c_row.get("實收保費", 0), errors='coerce')
                
                # 險種判定邏輯
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
                else:
                    calc_comm = 0

                results.append({
                    "製表日期": datetime.now().strftime("%Y-%m-%d"),
                    "被保險人姓名": c_name,
                    "保單號碼": c_p_no,
                    "車牌號碼": plate,
                    "實收保費": premium,
                    "應付佣金": round(calc_comm, 0),
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
        m1.metric("📌 總所得稅金 (5%)", f"${total_income_tax:,.0f}")
        m2.metric("🏦 留凱基 (應付佣金)", f"${total_calc_comm:,.0f}")
        m3.metric("🏦 匯國泰", f"${remit_cathay:,.0f}")
        
        st.dataframe(res_df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            res_df.to_excel(writer, index=False, sheet_name='佣金明細')
            pd.DataFrame([{"總所得稅金":total_income_tax, "留凱基":total_calc_comm, "匯國泰":remit_cathay}]).to_excel(writer, index=False, sheet_name='統計摘要')
        st.download_button(label="📥 下載結算報表", data=output.getvalue(), file_name="佣金結算單.xlsx")
    else:
        st.warning("⚠️ 檔案讀取成功，但找不到完全匹配的資料。")
        st.info("💡 請確認：進度表中的「新年度保單號碼」是否與佣金表的「保單號碼」完全一致。")