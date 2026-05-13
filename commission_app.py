import streamlit as st
import pandas as pd
import io
import math
from datetime import datetime

# --- 1. 系統設定 ---
st.set_page_config(page_title="產險佣金計算系統", layout="wide")

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

# --- 3. 險種參數管理 (核心新功能) ---
st.sidebar.title("⚙️ 參數設定中心")

# 預設參數內容
default_rates = {
    "險種名稱": ["車險任意險", "機車強制險", "汽車強制險", "住火險", "旅平險", "產品責任險"],
    "佣金趴數或金額": [0.07, 0.04, 60.0, 0.03, 0.10, 0.10],
    "計算類型": ["百分比", "百分比", "固定金額", "百分比", "百分比", "百分比"]
}

# 讓使用者可以上傳之前的設定檔
config_file = st.sidebar.file_uploader("📂 載入參數設定檔 (.xlsx)", type="xlsx")

if config_file:
    df_rates = pd.read_excel(config_file)
else:
    if 'df_rates' not in st.session_state:
        st.session_state.df_rates = pd.DataFrame(default_rates)
    df_rates = st.session_state.df_rates

# 顯示並編輯參數表
st.sidebar.subheader("編輯險種與佣金")
edited_df = st.sidebar.data_editor(df_rates, num_rows="dynamic", hide_index=True)
st.session_state.df_rates = edited_df

# 下載設定檔功能
output_config = io.BytesIO()
with pd.ExcelWriter(output_config, engine='xlsxwriter') as writer:
    edited_df.to_excel(writer, index=False)
st.sidebar.download_button(
    label="💾 下載當前參數設定",
    data=output_config.getvalue(),
    file_name=f"佣金參數設定_{datetime.now().strftime('%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# 將設定轉換成程式可讀的 dictionary
current_rates = {}
for _, row in edited_df.iterrows():
    current_rates[row["險種名稱"]] = {
        "val": row["佣金趴數或金額"],
        "type": row["計算類型"]
    }

# --- 4. 主介面：上傳資料 ---
st.title("💰 佣金對帳與計算中心")

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
        st.error("❌ 佣金表標題定位失敗。")
        st.stop()

    # --- 5. 核心運算邏輯 ---
    results = []
    raw_comm_total = pd.to_numeric(df_comm["應領佣金"], errors='coerce').sum()
    total_income_tax = math.ceil(raw_comm_total * 0.10) # 所得稅 10% 進位

    for i, c_row in df_comm.iterrows():
        c_name = str(c_row.get("被保險人", c_row.get("被保險人姓名", ""))).strip()
        c_p_no = str(c_row.get("保單號碼", c_row.get("新年度保單號碼", ""))).strip()
        
        if not c_name or c_name == "nan" or "備註" in c_name: continue

        match = df_prog[
            ((df_prog.get("被保險人姓名", pd.Series()).astype(str).str.strip() == c_name) |
             (df_prog.get("被保險人", pd.Series()).astype(str).str.strip() == c_name)) &
            (df_prog.get("新年度保單號碼", pd.Series()).astype(str).str.strip() == c_p_no)
        ]
        
        if not match.empty:
            p_row = match.iloc[0]
            servicer = str(p_row.get("實際服務人員", "")).strip()
            
            if servicer == "謝騏鴻": continue # 過濾謝騏鴻
            
            if servicer:
                premium = pd.to_numeric(c_row.get("實收保費", 0), errors='coerce')
                ins_desc = str(c_row.get("險種", "")) + str(p_row.get("保險種類", ""))
                
                calc_comm = 0
                # 根據側邊欄設定的「險種名稱」進行關鍵字比對
                for target_name, config in current_rates.items():
                    if target_name in ins_desc:
                        if config["type"] == "百分比":
                            calc_comm = premium * config["val"]
                        else: # 固定金額
                            calc_comm = config["val"]
                        break # 匹配到第一個就跳出

                final_comm = math.floor(calc_comm) # 應付佣金無條件捨去

                results.append({
                    "被保險人姓名": c_name,
                    "保單號碼": c_p_no,
                    "車牌號碼": p_row.get("牌照號碼", ""),
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
        m1.metric("📌 總所得稅金 (10% 進位)", f"${int(total_income_tax):,}")
        m2.metric("🏦 留凱基 (捨去後佣金總額)", f"${int(total_calc_comm):,}")
        m3.metric("🏦 匯國泰", f"${int(remit_cathay):,}")
        
        st.dataframe(res_df, use_container_width=True)
        
        output_res = io.BytesIO()
        with pd.ExcelWriter(output_res, engine='xlsxwriter') as writer:
            res_df.to_excel(writer, index=False, sheet_name='佣金明細')
        st.download_button(label="📥 下載結算報表", data=output_res.getvalue(), file_name="佣金結算單.xlsx")