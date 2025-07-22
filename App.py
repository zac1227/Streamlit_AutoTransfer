import streamlit as st
import pandas as pd
from test import generate_codebook
import base64, os

st.set_page_config(page_title="Codebook 產生器", layout="wide")
st.title("📄 自動化 Codebook 產生工具")

uploaded_file = st.file_uploader("請上傳資料檔（CSV 或 Excel）", type=["csv", "xlsx"])
meta_file = st.file_uploader("請上傳欄位型別設定（code.csv）", type=["csv"])

if uploaded_file and meta_file:
    try:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)
        meta_df = pd.read_csv(meta_file)

        if not {"Column", "Type"}.issubset(meta_df.columns):
            st.error("❌ code.csv 檔案需包含「Column」與「Type」欄位")
        else:
            mapping = {"0": "略過", "1": "連續型", "2": "類別型"}
            column_types = {r["Column"]: mapping.get(str(r["Type"]), "略過") for _, r in meta_df.iterrows()}
            variable_names = {col: col for col in df.columns}
            category_definitions = {}

            st.success(f"讀取完畢，資料共有 {df.shape[0]} 筆，{df.shape[1]} 欄位")
            st.dataframe(meta_df)

            if st.button("🚀 產出 Codebook"):
                output = generate_codebook(df, column_types, variable_names, category_definitions)
                with open(output, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{os.path.basename(output)}">📥 下載 Codebook</a>'
                st.markdown(href, unsafe_allow_html=True)
                st.success("✅ 報告產出成功")
    except Exception as e:
        st.error(f"串流執行時發生錯誤：{e}")
