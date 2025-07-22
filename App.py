import streamlit as st
import pandas as pd
from test import generate_codebook
import base64
import os

st.set_page_config(page_title="Codebook 產生器", layout="wide")
st.title("📄 自動化 Codebook 產生工具")

uploaded_file = st.file_uploader("請上傳資料檔案（CSV 或 Excel）", type=["csv", "xlsx"])
meta_file = st.file_uploader("請上傳欄位型別設定檔（code.csv）", type=["csv"])

if uploaded_file is not None and meta_file is not None:
    # 讀取資料
    df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith(".csv") else pd.read_excel(uploaded_file)
    meta_df = pd.read_csv(meta_file)

    # 解析型別設定
    type_mapping = {"0": "略過", "1": "連續型", "2": "類別型"}
    column_types = {row["colname"]: type_mapping.get(str(row["type"]), "略過") for _, row in meta_df.iterrows()}
    variable_names = {col: col for col in df.columns}  # 無 label，因此使用原欄位名稱
    category_definitions = {}  # 不提供定義

    st.success(f"成功讀取資料，共 {df.shape[0]} 筆資料、{df.shape[1]} 欄位。")

    # 產出報告
    if st.button("🚀 產出 Codebook"):
        with st.spinner("產生中..."):
            try:
                output_path = generate_codebook(df, column_types, variable_names, category_definitions)
                with open(output_path, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode()
                    href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{os.path.basename(output_path)}">📥 點我下載 Codebook 報告 (Word)</a>'
                    st.markdown(href, unsafe_allow_html=True)
                st.success("✅ 報告產出成功")
            except Exception as e:
                st.error(f"發生錯誤：{e}")
