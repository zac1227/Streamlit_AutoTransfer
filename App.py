# app.py

import streamlit as st
import pandas as pd
from test import generate_codebook
import base64
import os

st.set_page_config(page_title="Codebook 產生器", layout="wide")
st.title("📄 自動化 Codebook 產生工具")

uploaded_maindata = st.file_uploader("請上傳 CSV 檔案", type=["csv"])
uploaded_codebook = st.file_uploader("請上傳 Codebook 檔案（選填，需含 Column 與 Type 欄位）", type=["csv"])

df = pd.read_csv(uploaded_maindata) if uploaded_maindata else None
code_df = pd.read_csv(uploaded_codebook) if uploaded_codebook else None

if df is not None:
    st.success(f"成功讀取檔案，共 {df.shape[0]} 筆資料，{df.shape[1]} 欄位。")

    with st.expander("🔍 預覽資料"):
        st.dataframe(df.head())

    st.markdown("---")
    st.subheader("📤 報告產出")

    if code_df is not None:
        # 建立 column_types dict
        column_types = dict(zip(code_df["Column"], code_df["Type"]))

        # 可選：分類說明與欄位標籤
        variable_names = {col: col for col in df.columns}  # 可改為自動讀 label
        category_definitions = {}  # 若你有類別說明可以加入

        if st.button("🚀 產出 Codebook"):
            with st.spinner("產生中..."):
                try:
                    output_path = generate_codebook(df, column_types, variable_names, category_definitions)
                    with open(output_path, "rb") as f:
                        file_data = f.read()
                        b64 = base64.b64encode(file_data).decode()
                        href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{os.path.basename(output_path)}">📥 點我下載 Codebook 報告 (Word)</a>'
                        st.markdown(href, unsafe_allow_html=True)
                    st.success("✅ 報告產出成功，可直接下載。")
                except PermissionError as e:
                    st.error(f"⚠️ 檔案處理失敗：{e}")
    else:
        st.warning("⚠️ 請上傳 Codebook 檔案（需含 Column 與 Type 欄位）")

    st.markdown("---")
    st.caption("💡 註：若選擇『略過』，該欄位將不納入報告產出。")
