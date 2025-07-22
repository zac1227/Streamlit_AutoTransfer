# app.py

import streamlit as st
import pandas as pd
from test import generate_codebook
import base64
import os

st.set_page_config(page_title="Codebook 產生器", layout="wide")
st.title("📄 自動化 Codebook 產生工具")

uploaded_file = st.file_uploader("請上傳 CSV 檔案", type=["csv", "xlsx"])

if uploaded_file is not None:
    # 讀取資料
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.success(f"成功讀取檔案，共 {df.shape[0]} 筆資料，{df.shape[1]} 欄位。")

    with st.expander("🔍 預覽資料"):
        st.dataframe(df.head())

    st.subheader("🧠 自動判斷欄位型別（可修改）")
    column_types = {}
    variable_names = {}
    category_definitions = {}

    type_options = ["連續型", "類別型", "時間型", "略過"]

    for col in df.columns:
        with st.container():
            st.markdown(f"**欄位：{col}**")
            col1, col2 = st.columns(2)
            with col1:
                # 判斷邏輯加強：
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    guess = "時間型"
                elif pd.api.types.is_numeric_dtype(df[col]):
                    guess = "連續型"
                elif df[col].nunique() < 20:
                    guess = "類別型"
                else:
                    guess = "略過"

                column_types[col] = st.selectbox(
                    "變數型別", type_options, index=type_options.index(guess), key=f"type_{col}"
                )
            with col2:
                variable_names[col] = st.text_input("變數名稱（選填）", value=col, key=f"name_{col}")

            if column_types[col] == "類別型":
                if df[col].nunique() <= 20:
                    unique_vals = df[col].dropna().unique()
                    defs = {}
                    for val in unique_vals:
                        defs[val] = st.text_input(f"定義：{val}", key=f"def_{col}_{val}")
                    category_definitions[col] = defs
                else:
                    st.info("類別數過多，略過定義填寫。")

    st.markdown("---")
    st.subheader("📤 報告產出")

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

    st.markdown("---")
    st.caption("💡 註：若選擇『略過』，該欄位將不納入報告產出。")