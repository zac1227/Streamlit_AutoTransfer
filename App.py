import streamlit as st
import pandas as pd
from test import generate_codebook  # 你自己的函數檔案
import os

st.title("📊 Codebook 統計摘要產生器")

# 上傳資料與欄位說明
uploaded_data = st.file_uploader("📎 請上傳主資料（data.csv）", type=["csv"])
uploaded_meta = st.file_uploader("📎 請上傳變數說明（code.csv）", type=["csv"])

if uploaded_data and uploaded_meta:
    try:
        # 讀取主資料與變數定義
        df = pd.read_csv(uploaded_data)
        meta_df = pd.read_csv(uploaded_meta)

        # 建立 type 對應表
        type_map = {0: "略過", 1: "連續型", 2: "類別型"}
        column_types = {row["colname"]: type_map.get(row["type"], "略過") for _, row in meta_df.iterrows()}
        variable_names = {row["colname"]: row["variable_name"] for _, row in meta_df.iterrows()}
        category_definitions = {}  # 若無額外定義可留空

        # 顯示簡易預覽
        st.success("✅ 檔案上傳成功！以下是欄位資訊預覽：")
        st.dataframe(meta_df)

        # 產生報告
        if st.button("📄 產生 Codebook"):
            output_path = generate_codebook(df, column_types, variable_names, category_definitions)
            st.success("📘 Codebook 產生成功！")

            # 提供下載
            with open(output_path, "rb") as f:
                st.download_button("⬇️ 下載報告", f, file_name="codebook.docx")

    except Exception as e:
        st.error(f"❌ 發生錯誤：{e}")
