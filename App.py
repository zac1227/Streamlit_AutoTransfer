import streamlit as st
import pandas as pd
from test import generate_codebook
import os
import io

st.title("📘 Codebook 產生器")

# 上傳資料表
data_file = st.file_uploader("請上傳主資料表（CSV）", type=["csv"])
code_file = st.file_uploader("請上傳變數類型定義檔（code.csv）", type=["csv"])

if data_file and code_file:
    # 載入成暫存檔（避免 winerror 32）
    data_bytes = data_file.read()
    code_bytes = code_file.read()

    with open("temp_data.csv", "wb") as f:
        f.write(data_bytes)
    with open("temp_code.csv", "wb") as f:
        f.write(code_bytes)

    if st.button("📊 產生 Codebook"):
        output_path = "codebook.docx"
        with st.spinner("正在產生報告..."):
            try:
                generate_codebook("temp_data.csv", "temp_code.csv", output_path)
                st.success("✅ 完成！以下為下載連結：")
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 下載 Word 報告",
                        data=f,
                        file_name="codebook.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                import traceback
                st.error("⚠️ 發生錯誤，以下是詳細訊息：")
                st.code(traceback.format_exc())
