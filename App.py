import streamlit as st
import pandas as pd
<<<<<<< Updated upstream
from test import generate_codebook
import base64, os
=======
from codebook_generator import generate_codebook
import os
import io
>>>>>>> Stashed changes

st.title("📘 Codebook 產生器")

<<<<<<< Updated upstream
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
=======
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
            generate_codebook("temp_data.csv", "temp_code.csv", output_path)
        st.success("✅ 完成！以下為下載連結：")

        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 下載 Word 報告",
                data=f,
                file_name="codebook.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        # 清理暫存
        os.remove("temp_data.csv")
        os.remove("temp_code.csv")
        os.remove("codebook.docx")
>>>>>>> Stashed changes
