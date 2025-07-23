import streamlit as st
import pandas as pd
import base64
import os
import io

from test import generate_codebook  # 確保 test.py 有放對位置並含有該函式
tab1, tab2 = st.tabs(["📄 Codebook 產生器","📊 進階分析工具", ])


with tab1:
    st.title("📄 自動化 Codebook 產生工具")
    # ---------- 📥 上傳區塊 ----------
    st.set_page_config(page_title="Codebook 產生器", layout="wide")
    st.title("📄 自動化 Codebook 產生工具")

    def read_uploaded_csv(uploaded_file):
        for enc in ["utf-8", "utf-8-sig", "cp950", "big5"]:
            try:
                return pd.read_csv(io.TextIOWrapper(uploaded_file, encoding=enc))
            except Exception:
                uploaded_file.seek(0)
                continue
        st.error("❌ 檔案無法讀取，請確認是否為有效的 CSV 並使用常見編碼（UTF-8、BIG5、CP950）")
        return None

    uploaded_maindata = st.file_uploader("📂 請上傳主資料檔（CSV）", type=["csv"])
    uploaded_codebook = st.file_uploader("📋 請上傳變數類型檔 code.csv（需包含 Column 與 Type 欄位）", type=["csv"])

    df = None
    code_df = None

    if uploaded_maindata:
        df = read_uploaded_csv(uploaded_maindata)
        if df is not None:
            st.success(f"✅ 成功讀取主資料，共 {df.shape[0]} 筆，{df.shape[1]} 欄位。")
            with st.expander("🔍 預覽主資料"):
                st.dataframe(df.head())

    if uploaded_codebook:
        code_df = read_uploaded_csv(uploaded_codebook)
        if code_df is not None:
            st.success(f"✅ 成功讀取 Codebook，共定義 {len(code_df)} 欄位。")

    # ---------- 🧠 資料處理 ----------
    st.markdown("---")
    st.subheader("📤 報告產出區")

    if df is not None and code_df is not None:
        # 建立欄位型別設定
        column_types = dict(zip(code_df["Column"], code_df["Type"]))
        variable_names = {col: col for col in df.columns}
        category_definitions = {}  # 若你之後想加對應定義，可放這裡

        if st.button("🚀 產出 Codebook 報告"):
            with st.spinner("📄 產出中，請稍候..."):
                try:
                    output_path = "codebook.docx"  # 明確指定 Word 檔案名稱
                    output_path = generate_codebook(
                        df, column_types, variable_names, category_definitions, output_path=output_path
                    )
                    with open(output_path, "rb") as f:
                        file_data = f.read()
                        b64 = base64.b64encode(file_data).decode()
                        href = f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{os.path.basename(output_path)}">📥 點我下載 Codebook 報告（Word）</a>'
                        st.markdown(href, unsafe_allow_html=True)
                    st.success("✅ 報告已產出，可直接下載。")
                except Exception as e:
                    st.error(f"❌ 產出失敗：{e}")

# ---------- Tab 1 ----------
with tab2:
    st.title("📄")

    # 上傳區塊、read_uploaded_csv、資料預覽等原始程式碼照貼這裡
    ...
    # 產出報告的邏輯也放這裡
    ...

# ---------- Tab 2 ----------