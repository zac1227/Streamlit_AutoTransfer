import streamlit as st
import pandas as pd
import base64
import os
import io

from test import generate_codebook  # 確保 test.py 有放對位置並含有該函式
st.set_page_config(page_title="Codebook 產生器", layout="wide")
tab1, tab2 = st.tabs(["📄 Codebook 產生器","📊 進階分析工具", ])


with tab1:
    st.title("📄 自動化 Codebook 產生工具")
    # ---------- 📥 上傳區塊 ----------
    
    

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
    # 移除 Type 為 0 的欄位
        code_df = code_df[~code_df["Type"].astype(str).str.lower().eq("0")]

        # 建立欄位型別與角色（X 或 Y）
        column_types = {}
        variable_names = {}
        column_roles = {}
        x_counter = 1
        y_counter = 1

        for _, row in code_df.iterrows():
            col = row["Column"]
            t = str(row["Type"]).lower()

            if t == "y1":
                column_roles[col] = "Y"
                column_types[col] = 1
            elif t == "y2":
                column_roles[col] = "Y"
                column_types[col] = 2
            elif t in ["1", "2"]:
                column_roles[col] = f"X{x_counter}"
                column_types[col] = int(t)
                x_counter += 1
            else:
                st.warning(f"⚠️ Unknown Type '{t}' for column '{col}' — skipped.")
                continue

            variable_names[col] = column_roles.get(col, col)

                
        category_definitions = {}  # 若你之後想加對應定義，可放這裡
        if st.button("🚀 產出 Codebook 報告"):
            with st.spinner("📄 產出中，請稍候..."):
                try:
                    output_path = "codebook.docx"  # 明確指定 Word 檔案名稱
                    output_path = generate_codebook(
                        df, column_types, variable_names, category_definitions,code_df=code_df ,output_path=output_path
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
# ---------- Tab 2 ----------
with tab2:
    st.title("📊 進階分析工具")
    st.markdown("""
    ### 📘 功能說明
    本工具可根據 `code.csv` 中的 Transform 欄位，對主資料進行以下轉換：

    - 若 Transform 欄為空或 'none'，則不進行任何轉換。
    - 若為 `cut:[0,10,20,30]`，將以手動分箱方式進行區間分類（含邊界）。
    - 若為 `cut:quantile:4`，則會進行四等分的分位數切分。
    - 若為 `cut:uniform:3`，則會將資料等寬切為三段。
    - 若欄位為類別型態，或 Transform 欄為 `onehot`，則會進行 one-hot encoding。

    所有轉換後的欄位名稱將自動加上 `_binned` 或對應欄位前綴。
    """)
    def read_uploaded_csv(uploaded_file):
        for enc in ["utf-8", "utf-8-sig", "cp950", "big5"]:
            try:
                return pd.read_csv(io.TextIOWrapper(uploaded_file, encoding=enc))
            except Exception:
                uploaded_file.seek(0)
                continue
        st.error("❌ 檔案無法讀取，請確認是否為有效的 CSV 並使用常見編碼（UTF-8、BIG5、CP950）")
        return None

    uploaded_main = st.file_uploader("📂 請上傳主資料（CSV）", type=["csv"], key="main2")
    uploaded_code = st.file_uploader("📋 請上傳 code.csv（含 Transform 欄位）", type=["csv"], key="code2")

    df2, code2 = None, None

    if uploaded_main:
        df2 = read_uploaded_csv(uploaded_main)
    if uploaded_code:
        code2 = read_uploaded_csv(uploaded_code)

    if df2 is not None and code2 is not None:
        st.success("✅ 資料與 code.csv 載入成功")
        transform_col = []
        for _, row in code2.iterrows():
            col = row["Column"]
            transform = str(row.get("Transform", "")).strip()
            if transform.lower().startswith("cut:["):
                try:
                    bins = eval(transform[4:])
                    df2[col + "_binned"] = pd.cut(df2[col], bins=bins, include_lowest=True)
                except Exception as e:
                    st.warning(f"🔸 {col} 分箱失敗：{e}")
            elif transform.lower().startswith("cut:quantile:"):
                try:
                    q = int(transform.split(":")[-1])
                    df2[col + "_binned"] = pd.qcut(df2[col], q=q, duplicates='drop')
                except Exception as e:
                    st.warning(f"🔸 {col} 分位數切分失敗：{e}")
            elif transform.lower().startswith("cut:uniform:"):
                try:
                    k = int(transform.split(":")[-1])
                    df2[col + "_binned"] = pd.cut(df2[col], bins=k)
                except Exception as e:
                    st.warning(f"🔸 {col} 均分切分失敗：{e}")
            elif df2[col].dtype == 'object' or transform.lower() == 'onehot':
                try:
                    onehot = pd.get_dummies(df2[col], prefix=col)
                    df2 = pd.concat([df2, onehot], axis=1)
                except Exception as e:
                    st.warning(f"🔸 {col} one-hot 編碼失敗：{e}")
            elif transform == '' or transform.lower() == 'none':
                continue
            else:
                st.warning(f"🔸 未知 Transform 指令：{transform}（欄位 {col}）")

        st.markdown("---")
        st.subheader("🔍 預覽轉換後資料")
        st.dataframe(df2.head())

        csv = df2.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 下載轉換後的 CSV", data=csv, file_name="transformed_data.csv", mime="text/csv")
