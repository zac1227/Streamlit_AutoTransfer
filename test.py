from docx import Document
from docx.shared import Inches
import tempfile
import os

def generate_codebook(df, column_types, output_path="codebook.docx"):
    doc = Document()
    doc.add_heading("Codebook 統計摘要報告", level=1)

    # 🟦 首頁摘要統計
    try:
        total_rows, total_cols = df.shape
        doc.add_paragraph(f"總筆數：{total_rows}")
        doc.add_paragraph(f"欄位數：{total_cols}")
        doc.add_paragraph("含缺失值欄位：")

        na_series = df.isnull().sum()
        na_table = doc.add_table(rows=1, cols=3)
        na_table.style = "Table Grid"
        na_table.cell(0, 0).text = "欄位名稱"
        na_table.cell(0, 1).text = "缺失數"
        na_table.cell(0, 2).text = "缺失比例"

        for col in df.columns:
            na_count = na_series[col]
            if na_count > 0:
                row = na_table.add_row().cells
                row[0].text = col
                row[1].text = str(na_count)
                row[2].text = f"{na_count / total_rows:.2%}"
    except Exception as e:
        doc.add_paragraph(f"⚠️ 摘要統計錯誤：{e}")

    # 🟦 每個變數摘要
    for col in df.columns:
        if col not in column_types:
            continue

        vartype = column_types[col]
        if vartype == "0" or vartype == "略過":
            continue

        doc.add_heading(f"變數：{col}", level=2)

        if vartype == "1" or vartype == "連續":
            clean_series = df[col].dropna()

            if not pd.api.types.is_numeric_dtype(clean_series):
                doc.add_paragraph("⚠️ 資料非數值型，略過")
                continue

            mean_val = clean_series.mean()
            std_val = clean_series.std()
            min_val = clean_series.min()
            max_val = clean_series.max()

            table = doc.add_table(rows=3, cols=4)
            table.style = "Table Grid"
            table.cell(0, 0).text = "平均數"
            
            #table.cell(0, 1).text = f"{mean_val:.3f}" if pd.notna(mean_val) else ""
            table.cell(0, 2).text = "標準差"
            table.cell(0, 3).text = f"{std_val:.3f}" if pd.notna(std_val) else ""
            table.cell(1, 0).text = "最大值"
            table.cell(1, 1).text = f"{max_val:.3f}" if pd.notna(max_val) else ""
            table.cell(1, 2).text = "最小值"
            table.cell(1, 3).text = f"{min_val:.3f}" if pd.notna(min_val) else ""

            # Histogram
            if len(clean_series) > 0:
                fig, ax = plt.subplots()
                clean_series.plot(kind="hist", bins=10, color="skyblue", edgecolor="black", ax=ax)
                ax.set_title(f"Histogram of {col}")
                tmp1 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.tight_layout()
                plt.savefig(tmp1.name)
                plt.close("all")
                doc.add_picture(tmp1.name, width=Inches(4.5))
                os.unlink(tmp1.name)

        elif vartype == "2" or vartype == "類別":
            value_counts = df[col].value_counts(dropna=False)
            total = len(df)

            table = doc.add_table(rows=1 + len(value_counts), cols=3)
            table.style = "Table Grid"
            table.cell(0, 0).text = "類別"
            table.cell(0, 1).text = "數量"
            table.cell(0, 2).text = "比例"

            for i, (cat, count) in enumerate(value_counts.items(), start=1):
                table.cell(i, 0).text = str(cat)
                table.cell(i, 1).text = str(count)
                table.cell(i, 2).text = f"{count / total:.2%}"

            if len(value_counts) > 0:
                fig, ax = plt.subplots()
                value_counts.plot(kind="bar", color="cornflowerblue", ax=ax)
                ax.set_title(f"Bar Chart of {col}")
                tmp2 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.tight_layout()
                plt.savefig(tmp2.name)
                plt.close("all")
