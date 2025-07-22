from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt
import pandas as pd
import tempfile
import os

def generate_codebook(df, column_types, variable_names, category_definitions, output_path="codebook.docx"):
    doc = Document()
    doc.add_heading("Codebook 統計摘要報告", level=1)

    # 摘要統計
    total_rows, total_cols = df.shape
    doc.add_paragraph(f"總筆數：{total_rows}")
    doc.add_paragraph(f"欄位數：{total_cols}")
    doc.add_paragraph("欄位缺失值統計（僅顯示有缺失的欄位）：")

    na_table = doc.add_table(rows=1, cols=3)
    na_table.style = "Table Grid"
    na_table.cell(0, 0).text = "欄位名稱"
    na_table.cell(0, 1).text = "缺失數"
    na_table.cell(0, 2).text = "缺失比例"

    for col in df.columns:
        na_count = df[col].isnull().sum()
        if na_count > 0:
            row = na_table.add_row().cells
            row[0].text = col
            row[1].text = str(na_count)
            row[2].text = f"{na_count / total_rows:.2%}"

    for col in column_types:
        if col not in df.columns or column_types[col] == "略過":
            continue

        var_name = variable_names.get(col, col)
        doc.add_heading(f"變數：{col}（{var_name}）", level=2)

        if column_types[col] == "類別型":
            vc = df[col].value_counts(dropna=False)
            table = doc.add_table(rows=4, cols=4)
            table.style = "Table Grid"

            table.cell(0, 0).text = "變數編號"
            table.cell(0, 1).text = col
            table.cell(0, 2).text = "變數名稱"
            table.cell(0, 3).text = var_name

            keys = list(vc.index)
            defs = category_definitions.get(col, {})

            table.cell(1, 0).text = "變數類別"
            table.cell(1, 1).text = str(keys[0]) if len(keys) > 0 else ""
            table.cell(1, 2).text = str(keys[1]) if len(keys) > 1 else ""
            table.cell(1, 3).text = "變數定義"

            table.cell(2, 1).text = defs.get(keys[0], "") if len(keys) > 0 else ""
            table.cell(2, 2).text = defs.get(keys[1], "") if len(keys) > 1 else ""

            table.cell(3, 0).text = "數量與比例"
            total = len(df)
            table.cell(3, 1).text = f"{vc[keys[0]]} ({vc[keys[0]]/total:.2%})" if len(keys) > 0 else ""
            table.cell(3, 2).text = f"{vc[keys[1]]} ({vc[keys[1]]/total:.2%})" if len(keys) > 1 else ""
            table.cell(3, 3).text = "圖片"

            if df[col].dropna().shape[0] > 0:
                fig, ax = plt.subplots()
                vc.sort_index().plot(kind="bar", color="cornflowerblue", ax=ax)
                ax.set_title(f"Count Plot of {col}")
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.tight_layout()
                plt.savefig(tmp.name)
                plt.close("all")
                doc.add_picture(tmp.name, width=Inches(4.5))
                os.unlink(tmp.name)

        elif column_types[col] == "連續型" and pd.api.types.is_numeric_dtype(df[col]):
            mean_val = df[col].mean()
            std_val = df[col].std()
            min_val = df[col].min()
            max_val = df[col].max()

            table = doc.add_table(rows=3, cols=4)
            table.style = "Table Grid"
            table.cell(0, 0).text = "變數編號"
            table.cell(0, 1).text = col
            table.cell(0, 2).text = "變數名稱"
            table.cell(0, 3).text = var_name
            table.cell(1, 0).text = "平均數"
            table.cell(1, 1).text = f"{mean_val:.3f}"
            table.cell(1, 2).text = "標準差"
            table.cell(1, 3).text = f"{std_val:.3f}"
            table.cell(2, 0).text = "最大值"
            table.cell(2, 1).text = f"{max_val:.3f}"
            table.cell(2, 2).text = "最小值"
            table.cell(2, 3).text
