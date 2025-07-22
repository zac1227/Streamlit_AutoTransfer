from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt
import pandas as pd
import tempfile
import os

def generate_codebook(df, column_types, variable_names, category_definitions, output_path="codebook.docx", preview_mode=False):
    doc = Document()
    doc.add_heading("Codebook 統計摘要報告", level=1)

    # ✅ 首頁摘要統計區塊
    try:
        total_rows, total_cols = df.shape
        doc.add_paragraph(f"總筆數 - 資料列數: {total_rows}")
        doc.add_paragraph(f"欄位數 - 變數數量: {total_cols}")
        doc.add_paragraph("欄位缺失值統計（僅顯示有缺失的欄位）")

        na_series = df.isnull().sum()
        na_table = doc.add_table(rows=1, cols=3)
        na_table.style = "Table Grid"
        na_table.cell(0, 0).text = "欄位名稱"
        na_table.cell(0, 1).text = "缺失數"
        na_table.cell(0, 2).text = "缺失比例"

        for col in df.columns:
            na_count = na_series[col]
            if na_count > 0:
                row_cells = na_table.add_row().cells
                row_cells[0].text = str(col)
                row_cells[1].text = str(na_count)
                row_cells[2].text = f"{na_count / total_rows:.2%}"
    except Exception as e:
        doc.add_paragraph(f"⚠️ 無法產生摘要統計資訊：{e}")

    # ✅ 處理每個欄位
    for col in column_types:
        if col not in df.columns or column_types[col] == "略過":
            continue

        var_name = variable_names.get(col, col)
        doc.add_heading(f"變數：{col}（{var_name}）", level=2)

        if column_types[col] == "類別型":
            value_counts = df[col].value_counts(dropna=False)
            total = len(df)
            defs = category_definitions.get(col, {})

            categories = list(value_counts.index)
            percents = [f"{value_counts[k]} ({value_counts[k]/total:.2%})" for k in categories]
            cat_labels = [str(k) for k in categories]
            def_labels = [defs.get(k, "") for k in categories]

            table = doc.add_table(rows=4, cols=4)
            table.style = "Table Grid"
            table.cell(0, 0).text = "變數編號"
            table.cell(0, 1).text = col
            table.cell(0, 2).text = "變數名稱"
            table.cell(0, 3).text = var_name
            table.cell(1, 0).text = "變數類別"
            table.cell(1, 1).text = cat_labels[0] if len(cat_labels) > 0 else ""
            table.cell(1, 2).text = cat_labels[1] if len(cat_labels) > 1 else ""
            table.cell(1, 3).text = "變數定義"
            table.cell(2, 0).text = ""
            table.cell(2, 1).text = def_labels[0] if len(def_labels) > 0 else ""
            table.cell(2, 2).text = def_labels[1] if len(def_labels) > 1 else ""
            table.cell(2, 3).text = ""
            table.cell(3, 0).text = "數量與比例"
            table.cell(3, 1).text = percents[0] if len(percents) > 0 else ""
            table.cell(3, 2).text = percents[1] if len(percents) > 1 else ""
            table.cell(3, 3).text = "圖片"

            if df[col].dropna().shape[0] > 0:
                fig, ax = plt.subplots()
                value_counts.sort_index().plot(kind="bar", color="cornflowerblue", ax=ax)
                ax.set_title(f"Count Plot of {col}")
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.tight_layout()
                plt.savefig(tmp.name)
                plt.close("all")
                doc.add_picture(tmp.name, width=Inches(4.5))
                try:
                    os.unlink(tmp.name)
                except PermissionError:
                    pass

        elif column_types[col] == "連續型":
            if not pd.api.types.is_numeric_dtype(df[col]):
                continue

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
            table.cell(1, 1).text = f"{mean_val:.3f}" if not pd.isna(mean_val) else ""
            table.cell(1, 2).text = "標準差"
            table.cell(1, 3).text = f"{std_val:.3f}" if not pd.isna(std_val) else ""
            table.cell(2, 0).text = "最大值"
            table.cell(2, 1).text = f"{max_val:.3f}" if not pd.isna(max_val) else ""
            table.cell(2, 2).text = "最小值"
            table.cell(2, 3).text = f"{min_val:.3f}" if not pd.isna(min_val) else ""

            if df[col].dropna().shape[0] > 0:
                fig, ax = plt.subplots()
                df[col].plot(kind="hist", bins=10, color="skyblue", edgecolor="black", ax=ax)
                ax.set_title(f"Histogram of {col}")
                tmp1 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.tight_layout()
                plt.savefig(tmp1.name)
                plt.close("all")
                doc.add_picture(tmp1.name, width=Inches(4.5))
                try:
                    os.unlink(tmp1.name)
                except PermissionError:
                    pass

                fig2, ax2 = plt.subplots()
                df.boxplot(column=col, ax=ax2)
                ax2.set_title(f"Boxplot of {col}")
                tmp2 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.tight_layout()
                plt.savefig(tmp2.name)
                plt.close("all")
                doc.add_picture(tmp2.name, width=Inches(4.5))
                try:
                    os.unlink(tmp2.name)
                except PermissionError:
                    pass

    doc.save(output_path)
    return output_path
