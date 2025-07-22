import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Inches
import tempfile
import os

def generate_codebook(data_path, code_path, output_path="codebook.docx"):
    df = pd.read_csv(data_path)
    code_df = pd.read_csv(code_path)
    type_map = dict(zip(code_df["Column"], code_df["Type"]))

    doc = Document()
    doc.add_heading("Codebook 統計摘要報告", level=1)

    # ➤ 資料摘要
    doc.add_paragraph(f"總筆數：{len(df)}")
    doc.add_paragraph(f"欄位數：{df.shape[1]}")
    doc.add_paragraph(" ")

    # ➤ 缺失值統計表格
    doc.add_heading("缺失值統計", level=2)
    na_series = df.isnull().sum()
    na_percent = df.isnull().mean()

    na_table = doc.add_table(rows=1, cols=3)
    na_table.style = "Table Grid"
    na_table.cell(0, 0).text = "欄位名稱"
    na_table.cell(0, 1).text = "缺失數"
    na_table.cell(0, 2).text = "缺失比例"

    for col in df.columns:
        missing_count = na_series[col]
        missing_pct = na_percent[col]
        if missing_count > 0:
            row = na_table.add_row().cells
            row[0].text = col
            row[1].text = str(missing_count)
            row[2].text = f"{missing_pct:.2%}"

    doc.add_paragraph(" ")

    # ➤ 各變數報告（文字 + 圖表）
    for col, t in type_map.items():
        if t == 0 or col not in df.columns:
            continue

        doc.add_heading(f"變數：{col}", level=2)
        col_data = df[col].dropna()

        # 統計資訊 + 圖表放進同一個表格
        table = doc.add_table(rows=1, cols=2)
        table.style = "Table Grid"
        table.cell(0, 0).text = "統計資訊"
        table.cell(0, 1).text = "圖表"

        if t == 1:  # 🔹連續變數
            desc = col_data.describe()
            q1 = col_data.quantile(0.25)
            q3 = col_data.quantile(0.75)
            stats = (
                f"類型：連續變數（基於非缺失資料）\n"
                f"平均值：{desc['mean']:.2f}\n"
                f"標準差：{col_data.std():.2f}\n"
                f"最小值：{desc['min']}\n"
                f"Q1（25%）：{q1}\n"
                f"中位數（50%）：{desc['50%']}\n"
                f"Q3（75%）：{q3}\n"
                f"最大值：{desc['max']}"
            )
            table.cell(0, 0).text = stats

            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                fig, axs = plt.subplots(2, 1, figsize=(6, 6))
                col_data.hist(bins=30, ax=axs[0])
                axs[0].set_title(f"{col} - 分布圖")
                axs[1].boxplot(col_data, vert=False)
                axs[1].set_title(f"{col} - Boxplot")
                plt.tight_layout()
                plt.savefig(tmp.name)
                plt.close()
                run = table.cell(0, 1).paragraphs[0].add_run()
                run.add_picture(tmp.name, width=Inches(4.5))
                os.unlink(tmp.name)

        elif t == 2:  # 🔸類別變數
            counts = col_data.value_counts(dropna=False)
            total = counts.sum()
            stats_lines = [f"{k}: {v} 筆（{v/total:.2%}）" for k, v in counts.items()]
            stats = (
                f"類型：類別變數（含缺失值統計）\n" +
                "\n".join(stats_lines) +
                f"\n總樣本數：{total}"
            )
            table.cell(0, 0).text = stats

            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                plt.figure(figsize=(6, 4))
                counts.plot(kind='bar')
                plt.title(f"{col} - 長條圖")
                plt.tight_layout()
                plt.savefig(tmp.name)
                plt.close()
                run = table.cell(0, 1).paragraphs[0].add_run()
                run.add_picture(tmp.name, width=Inches(4.5))
                os.unlink(tmp.name)

        doc.add_paragraph(" ")

    doc.save(output_path)
