from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt
import pandas as pd
import tempfile
import os

def generate_codebook(df, column_types, variable_names, category_definitions, output_path="codebook.docx", preview_mode=False):
    if output_path is None:
        output_path = "codebook.docx"

    doc = Document()
    doc.add_heading("Codebook 統計摘要報告", level=1)
    df = df.dropna(how='all')  # 移除全為 NaN 的行

    for col in df.columns:
        if col not in column_types:
            continue
        type_code = column_types[col]
        if type_code == 0:
            continue

        var_name = variable_names.get(col, col)
        doc.add_heading(f"變數：{col}（{var_name}）", level=2)

        # 🟦 類別型欄位
        if type_code == 2:
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

            # 圖表
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

        # 🟩 連續型欄位
        elif type_code == 1:
            try:
                df[col] = pd.to_numeric(df[col], errors="coerce")
            except Exception as e:
                print(f"❌ 欄位 {col} 強制轉數值失敗：{e}")
                continue

            if df[col].dropna().empty:
                print(f"⚠️ 欄位 {col} 無有效數值（全為 NaN），略過")
                continue

            desc = df[col].describe()

            table = doc.add_table(rows=3, cols=4)
            table.style = "Table Grid"
            table.cell(0, 0).text = "變數編號"
            table.cell(0, 1).text = col
            table.cell(0, 2).text = "變數名稱"
            table.cell(0, 3).text = var_name

            table.cell(1, 0).text = "平均數"
            table.cell(1, 1).text = f"{desc['mean']:.3f}"
            table.cell(1, 2).text = "標準差"
            table.cell(1, 3).text = f"{desc['std']:.3f}"

            table.cell(2, 0).text = "最大值"
            table.cell(2, 1).text = f"{desc['max']:.3f}"
            table.cell(2, 2).text = "最小值"
            table.cell(2, 3).text = f"{desc['min']:.3f}"

            # Histogram
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

            # Boxplot
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
