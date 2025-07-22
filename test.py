from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt
import pandas as pd
import tempfile, os

def generate_codebook(df, column_types, variable_names, category_definitions, output_path="codebook.docx"):
    doc = Document()
    doc.add_heading("Codebook 統計摘要報告", level=1)

    # 摘要統計
    rows, cols = df.shape
    doc.add_paragraph(f"總筆數：{rows}")
    doc.add_paragraph(f"欄位數：{cols}")
    doc.add_paragraph("有缺失值的欄位：")
    na_table = doc.add_table(rows=1, cols=3)
    na_table.style = "Table Grid"
    na_table.cell(0, 0).text = "欄位"
    na_table.cell(0, 1).text = "缺失數"
    na_table.cell(0, 2).text = "缺失比例"
    for c in df.columns:
        n = df[c].isnull().sum()
        if n > 0:
            r = na_table.add_row().cells
            r[0].text, r[1].text, r[2].text = c, str(n), f"{n/rows:.2%}"

    # 各欄位摘要
    for col, tp in column_types.items():
        if col not in df.columns or tp == "略過": continue
        var = variable_names.get(col, col)
        doc.add_heading(f"{col}（{var}）", level=2)

        if tp == "類別型":
            vc = df[col].value_counts(dropna=False)
            t = len(df)
            doc.add_paragraph("類別分布：")
            for k, v in vc.items():
                doc.add_paragraph(f"{k}: {v} ({v/t:.2%})", style="List Bullet")
            if len(vc.dropna())>0:
                fig, ax = plt.subplots()
                vc.sort_index().plot(kind="bar", ax=ax, color="cornflowerblue")
                ax.set_title(f"{col} 分布")
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.tight_layout(); plt.savefig(tmp.name); plt.close()
                doc.add_picture(tmp.name, width=Inches(4))
                os.unlink(tmp.name)

        elif tp == "連續型":
            if not pd.api.types.is_numeric_dtype(df[col]): continue
            s = df[col]
            mean, std, mn, mx = s.mean(), s.std(), s.min(), s.max()
            doc.add_paragraph(f"平均數：{mean:.3f}")
            doc.add_paragraph(f"標準差：{std:.3f}")
            doc.add_paragraph(f"最小值：{mn:.3f}")
            doc.add_paragraph(f"最大值：{mx:.3f}")
            if s.dropna().shape[0]>0:
                fig, ax = plt.subplots()
                s.plot(kind="hist", bins=10, color="skyblue", ax=ax)
                ax.set_title(f"{col} 分布")
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.tight_layout(); plt.savefig(tmp.name); plt.close()
                doc.add_picture(tmp.name, width=Inches(4))
                os.unlink(tmp.name)
                fig2, ax2 = plt.subplots()
                df.boxplot(column=col, ax=ax2)
                ax2.set_title(f"{col} Boxplot")
                tmp2 = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
                plt.tight_layout(); plt.savefig(tmp2.name); plt.close()
                doc.add_picture(tmp2.name, width=Inches(4))
                os.unlink(tmp2.name)

    doc.save(output_path)
    return output_path
