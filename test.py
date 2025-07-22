from docx import Document
import pandas as pd


def generate_codebook(df, column_types, variable_names, category_definitions, output_path="codebook.docx"):
    doc = Document()
    doc.add_heading("Codebook 統計摘要報告", level=1)

    # 首頁摘要
    total_rows, total_cols = df.shape
    doc.add_paragraph(f"總筆數（資料列數）：{total_rows}")
    doc.add_paragraph(f"欄位數（變數數量）：{total_cols}")

    # 各欄位統計
    for col in column_types:
        if col not in df.columns:
            continue
        if column_types.get(col) == "略過":
            continue

        var_name = variable_names.get(col, col)
        doc.add_heading(f"變數：{col}（{var_name}）", level=2)

        if column_types[col] == "類別型":
            value_counts = df[col].value_counts(dropna=False)
            for val, count in value_counts.items():
                percent = count / len(df)
                doc.add_paragraph(f"{val}: {count} ({percent:.2%})", style="List Bullet")

        elif column_types[col] == "連續型":
            if not pd.api.types.is_numeric_dtype(df[col]):
                continue
            mean_val = df[col].mean()
            std_val = df[col].std()
            min_val = df[col].min()
            max_val = df[col].max()

            doc.add_paragraph(f"平均數: {mean_val:.3f}")
            doc.add_paragraph(f"標準差: {std_val:.3f}")
            doc.add_paragraph(f"最大值: {max_val:.3f}")
            doc.add_paragraph(f"最小值: {min_val:.3f}")

    doc.save(output_path)
    return output_path
