from pathlib import Path
import pandas as pd

# 文件的目录
this_dir = Path(__file__).resolve().parent

# 从sales_data的所有子文件夹中读取Excel文件
parts = []
for path in (this_dir / "sales_data").rglob("*.xls*"):
    print(f"Reading {path.name}")
    part = pd.read_excel(path, index_col="transaction_id")
    parts.append(part)

# 将从Excel文件生成的DataFrame结合成单个DataFrame，pandas会负责对列进行对齐
df = pd.concat(parts)

# 对每个营业厅进行数据透视，将同一天产生的交易全部加起来
pivot = pd.pivot_table(df, 
                       index="transaction_date", 
                       columns="store", 
                       values="amount", 
                       aggfunc="sum")

# 按月重采样，并赋予一个索引名称
summary = pivot.resample("M").sum()
summary.index.name = "Month"

# 将总结报表写入Excel文件
summary.to_excel(this_dir / "sales_report_pandas.xlsx")