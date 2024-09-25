# 数据读取
import pandas as pd

filepath = "data.xlsx"
df = pd.read_excel(filepath)
print(df.info)
# 准确率计算
# bertscore计算
# 数据分析