import os
import pandas as pd
from bert_score import score

# 指定文件路径和结果保存路径
input_file_path = 'C:/Users/panxiaohang/llm/bertscore/标准版表格.xlsx'
output_file_path = 'C:/Users/panxiaohang/llm/bertscore/标准版表格bert_score结果.xlsx'

# 读取 Excel 文件
print("正在读取 Excel 文件...")
data = pd.read_excel(input_file_path, header=0)
print(f"文件读取成功，共有 {len(data)} 行数据。")

# 使用本地模型路径
model_name_or_path = 'bert-base-chinese'

# 检查模型路径是否存在
if not os.path.exists(model_name_or_path):
    print(f"错误：模型路径 {model_name_or_path} 不存在，请确保模型文件夹放在代码运行目录下。")
    exit()

print("模型和分词器加载成功！")

# 初始化结果存储，同时初始化一个列表来记录哪些行需要更新
results = []
update_rows = []

# 遍历每一行数据，检查条件并计算 BERTScore
for index, row in data.iterrows():
    if row[5] == 1 and row[7] == 1:  # 第6列和第8列都为1
        col7_text = str(row[6])  # 获取第7列
        col9_text = str(row[8])  # 获取第9列

        print(f"正在处理第 {index + 1} 行数据...")

        try:
            _, similarity_score, _ = score([col7_text], [col9_text], model_type=model_name_or_path, lang="zh")

            # 存储需要更新的行索引和相似度得分
            update_rows.append((index, similarity_score.item()))
            print(f"第 {index + 1} 行数据处理完成。")
        except Exception as e:
            print(f"第 {index + 1} 行数据处理时出错: {e}")

# 更新符合条件的行的第10列（索引9）为BERTScore相似度，其余行设置为NULL
for idx, sim_score in update_rows:
    data.at[idx, 9] = sim_score
data.fillna({'Column10': None}, inplace=True)  # 假设第10列名为'Column10'，如果没有则需调整

# 保存结果到指定目录的 Excel 文件
data.to_excel(output_file_path, index=False)
print(f"结果已保存到 {output_file_path}")

# 计算BERTScore的中位数、平均数以及各阈值的百分比
similarity_scores = [score for _, score in update_rows]
similarity_scores.sort()

n_scores = len(similarity_scores)
median_score = similarity_scores[n_scores // 2] if n_scores % 2 != 0 else (similarity_scores[n_scores // 2 - 1] + similarity_scores[n_scores // 2]) / 2
average_score = sum(similarity_scores) / n_scores

percent_ge_06 = sum(score >= 0.6 for score in similarity_scores) / n_scores * 100
percent_ge_065 = sum(score >= 0.65 for score in similarity_scores) / n_scores * 100
percent_ge_07 = sum(score >= 0.7 for score in similarity_scores) / n_scores * 100
percent_ge_08 = sum(score >= 0.8 for score in similarity_scores) / n_scores * 100

# 输出统计结果
print(f"\nBERTScore 统计结果:")
print(f"中位数: {median_score:.4f}")
print(f"平均数: {average_score:.4f}")
print(f">= 0.6 的比例: {percent_ge_06:.2f}%")
print(f">= 0.65 的比例: {percent_ge_065:.2f}%")
print(f">= 0.7 的比例: {percent_ge_07:.2f}%")
print(f">= 0.8 的比例: {percent_ge_08:.2f}%")

# 结束提示
print("\n所有统计分析已完成并已输出。")