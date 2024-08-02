import pandas as pd

file_path = r'C:\Users\杰哥\PycharmProjects\pythonProject33\pandas处理学校数据\data\data\调研学生科目得分表.xlsx'
df = pd.read_excel(file_path)
numeric_columns = [
    '总分', '语文原始成绩', '数学原始成绩', '英语原始成绩', '历史原始成绩',
    '物理原始成绩', '政治赋分成绩', '政治原始成绩', '地理赋分成绩',
    '地理原始成绩', '化学赋分成绩', '化学原始成绩', '生物赋分成绩', '生物原始成绩'
]

# Convert specified columns to numeric, coercing errors
df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='coerce')

# Display the first few rows to verify the conversion
print(df.info())
mean_scores = df.groupby(['区县', '学校', '班级']).mean().reset_index()
def calculate_rank(df, group_cols, rank_col):
    df[f'{rank_col}_区排名'] = df.groupby(group_cols)[rank_col].rank(ascending=False, method='min')
    return df

subjects = ['总分', '语文', '数学', '英语', '历史', '物理', '政治赋分', '地理赋分', '化学赋分', '生物赋分']
for subject in subjects:
    mean_scores = calculate_rank(mean_scores, ['区县'], subject)
for subject in subjects:
    mean_scores[f'{subject}_市排名'] = mean_scores[subject].rank(ascending=False, method='min')
mean_scores.to_excel('output_file.xlsx', index=False)
#
#
# # # 按区县分组并统计总分大于 85 分的人数
result = df[df['总分'] > a ].groupby('区县').size()/df[df['总分']].groupby('区县').size()
#
# pd.to_numeric()
subjects = ["历史", "政治", "地理", "化学", "生物"]

# 创建一个函数来计算选课组合
def calculate_combinations(row):
    combination = []
    if not pd.isnull(row["政治赋分"]):
        combination.append("政治")
    if not pd.isnull(row["地理赋分"]):
        combination.append("地理")
    if not pd.isnull(row["化学赋分"]):
        combination.append("化学")
    if not pd.isnull(row["生物赋分"]):
        combination.append("生物")
    return combination

# 应用函数到每一行数据
df["选课组合"] = df.apply(calculate_combinations, axis=1)

# 打印每个学生的选课组合
for combination in df["选课组合"]:
    print("，".join(combination))


