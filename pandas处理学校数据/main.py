# import pandas as pd
#
# # 读取数据
# df = pd.read_excel(r'C:\Users\杰哥\PycharmProjects\pythonProject33\pandas处理学校数据\data\data\调研学生科目得分表.xlsx', dtype={'考号': str, '学籍号': str})
# # 将区县和学校列转换为 category 类型以减少内存占用
# df['区县'] = df['区县'].astype('category')
# df['学校'] = df['学校'].astype('category')
# subjects = [
#     '总分', '语文原始成绩', '数学原始成绩', '英语原始成绩', '历史原始成绩',
#     '物理原始成绩', '政治赋分成绩', '政治原始成绩', '地理赋分成绩',
#     '地理原始成绩', '化学赋分成绩', '化学原始成绩', '生物赋分成绩', '生物原始成绩'
# ]
#
#
# def determine_selection(row):
#     subjects = ['历史原始成绩', '物理原始成绩', '政治原始成绩', '地理原始成绩', '化学原始成绩', '生物原始成绩']
#     selection = []
#     for subject in subjects:
#         if not pd.isna(row[subject]):
#             selection.append(subject[:2])
#     return ','.join(selection)
#
#
# # 获取学生选的科目
# df['选择组合'] = df.apply(determine_selection, axis=1)
#
# # 统计函数
# def calculate_statistics(df, group_by_columns):
#     if group_by_columns:
#         # 统计每个组合的人数和总人数
#         count_df = df.groupby(group_by_columns + ['选择组合'], observed=True).size().reset_index(name='人数')
#         total_counts = df.groupby(group_by_columns, observed=True).size().reset_index(name='总人数')
#
#         # 合并数据以计算占比
#         stats_df = pd.merge(count_df, total_counts, on=group_by_columns)
#         stats_df['占比'] = stats_df['人数'] / stats_df['总人数']
#     else:
#         # 全市统计
#         count_df = df['选择组合'].value_counts().reset_index(name='人数')
#         count_df.columns = ['选择组合', '人数']
#         total_count = count_df['人数'].sum()
#         count_df['占比'] = count_df['人数'] / total_count
#         stats_df = count_df
#
#     return stats_df
#
#
# # 全市统计
# city_stats = calculate_statistics(df, [])
# # 分区县统计
# district_stats = calculate_statistics(df, ['区县'])
# # 分学校统计
# school_stats = calculate_statistics(df, ['区县', '学校'])
#
#
# # 清理无效数据
# df[subjects] = df[subjects].apply(pd.to_numeric, errors='coerce')
#
# # 定义上线标准
# standards = {
#     '物理类': {
#         '总分': {'high': 490, 'medium': 420},
#         '语文原始成绩': {'high': 96, 'medium': 85},
#         '数学原始成绩': {'high': 100, 'medium': 90},
#         '英语原始成绩': {'high': 95, 'medium': 85},
#         '历史原始成绩': {'high': 90, 'medium': 80},
#         '物理原始成绩': {'high': 92, 'medium': 82},
#         '政治赋分成绩': {'high': 88, 'medium': 78},
#         '地理赋分成绩': {'high': 87, 'medium': 77},
#         '化学赋分成绩': {'high': 89, 'medium': 79},
#         '生物赋分成绩': {'high': 83, 'medium': 73},
#     },
#     '历史类': {
#         '总分': {'high': 490, 'medium': 420},
#         '语文原始成绩': {'high': 100, 'medium': 87},
#         '数学原始成绩': {'high': 98, 'medium': 88},
#         '英语原始成绩': {'high': 97, 'medium': 86},
#         '历史原始成绩': {'high': 96, 'medium': 85},
#         '物理原始成绩': {'high': 95, 'medium': 84},
#         '政治赋分成绩': {'high': 94, 'medium': 83},
#         '地理赋分成绩': {'high': 92, 'medium': 81},
#         '化学赋分成绩': {'high': 90, 'medium': 79},
#         '生物赋分成绩': {'high': 88, 'medium': 77},
#     }
# }
#
#
# # 计算单科参考人数、平均分、上线人数和上线率的函数
# def calculate_pass_rate(df, category, subject, standards):
#     thresholds = standards[category][subject]
#     high_threshold = thresholds['high']
#     medium_threshold = thresholds['medium']
#
#     valid_students = df[df[subject] > 0]
#     total_students = len(valid_students)
#
#     avg = valid_students[subject].mean()
#     high_pass = valid_students[subject].ge(high_threshold).sum()
#     medium_pass = valid_students[subject].ge(medium_threshold).sum()
#
#     high_hit_pass = valid_students[
#         valid_students[subject].ge(high_threshold) &
#         valid_students['总分'].ge(standards[category]['总分']['high'])
#     ].shape[0]
#     medium_hit_pass = valid_students[
#         valid_students[subject].ge(medium_threshold) &
#         valid_students['总分'].ge(standards[category]['总分']['medium'])
#     ].shape[0]
#
#     high_pass_rate = high_pass / total_students if total_students > 0 else 0
#     medium_pass_rate = medium_pass / total_students if total_students > 0 else 0
#     high_hit_pass_rate = high_hit_pass / high_pass if high_pass > 0 else 0
#     medium_hit_pass_rate = medium_hit_pass / medium_pass if medium_pass > 0 else 0
#
#     return {
#         '参考人数': total_students,
#         '科目均分': avg,
#         '高线人数': high_pass,
#         '高线率': high_pass_rate,
#         '高线命中人数': high_hit_pass,
#         '高线命中率': high_hit_pass_rate,
#         '中线人数': medium_pass,
#         '中线率': medium_pass_rate,
#         '中线命中人数': medium_hit_pass,
#         '中线命中率': medium_hit_pass_rate
#     }
#
# def aggregate_results(df, group_by_cols, standards):
#     results = []
#
#     for group_values, group in df.groupby(group_by_cols):
#         if len(group_by_cols) == 3:
#             region, school, category = group_values
#         elif len(group_by_cols) == 2:
#             region, category = group_values
#             school = None
#         elif len(group_by_cols) == 1:
#             category = group_values
#             region = school = None
#
#         if category not in standards:
#             print(f"Category '{category}' not found in standards.")
#             continue
#
#         for subject in standards[category]:
#             pass_rates = calculate_pass_rate(group, category, subject, standards)
#             result = {
#                 **dict(zip(group_by_cols, group_values)),
#                 '类别': category,
#                 '科目': subject,
#                 **pass_rates
#             }
#             results.append(result)
#
#     return pd.DataFrame(results)
#
# def add_ranking(df, group_by_cols, rank_cols, prefix):
#     for col in rank_cols:
#         rank_col_name = f'{prefix}{col}排名'
#         df[rank_col_name] = df.groupby(group_by_cols)[col].rank(ascending=False, method='min')
#
# def calculate_combination_statistics(df, standards):
#     # 统计每个选择组合的人数和总分上线情况
#     combination_stats = []
#
#     for category, group in df.groupby('类别'):
#         high_threshold = standards[category]['总分']['high']
#         medium_threshold = standards[category]['总分']['medium']
#
#         comb_stats = group.groupby(['选择组合']).agg(
#             选择组合人数=('总分', 'size'),
#             高线总分人数=('总分', lambda x: (x >= high_threshold).sum()),
#             中线总分人数=('总分', lambda x: (x >= medium_threshold).sum())
#         ).reset_index()
#
#         comb_stats['高线总分率'] = comb_stats['高线总分人数'] / comb_stats['选择组合人数']
#         comb_stats['中线总分率'] = comb_stats['中线总分人数'] / comb_stats['选择组合人数']
#         comb_stats['类别'] = category
#
#         combination_stats.append(comb_stats)
#
#     return pd.concat(combination_stats, ignore_index=True)
#
# # 计算每个区域、每个学校、每个类别的每个科目的上线人数和上线率
# school_results = aggregate_results(df, ['区县', '学校', '类别'], standards)
# county_results = aggregate_results(df, ['区县', '类别'], standards)
# city_results = aggregate_results(df, ['类别'], standards)
#
# # 为学校结果添加排名
# add_ranking(school_results, ['类别', '科目'], ['科目均分', '高线率', '中线率'], '市')
# add_ranking(school_results, ['区县', '类别', '科目'], ['科目均分', '高线率', '中线率', '高线命中率', '中线命中率'], '区')
#
# # 为区县结果添加排名
# add_ranking(county_results, ['类别', '科目'], ['科目均分', '高线率', '中线率'], '市')
# add_ranking(county_results, ['区县', '类别', '科目'], ['科目均分', '高线率', '中线率', '高线命中率', '中线命中率'], '区')
#
# # 计算选择组合的统计数据
# combination_stats = calculate_combination_statistics(df, standards)
#
# # 保存结果
# school_results.to_excel('school_pass_rate_results.xlsx', index=False)
# county_results.to_excel('county_pass_rate_results.xlsx', index=False)
# city_results.to_excel('city_pass_rate_results.xlsx', index=False)
# combination_stats.to_excel('combination_stats.xlsx', index=False)
#


import pandas as pd

# 读取数据
df = pd.read_excel(r'C:\Users\杰哥\PycharmProjects\pythonProject33\pandas处理学校数据\data\data\调研学生科目得分表.xlsx', dtype={'考号': str, '学籍号': str})
# 将区县和学校列转换为 category 类型以减少内存占用
df['区县'] = df['区县'].astype('category')
df['学校'] = df['学校'].astype('category')
subjects = [
    '总分', '语文原始成绩', '数学原始成绩', '英语原始成绩', '历史原始成绩',
    '物理原始成绩', '政治赋分成绩', '政治原始成绩', '地理赋分成绩',
    '地理原始成绩', '化学赋分成绩', '化学原始成绩', '生物赋分成绩', '生物原始成绩'
]


def determine_selection(row):
    subjects = ['历史原始成绩', '物理原始成绩', '政治原始成绩', '地理原始成绩', '化学原始成绩', '生物原始成绩']
    selection = []
    for subject in subjects:
        if not pd.isna(row[subject]):
            selection.append(subject[:2])
    return ','.join(selection)


# 获取学生选的科目
df['选择组合'] = df.apply(determine_selection, axis=1)

# 定义上线标准
standards = {
    '物理类': {
        '总分': {'high': 490, 'medium': 420},
        '语文原始成绩': {'high': 96, 'medium': 85},
        '数学原始成绩': {'high': 100, 'medium': 90},
        '英语原始成绩': {'high': 95, 'medium': 85},
        '历史原始成绩': {'high': 90, 'medium': 80},
        '物理原始成绩': {'high': 92, 'medium': 82},
        '政治赋分成绩': {'high': 88, 'medium': 78},
        '地理赋分成绩': {'high': 87, 'medium': 77},
        '化学赋分成绩': {'high': 89, 'medium': 79},
        '生物赋分成绩': {'high': 83, 'medium': 73},
    },
    '历史类': {
        '总分': {'high': 490, 'medium': 420},
        '语文原始成绩': {'high': 100, 'medium': 87},
        '数学原始成绩': {'high': 98, 'medium': 88},
        '英语原始成绩': {'high': 97, 'medium': 86},
        '历史原始成绩': {'high': 96, 'medium': 85},
        '物理原始成绩': {'high': 95, 'medium': 84},
        '政治赋分成绩': {'high': 94, 'medium': 83},
        '地理赋分成绩': {'high': 92, 'medium': 81},
        '化学赋分成绩': {'high': 90, 'medium': 79},
        '生物赋分成绩': {'high': 88, 'medium': 77},
    }
}

# 清理无效数据
df[subjects] = df[subjects].apply(pd.to_numeric, errors='coerce')

# 计算选择组合的统计数据
def calculate_statistics(df, group_by_columns, standards):
    results = []

    for group_values, group in df.groupby(group_by_columns):
        if len(group_by_columns) == 3:
            region, school, category = group_values
        elif len(group_by_columns) == 2:
            region, category = group_values
            school = None
        elif len(group_by_columns) == 1:
            category = group_values
            region = school = None

        if category not in standards:
            print(f"Category '{category}' not found in standards.")
            continue

        high_threshold = standards[category]['总分']['high']
        medium_threshold = standards[category]['总分']['medium']

        comb_stats = group.groupby('选择组合').agg(
            选择组合人数=('总分', 'size'),
            高线总分人数=('总分', lambda x: (x >= high_threshold).sum()),
            中线总分人数=('总分', lambda x: (x >= medium_threshold).sum())
        ).reset_index()

        comb_stats['高线总分率'] = comb_stats['高线总分人数'] / comb_stats['选择组合人数']
        comb_stats['中线总分率'] = comb_stats['中线总分人数'] / comb_stats['选择组合人数']
        comb_stats['类别'] = category

        for _, row in comb_stats.iterrows():
            result = {
                **dict(zip(group_by_columns, group_values)),
                '选择组合': row['选择组合'],
                '选择组合人数': row['选择组合人数'],
                '总人数': group.shape[0],
                '占比': row['选择组合人数'] / group.shape[0],
                '高线总分人数': row['高线总分人数'],
                '中线总分人数': row['中线总分人数'],
                '高线总分率': row['高线总分率'],
                '中线总分率': row['中线总分率'],
                '类别': row['类别']
            }
            results.append(result)

    return pd.DataFrame(results)

def add_ranking(df, group_by_cols, rank_cols, prefix):
    for col in rank_cols:
        rank_col_name = f'{prefix}{col}排名'
        df[rank_col_name] = df.groupby(group_by_cols)[col].rank(ascending=False, method='min')

# 计算每个区域、每个学校、每个类别的选择组合的统计数据
school_results = calculate_statistics(df, ['区县', '学校', '类别'], standards)
county_results = calculate_statistics(df, ['区县', '类别'], standards)
city_results = calculate_statistics(df, ['类别'], standards)

# 为学校结果添加排名
add_ranking(school_results, ['类别', '选择组合'], ['选择组合人数', '高线总分率', '中线总分率'], '市')
add_ranking(school_results, ['区县', '类别', '选择组合'], ['选择组合人数', '高线总分率', '中线总分率'], '区')

# 为区县结果添加排名
add_ranking(county_results, ['类别', '选择组合'], ['选择组合人数', '高线总分率', '中线总分率'], '市')
add_ranking(county_results, ['区县', '类别', '选择组合'], ['选择组合人数', '高线总分率', '中线总分率'], '区')

# 保存结果
school_results.to_excel('school_pass_rate_results.xlsx', index=False)
county_results.to_excel('county_pass_rate_results.xlsx', index=False)
city_results.to_excel('city_pass_rate_results.xlsx', index=False)
