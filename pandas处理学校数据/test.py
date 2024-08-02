import pandas as pd

# 读取数据 将区县和学校列转换为 category 类型以减少内存占用
df = pd.read_excel(r'C:\Users\杰哥\PycharmProjects\pythonProject33\pandas处理学校数据\data\data\调研学生科目得分表.xlsx', dtype={'考号': str, '学籍号': str})
# 将区县和学校列转换为 category 类型以减少内存占用
df['区县'] = df['区县'].astype('category')
df['学校'] = df['学校'].astype('category')
subjects = [
    '总分', '语文原始成绩', '数学原始成绩', '英语原始成绩', '历史原始成绩',
    '物理原始成绩', '政治赋分成绩', '政治原始成绩', '地理赋分成绩',
    '地理原始成绩', '化学赋分成绩', '化学原始成绩', '生物赋分成绩', '生物原始成绩'
]

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


def determine_selection(row):
    subjects = ['历史原始成绩', '物理原始成绩', '政治原始成绩', '地理原始成绩', '化学原始成绩', '生物原始成绩']
    selection = []
    for subject in subjects:
        if not pd.isna(row[subject]):
            selection.append(subject[:2])
    return ','.join(selection)


# 获取学生选的科目
df['选科组合'] = df.apply(determine_selection, axis=1)

# 清理无效数据
df[subjects] = df[subjects].apply(pd.to_numeric, errors='coerce')


######################################################
#     计算市、区、校的每个选科组合的人数、占比、上线人数
######################################################

# 创建一个函数来计算每个选科组合的统计数据
def calculate_statistics(df, category):
    # 筛选类别
    df_category = df[df['类别'] == category]

    # 分组统计
    group = df_category.groupby('选科组合').agg(人数=('总分', 'size')).reset_index()

    # 计算人数占比
    total_students = len(df_category)
    group['人数占比'] = group['人数'] / total_students

    # 计算总分上高线和中线的人数
    high_line = standards[category]['总分']['high']
    medium_line = standards[category]['总分']['medium']

    # 计算总分上高线的人数
    # 筛选总分大于等于高线的学生
    high_line_students = df_category[df_category['总分'] >= high_line]
    # 按选科组合进行分组并统计每组上线的人数
    high_line_counts = high_line_students.groupby('选科组合')['总分'].size()
    # 将统计结果重新索引以匹配所有选科组合，并用0填充缺失值
    high_line_counts_reindexed = high_line_counts.reindex(group['选科组合']).fillna(0)
    # 将统计结果转换为整数
    group['总分上高线人数'] = high_line_counts_reindexed.astype(int).values

    # 计算总分上中线的人数
    # 筛选总分大于等于中线的学生
    medium_line_students = df_category[df_category['总分'] >= medium_line]
    # 按选科组合进行分组并统计每组上线的人数
    medium_line_counts = medium_line_students.groupby('选科组合')['总分'].size()
    # 将统计结果重新索引以匹配所有选科组合，并用0填充缺失值
    medium_line_counts_reindexed = medium_line_counts.reindex(group['选科组合']).fillna(0)
    # 将统计结果转换为整数
    group['总分上中线人数'] = medium_line_counts_reindexed.astype(int).values

    return group


# 计算全市的统计数据
physics_stats_city = calculate_statistics(df, '物理类')
history_stats_city = calculate_statistics(df, '历史类')

# 按区县计算统计数据并存储为DataFrame对象
county_stats_list = []
for county, df_county in df.groupby('区县'):
    physics_stats_county = calculate_statistics(df_county, '物理类')
    history_stats_county = calculate_statistics(df_county, '历史类')

    physics_stats_county['区县'] = county
    history_stats_county['区县'] = county

    county_stats_list.append(physics_stats_county)
    county_stats_list.append(history_stats_county)

county_stats_df = pd.concat(county_stats_list, ignore_index=True)

# 按学校计算统计数据并存储为DataFrame对象
school_stats_list = []
for school, df_school in df.groupby('学校'):
    physics_stats_school = calculate_statistics(df_school, '物理类')
    history_stats_school = calculate_statistics(df_school, '历史类')

    physics_stats_school['学校'] = school
    history_stats_school['学校'] = school

    school_stats_list.append(physics_stats_school)
    school_stats_list.append(history_stats_school)

school_stats_df = pd.concat(school_stats_list, ignore_index=True)

# 输出结果
# print("全市物理类统计数据：")
# print(physics_stats_city)
# print("\n全市历史类统计数据：")
# print(history_stats_city)

# print("\n按区县统计数据：")
# print(county_stats_df)
county_stats_df.to_excel('0.xlsx', index=False)

# print("\n按学校统计数据：")
# print(school_stats_df)




######################################################
#     计算市、区、校的参考人数、平均分、上线人数和上线率
######################################################

# 计算单科参考人数、平均分、上线人数和上线率的函数
def calculate_pass_rate(df, category, subject, standards):
    high_threshold = standards[category][subject]['high']
    medium_threshold = standards[category][subject]['medium']

    # 只计算分数大于0分的学生人数
    valid_students = df[df[subject] > 0]
    total_students = len(valid_students)

    avg = valid_students[subject].mean()
    # 单科上线人数
    high_pass = len(valid_students[valid_students[subject] >= high_threshold])
    medium_pass = len(valid_students[valid_students[subject] >= medium_threshold])
    # 单科和总分同时上线人数
    high_hit_pass = len(valid_students[(valid_students[subject] >= high_threshold) & (valid_students['总分'] >= standards[category]['总分']['high'])])
    medium_hit_pass = len(valid_students[(valid_students[subject] >= medium_threshold) & (valid_students['总分'] >= standards[category]['总分']['medium'])])
    # 单科上线率
    high_pass_rate = high_pass / total_students if total_students > 0 else 0
    medium_pass_rate = medium_pass / total_students if total_students > 0 else 0
    # 单科命中率
    high_hit_pass_rate = high_hit_pass / high_pass if high_pass > 0 else 0
    medium_hit_pass_rate = medium_hit_pass / medium_pass if medium_pass > 0 else 0

    return total_students, avg, high_pass, high_pass_rate, high_hit_pass, high_hit_pass_rate, medium_pass, medium_pass_rate, medium_hit_pass, medium_hit_pass_rate


# 计算每个区域、每个学校、每个类别的每个科目的上线人数和上线率
def calculate_results(df, group_by_columns):
    results = []
    for group_keys, group in df.groupby(group_by_columns):
        if len(group_by_columns) == 3:
            region, school, category = group_keys
        elif len(group_by_columns) == 2:
            region, category = group_keys
            school = None
        else:
            category = group_keys[0]
            region = school = None

        # Ensure category is not None and is a valid key
        if category not in standards:
            continue

        for subject in standards[category]:
            total_students, avg, high_pass, high_pass_rate, high_hit_pass, high_hit_pass_rate, medium_pass, medium_pass_rate, medium_hit_pass, medium_hit_pass_rate = calculate_pass_rate(
                group,
                category,
                subject,
                standards)
            results.append({
                '区县': region,
                '学校': school,
                '类别': category,
                '科目': subject,
                '参考人数': total_students,
                '科目均分': avg,
                '高线人数': high_pass,
                '高线率': high_pass_rate,
                '高线命中人数': high_hit_pass,
                '高线命中率': high_hit_pass_rate,
                '中线人数': medium_pass,
                '中线率': medium_pass_rate,
                '中线命中人数': medium_hit_pass,
                '中线命中率': medium_hit_pass_rate,
            })

    return pd.DataFrame(results)


# 计算结果
school_results_df = calculate_results(df, ['区县', '学校', '类别'])
district_results_df = calculate_results(df, ['区县', '类别'])
city_results_df = calculate_results(df, ['类别'])


# 计算排名函数
def add_ranking(df, group_by_cols, rank_cols, prefix):
    for col in rank_cols:
        rank_col_name = f'{prefix}{col}排名'
        df[rank_col_name] = df.groupby(group_by_cols)[col].rank(ascending=False, method='min')


# 为学校结果添加排名
add_ranking(school_results_df, ['类别', '科目'], ['科目均分', '高线率', '中线率'], '市')
add_ranking(school_results_df, ['区县', '类别', '科目'], ['科目均分', '高线率', '中线率', '高线命中率', '中线命中率'], '区')

# 为区县结果添加排名
add_ranking(district_results_df, ['类别', '科目'], ['科目均分', '高线率', '中线率'], '市')
add_ranking(district_results_df, ['区县', '类别', '科目'], ['科目均分', '高线率', '中线率', '高线命中率', '中线命中率'], '区')
print(city_results_df)
# 保存结果
# school_results_df.to_excel('school_731.xlsx', index=False)
# district_results_df.to_excel('district_731.xlsx', index=False)
# city_results_df.to_excel('city_731.xlsx', index=False)
