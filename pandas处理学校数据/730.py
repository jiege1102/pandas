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
        '总分': {'high': 490, 'medium': 391},
        '语文原始成绩': {'high': 96, 'medium': 85},
        '数学原始成绩': {'high': 91, 'medium': 64},
        '英语原始成绩': {'high': 106, 'medium': 74},
        '历史原始成绩': {'high': 90, 'medium': 80},
        '物理原始成绩': {'high': 50, 'medium': 28},
        '政治赋分成绩': {'high': 68, 'medium': 55},
        '地理赋分成绩': {'high': 74, 'medium': 62},
        '化学赋分成绩': {'high': 68, 'medium': 53},
        '生物赋分成绩': {'high': 71, 'medium': 58},
    },
    '历史类': {
        '总分': {'high': 495, 'medium': 410},
        '语文原始成绩': {'high': 100, 'medium': 87},
        '数学原始成绩': {'high': 77, 'medium': 49},
        '英语原始成绩': {'high': 110, 'medium': 78},
        '历史原始成绩': {'high': 65, 'medium': 50},
        '物理原始成绩': {'high': 95, 'medium': 84},
        '政治赋分成绩': {'high': 89, 'medium': 40},
        '地理赋分成绩': {'high': 68, 'medium': 51},
        '化学赋分成绩': {'high': 75, 'medium': 60},
        '生物赋分成绩': {'high': 77, 'medium': 61},
    }
}

# 清理无效数据
df[subjects] = df[subjects].apply(pd.to_numeric, errors='coerce')



def calculate_pass_rate(df, category, subject, standards):
    thresholds = standards[category][subject]
    high_threshold = thresholds['high']
    medium_threshold = thresholds['medium']

    valid_students = df[df[subject] > 0]
    total_students = len(valid_students)

    avg = valid_students[subject].mean()
    high_pass = valid_students[subject].ge(high_threshold).sum()
    medium_pass = valid_students[subject].ge(medium_threshold).sum()

    high_hit_pass = valid_students[
        valid_students[subject].ge(high_threshold) &
        valid_students['总分'].ge(standards[category]['总分']['high'])
    ].shape[0]
    medium_hit_pass = valid_students[
        valid_students[subject].ge(medium_threshold) &
        valid_students['总分'].ge(standards[category]['总分']['medium'])
    ].shape[0]

    high_pass_rate = high_pass / total_students if total_students > 0 else 0
    medium_pass_rate = medium_pass / total_students if total_students > 0 else 0
    high_hit_pass_rate = high_hit_pass / high_pass if high_pass > 0 else 0
    medium_hit_pass_rate = medium_hit_pass / medium_pass if medium_pass > 0 else 0

    return {
        '参考人数': total_students,
        '科目均分': avg,
        '高线人数': high_pass,
        '高线率': high_pass_rate,
        '高线命中人数': high_hit_pass,
        '高线命中率': high_hit_pass_rate,
        '中线人数': medium_pass,
        '中线率': medium_pass_rate,
        '中线命中人数': medium_hit_pass,
        '中线命中率': medium_hit_pass_rate
    }

def calculate_combination_statistics(df, group_by_columns, standards):
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

        # 计算选择组合统计
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

        # 计算均分、上线率等统计
        for subject in standards[category]:
            pass_rates = calculate_pass_rate(group, category, subject, standards)
            result = {
                **dict(zip(group_by_columns, group_values)),
                '类别': category,
                '科目': subject,
                **pass_rates
            }
            results.append(result)

    return pd.DataFrame(results)

def add_ranking(df, group_by_cols, rank_cols, prefix):
    for col in rank_cols:
        rank_col_name = f'{prefix}{col}排名'
        df[rank_col_name] = df.groupby(group_by_cols)[col].rank(ascending=False, method='min')


def calculate_grade_statistics(df, group_by_columns, grade_columns):
    results = []

    for group_values, group in df.groupby(group_by_columns):
        for subject in grade_columns:
            subject_stats = group[subject].value_counts().sort_index()
            total_students = subject_stats.sum()

            result = {
                **dict(zip(group_by_columns, group_values)),
                '科目': subject,
                '总人数': total_students
            }

            for grade, count in subject_stats.items():
                result[f'{grade}等人数'] = count

            results.append(result)

    return pd.DataFrame(results)



# 定义筛选高线踩线生函数
def mark_high_borderline_students(row, standards):
    if row['类别'] not in standards:
        return 0

    high_threshold = standards[row['类别']]['总分']['high']
    total_score = row['总分']

    if (high_threshold - 10) <= total_score <= (high_threshold + 10):
        return 1 if row['类别'] == '物理类' else 3
    elif (high_threshold - 20) <= total_score <= (high_threshold + 20):
        return 2 if row['类别'] == '物理类' else 4
    else:
        return 0


# 定义筛选中线踩线生函数
def mark_medium_borderline_students(row, standards):
    if row['类别'] not in standards:
        return 0

    medium_threshold = standards[row['类别']]['总分']['medium']
    total_score = row['总分']

    if (medium_threshold - 10) <= total_score <= (medium_threshold + 10):
        return 1 if row['类别'] == '物理类' else 3
    elif (medium_threshold - 20) <= total_score < (medium_threshold - 10) or (medium_threshold + 10) < total_score <= (
            medium_threshold + 20):
        return 2 if row['类别'] == '物理类' else 4
    else:
        return 0


# 在 df 的基础上增加两个辅助列，标记高线和中线踩线生
df['类别'] = df.apply(lambda row: '物理类' if '物理' in row['选择组合'] else '历史类', axis=1)
df['高线踩线生标记'] = df.apply(lambda row: mark_high_borderline_students(row, standards), axis=1)
df['中线踩线生标记'] = df.apply(lambda row: mark_medium_borderline_students(row, standards), axis=1)

required_columns = [
    '区县', '学校', '总分', '语文原始成绩', '数学原始成绩', '英语原始成绩',
    '历史原始成绩', '物理原始成绩', '政治赋分成绩', '地理赋分成绩',
    '化学赋分成绩', '生物赋分成绩', '类别', '学籍号'
]

# 过滤只包含踩线生的数据
borderline_students_df = df[
    (df['高线踩线生标记'] != 0) | (df['中线踩线生标记'] != 0)
]

# 选择所需的列
result_df = borderline_students_df[required_columns]


# 统计函数
def calculate_statistics(group, standards):
    category = group['类别'].iloc[0]
    high_line = standards[category]['总分']['high']
    medium_line = standards[category]['总分']['medium']

    high_10_down = group[(high_line - 10 <= group['总分']) & (group['总分'] < high_line)].shape[0]
    high_10_up = group[(high_line < group['总分']) & (group['总分'] <= high_line + 10)].shape[0]
    high_20_down = group[(high_line - 20 <= group['总分']) & (group['总分'] < high_line)].shape[0]
    high_20_up = group[(high_line < group['总分']) & (group['总分'] <= high_line + 20)].shape[0]

    medium_10_down = group[(medium_line - 10 <= group['总分']) & (group['总分'] < medium_line)].shape[0]
    medium_10_up = group[(medium_line < group['总分']) & (group['总分'] <= medium_line + 10)].shape[0]
    medium_20_down = group[(medium_line - 20 <= group['总分']) & (group['总分'] < medium_line)].shape[0]
    medium_20_up = group[(medium_line < group['总分']) & (group['总分'] <= medium_line + 20)].shape[0]

    # 计算单科未上线人数
    total_students = high_20_down + high_10_up
    high_subject_not_passed = {subject: 0 for subject in standards[category] if subject != '总分'}

    for _, row in group.iterrows():
        # 考虑在高线踩线范围内的学生
        if (high_line - 20 <= row['总分'] <= high_line + 10):
            for subject in high_subject_not_passed:
                if row[subject] < standards[category][subject]['high']:
                    high_subject_not_passed[subject] += 1

    medium_subject_not_passed = {subject: 0 for subject in standards[category] if subject != '总分'}

    for _, row in group.iterrows():
        # 考虑在中线踩线范围内的学生
        if (medium_line - 20 <= row['总分'] <= medium_line + 10):
            for subject in medium_subject_not_passed:
                if row[subject] < standards[category][subject]['medium']:
                    medium_subject_not_passed[subject] += 1

    # 构建结果
    result = {
        '高线下10分人数': high_10_down,
        '高线上10分人数': high_10_up,
        '高线下20分人数': high_20_down,
        '高线上20分人数': high_20_up,
        '中线下10分人数': medium_10_down,
        '中线上10分人数': medium_10_up,
        '中线下20分人数': medium_20_down,
        '中线上20分人数': medium_20_up,
    }

    # 添加单科未上线人数和比例
    for subject, count in high_subject_not_passed.items():
        result[f'高线_{subject}未上线人数'] = count

    for subject, count in medium_subject_not_passed.items():
        result[f'中线_{subject}未上线人数'] = count


    return pd.Series(result)


# 按市、区、校、类别统计
statistics_df = result_df.groupby(['区县', '学校', '类别']).apply(
    calculate_statistics, standards=standards
).reset_index()
statistics_df.to_excel('step_student_tj.xlsx')

# # 计算等级人数统计
# grade_columns = ['政治等级', '地理等级', '生物等级', '化学等级']
# grade_results_school = calculate_grade_statistics(df, ['学校', '类别'], grade_columns)
# grade_results_county = calculate_grade_statistics(df, ['区县', '类别'], grade_columns)
# grade_results_city = calculate_grade_statistics(df, ['类别'], grade_columns)
#
# # 保存等级统计结果
# grade_results_school.to_excel('school_grades_730.xlsx', index=False)
# grade_results_county.to_excel('county_grades_730.xlsx', index=False)
# grade_results_city.to_excel('city_grades_730.xlsx', index=False)

# # 计算每个区域、每个学校、每个类别的选择组合和科目均分等统计数据
# school_results = calculate_combination_statistics(df, ['区县', '学校', '类别'], standards)
# county_results = calculate_combination_statistics(df, ['区县', '类别'], standards)
# city_results = calculate_combination_statistics(df, ['类别'], standards)
#
# # 为学校结果添加排名
# # add_ranking(school_results, ['类别', '选择组合'], ['选择组合人数', '高线总分率', '中线总分率'], '市')
# # add_ranking(school_results, ['区县', '类别', '选择组合'], ['选择组合人数', '高线总分率', '中线总分率'], '区')
# add_ranking(school_results, ['类别', '科目'], ['科目均分', '高线率', '中线率'], '市')
# add_ranking(school_results, ['区县', '类别', '科目'], ['科目均分', '高线率', '中线率', '高线命中率', '中线命中率'], '区')
#
# # 为区县结果添加排名
# # add_ranking(county_results, ['类别', '选择组合'], ['选择组合人数', '高线总分率', '中线总分率'], '市')
# # add_ranking(county_results, ['区县', '类别', '选择组合'], ['选择组合人数', '高线总分率', '中线总分率'], '区')
# add_ranking(county_results, ['类别', '科目'], ['科目均分', '高线率', '中线率'], '市')
# add_ranking(county_results, ['区县', '类别', '科目'], ['科目均分', '高线率', '中线率', '高线命中率', '中线命中率'], '区')
#
# # 为市结果添加排名
# # add_ranking(city_results, ['类别', '选择组合'], ['选择组合人数', '高线总分率', '中线总分率'], '市')
# add_ranking(city_results, ['类别', '科目'], ['科目均分', '高线率', '中线率'], '市')
#
# # 保存结果
# school_results.to_excel('school_730.xlsx', index=False)
# county_results.to_excel('county_730.xlsx', index=False)
# city_results.to_excel('city_730.xlsx', index=False)
