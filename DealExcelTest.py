import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

filename = 'd:/3/1.xlsx'


#处理excel文件
def deal_excel(filename):
    df = pd.read_excel(filename,header=1)
    print(df)
    # 使用groupby和sum进行聚合，并使用tsransform广播总和值
    df_group = df.groupby(['品名', '规格'])
    df['合计'] = df_group['数量'].transform('sum')
    needGroupIdxArray = df_group.groups
    # 获取所有的value并存储在一个列表中
    values = [[index.min(), index.max()] for index in needGroupIdxArray.values()]
    # 创建一个Excel工作簿
    wb = Workbook()
    ws = wb.active
    # 将DataFrame中的数据写入工作表
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        ws.append(row)
    # 保存需要合并单元格的范围
    merge_ranges = []
    for value in values:
        min_value = value[0]
        max_value = value[1]
        merge_ranges.append((min_value + 2, max_value + 2))  # +2是因为在Excel中行索引从1开始，并且有标题行
    # 合并单元格
    for merge_range in merge_ranges:
        min_row, max_row = merge_range
        ws.merge_cells(start_row=min_row, start_column=16, end_row=max_row, end_column=16)  # '合计'列是第7列
    # 保存结果为xlsx文件
    okFileName = filename.replace('.xlsx', '_完成流向.xlsx')
    wb.save(okFileName)

deal_excel(filename)