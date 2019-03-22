import os
import pandas as pd

# 获得所有的小表格
file_list = []
for root, _, files in os.walk('../各科室表格'):
    for file in files:
        file_list.append(os.path.join(root, file))

print('文件名', file_list)

# 统计天数
person_name, full, ill, work, holiday, other, tran_holiday, extra_work = [], [], [], [], [], [], [], []
for xls_name in file_list:
    print('当前处理表格名字', xls_name)
    file = pd.read_excel(xls_name, sheetname=None, skiprows=[0, 1], index_col=0)  # 返回的是{sheet:dataframe}

    for sheet_name, sheet_value in file.items():
        print('当前处理sheet:', sheet_name)
        for index, person in sheet_value.iterrows():
            person_name.append(person['姓名'])
            if person['全勤（天）'] == '√':
                full.append('√')
                ill.append('')
                work.append('')
                holiday.append('')
                other.append('')
                tran_holiday.append('')
                extra_work.append('')
            else:
                full.append(person['全勤（天）'])
                ill.append(person['病假（天）'])
                work.append(person['事假（天）'])
                holiday.append(person['休假（天）'])
                other.append(person['旷工（天）'])  # 其它这个项目没有对应关系
                tran_holiday.append('')  # 其它这个项目没有对应关系
                extra_work.append(person['加班（小时）'])

data = {'姓名': person_name, '全勤': full, '病假（天）': ill, '事假（天）': work, '休假（天）': holiday, '其它': other,
        '倒休（天）': tran_holiday, '加班（小时）': extra_work}
save_data = pd.DataFrame(data, columns=['姓名', '全勤', '病假（天）', '事假（天）', '休假（天）', '其它',
                                        '倒休（天）', '加班（小时）'])
print(save_data.head())
save_data.to_excel('粗数据.xlsx')
