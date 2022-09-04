"""
@author： jmchen
@time： 2022/8/27
"""
import pandas as pd
import warnings
# import datacompy
from time import sleep
import tkinter as tk
from tkinter import filedialog

warnings.filterwarnings('ignore')


def add_new_col(raw_data, check_data):
    raw = pd.read_excel(raw_data)
    check = pd.read_excel(check_data)
    cls_raw = set(raw['分类'])
    cls_check = list(check.columns)
    begin_idx = cls_check.index('流程图') + 1
    for label in cls_raw:
        if label not in cls_check and label not in ['IO表', '3D总装图', '电气接线图']:
            # 增加这列
            check.insert(begin_idx, label, value='×', allow_duplicates=False)
            begin_idx += 1
    # check.to_excel('./data/temp.xls')
    return check


def auto_write(raw_data, check_data, new_data):
    raw = pd.read_excel(raw_data)
    check = check_data

    all_labels = list(check.columns)
    first_check_label_idx = all_labels.index('所有者') + 1
    end_check_label_idx = all_labels.index('工程师')
    idx_ipi_5 = all_labels.index('IPI-5-电气接线图(含IO表)')
    idx_ipi_6 = all_labels.index('IPI-6-装配图（含气路图/FAT打表表）')

    # print('所有列名：', all_labels)
    check_label = all_labels[first_check_label_idx: end_check_label_idx]
    print('所有分类：', check_label, end='\n\n')
    # 获取原数据的分类和项目编号
    cls = list(raw['分类'])
    numbers = list(raw['项目编号'])

    num_cls_dict = dict()
    for idx, num in enumerate(numbers):
        if num not in num_cls_dict.keys():
            num_cls_dict[num] = []
        num_cls_dict[num].append(cls[idx])

    # 在新表中判断和写入
    for idx, num in enumerate(check['项目号']):
        if num not in num_cls_dict.keys():
            # 项目编号不在的情况
            # print(num, end=' ')
            continue
        else:
            have_labels = num_cls_dict[num]

            for each_label in check_label:
                if '电气接线图' in have_labels or 'IO表' in have_labels:
                    check.iloc[idx, idx_ipi_5] = '√'
                if '3D总装图' in have_labels:
                    check.iloc[idx, idx_ipi_6] = '√'
                if each_label in have_labels:
                    # check.iloc[idx, first_check_label_idx + check_label.index(each_label)] = '√'
                    check.iloc[idx, all_labels.index(each_label)] = '√'
                if each_label not in have_labels:
                    # print(each_label, idx)
                    check.iloc[idx, all_labels.index(each_label)] = '×'

    # check.to_excel('./data/res.xlsx')
    check.to_excel(new_data)

    print('新文件写入完毕\n\n')


'''

def check_is_equal(data2_path, data3_path):
    data2 = pd.read_excel(data2_path)
    data3 = pd.read_excel(data3_path).iloc[:, 1:]
    d3_labels = list(data3.columns)
    # print(d3_labels)
    begin_label = d3_labels.index('所有者') + 1
    end_label = d3_labels.index('流程图') + 1
    data3 = data3.iloc[:, begin_label: end_label]
    data2 = data2.iloc[:, begin_label: end_label]

    compy = datacompy.Compare(data2, data3, on_index=True)
    print(compy.matches())
    print(compy.report())
    # print(data2 == data3)
    # a = data2.equals(data3)
    # print(a)
    # a = data2.append(data3).drop_duplicates(keep=False)
    # print(a)
    # for idx, v in a.iterrows():
    #     print(list(v))
    #     break
    # print(f'宝，有 {len(a)} 行是不一样的')
'''


def choose_file():
    # 实例化
    root = tk.Tk()
    root.withdraw()
    # 获取文件夹路径
    f_path = filedialog.askopenfilename()
    return f_path


def interface_UI():
    while True:
        # data1_path, data2_path, data3_path = 0, 0, 0
        print('#*=' * 15, 'help_cc', '=*#' * 15)
        print('#*=' * 16, '=*#' * 16, end='\n\n')
        op1 = input('请输入 c 来选择原来完整的文件：')
        if op1 == 'c':
            data1_path = choose_file()
            print(data1_path)
        else:
            print('输入其他值是无效的，请重新输入c')
            continue
        op2 = input('请输入 q 来选择需要填写的文件：')
        if op2 == 'q':
            data2_path = choose_file()
            print(data2_path)
        else:
            print('输入其他值是无效的，请重新输入q')
            continue
        op3 = input('请输入 l 来输入你要保存的文件名：')
        if op3 == 'l':
            save_file_name = input('请输入文件名称，然后按回车：')
            # data3_path = data1_path[:data1_path.rfind('/') + 1] + save_file_name + '.xls'
            data3_path = data1_path[:data1_path.rfind('/') + 1] + save_file_name + '.xlsx'
            print(data3_path)
        else:
            print('输入其他值是无效的，请重新输入文件名')
            continue

        check_data = add_new_col(data1_path, data2_path)
        auto_write(data1_path, check_data, data3_path)
        find_not_in(data1_path, data2_path)

        statue = input('\n\n若要退出请按 pp，再回车；继续生成文件请按任意键: ')
        if statue == 'pp':
            print('程序结束......')
            for i in range(5):
                sleep(1)
                print(f'程序将在{5 - i}秒后退出程序')
            print('已结束程序')
            exit(0)
        else:
            continue


def find_not_in(raw_data, new_data):
    raw = pd.read_excel(raw_data)
    new = pd.read_excel(new_data)
    raw_number = set(raw['项目编号'])
    new_number = set(new['项目号'])
    x, y = raw_number - new_number, new_number - raw_number
    print(f'原来有的，新的没有的编号有 {len(x)}人：\n', x, '\n\n')
    print(f'原来没有的，新的有的编号有 {len(y)}人：\n', y, '\n\n')
    with open('./系统导出有新建无.txt', 'w', encoding='utf-8') as f:
        for num in x:
            f.write(num + '\n')
    with open('系统导出无新建有.txt', 'w', encoding='utf-8') as f:
        for num in y:
            f.write(num + '\n')
    print('txt保存完毕')


if __name__ == '__main__':
    try:
        print('程序启动......\n\n')
        interface_UI()
    except:
        print('若出现问题，请联系张学友')
        sleep(5)
