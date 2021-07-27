import xlrd
import xlwt
from collections import Counter
import pandas as pd

def read_excel():
    # 打开文件
    workBook = xlrd.open_workbook('D:\\PycharmProjects\\flaskProject1\\templates\\xls\\豆瓣电影Top250.xls');

    # 1.获取sheet的名字
    # 1.1 获取所有sheet的名字(list类型)
    allSheetNames = workBook.sheet_names();
    print(allSheetNames);

    # 1.2 按索引号获取sheet的名字（string类型）
    sheet2Name = workBook.sheet_names()[0];
    print(sheet2Name);

    sheet2 = workBook.sheets()[0]  # 获得第1张sheet，索引从0开始
    print(sheet2)
    sheet2_name = sheet2.name  # 获得名称
    sheet2_cols = sheet2.ncols  # 获得列数
    sheet2_nrows = sheet2.nrows  # 获得行数
    print('sheet2 Name: %s\nsheet2 cols: %s\nsheet2 rows: %s' % (sheet2_name, sheet2_cols, sheet2_nrows))

    sheet2_nrows4 = sheet2.row_values(4)  # 获得第4行数据
    sheet2_cols2 = sheet2.col_values(2)  # 获得第2列数据
    cell23 = sheet2.row(2)[5].value  # 查看第3行第6列数据
    print('Row 4: %s\nCol 2: %s\ncell23: %s\n' % (sheet2_nrows4, sheet2_cols2, cell23))

    print("-------------------kkkk-------------------")
    score = []  # 评分
    num = []  # 每个评分所统计出的电影数量
    sheet2 = workBook.sheets()[0]  # 获得第1张sheet，索引从0开始
    score1 = sheet2.col_values(4)  # 获取评分信息
    print(score1)
    print(len(score1)) # 250
    score2 = Counter(score1) # ounter({'8.8': 44, '8.7': 39, '8.9': 30, '8.6': 28})
    print(score2)
    print(len(score2)) # 15
    score3 = sorted(score2.items(), key=lambda dict1: score2[0], reverse=True)
    print(score3) # 按顺序排列
    for item in score3:
        score.append(item[0]) # 键
        num.append(item[1])  # 值


    print(score)
    print(num)

    print("-------------------jjjj-------------------")
    a = [1, 2, 3, 1, 1, 2]
    result = pd.value_counts(a)
    print(result)

def read1():
    '''两个列表合并为字典 示例：'''
    keys = ['a', 'b', 'c','a']
    values = [1, 2, 3,9]
    dictionary = dict(zip(keys, values))
    print(dictionary)

    # 打开文件
    workBook = xlrd.open_workbook('D:\\ProgramFiles\\docTest\excel\\TeamSettlementDetails.xls')
    mouduleName = []
    totalPay = []
    sheet = workBook.sheets()[0]
    mouduleName1 = sheet.col_values(3) # 业务模块
   # print(mouduleName1)
    totalPay1 = sheet.col_values(6) # 结算金额
    # print(totalPay1)
    modulePay = dict(zip(mouduleName1,totalPay1))
    print(modulePay)

    # for item in mouduleName1:

def aa():
    workBook2 = xlrd.open_workbook('D:\\ProgramFiles\\docTest\excel\\TeamSettlementDetails.xls')
    sheet2 = workBook2.sheets()[0]

    aa=Counter(sheet2.col_values(4))
    print(aa) # Counter({'other': 7862, 'catering': 2605, 'ticket': 2486, 'hotel': 1343, 'meeting': 979, 'training': 617, 'guid': 407, 'party': 84})
    moduleName=sorted(set(aa))
    print(moduleName) # ['catering', 'guid', 'hotel', 'meeting', 'other', 'party', 'ticket', 'training']
    print(moduleName[0]) # catering

    otherTotal = 0
    cateringTotal = 0
    ticketTotal = 0
    hotelTotal = 0
    meetingTotal = 0
    trainingTotal = 0
    guidTotal = 0
    partyTotal = 0

    list = []
    sheet2_nrows = sheet2.nrows  # 获得行数
    for i in range(sheet2_nrows):  # 逐行打印sheet2数据
        if sheet2.row_values(i)[4] == 'catering':
            # print(sheet2.row_values(i)[6])
            cateringTotal += sheet2.row_values(i)[6]
        if sheet2.row_values(i)[4] == 'guid':
            # print(sheet2.row_values(i)[6])
            guidTotal += sheet2.row_values(i)[6]
        if sheet2.row_values(i)[4] == 'ticket':
            # print(sheet2.row_values(i)[6])
            ticketTotal += sheet2.row_values(i)[6]
        if sheet2.row_values(i)[4] == 'hotel':
            # print(sheet2.row_values(i)[6])
            hotelTotal += sheet2.row_values(i)[6]
        if sheet2.row_values(i)[4] == 'meeting':
            # print(sheet2.row_values(i)[6])
            meetingTotal += sheet2.row_values(i)[6]
        if sheet2.row_values(i)[4] == 'other':
            # print(sheet2.row_values(i)[6])
            otherTotal += sheet2.row_values(i)[6]
        if sheet2.row_values(i)[4] == 'party':
            # print(sheet2.row_values(i)[6])
            partyTotal += sheet2.row_values(i)[6]
        if sheet2.row_values(i)[4] == 'training':
            # print(sheet2.row_values(i)[6])
            trainingTotal += sheet2.row_values(i)[6]

    print(hotelTotal)

if __name__ == '__main__':
    aa()
   # read_excel();
   #  read1()


'''
        if sheet2.row_values(i)[4] == 'catering':
            # print(sheet2.row_values(i)[6])
            cateringTotal += sheet2.row_values(i)[6]
        if sheet2.row_values(i)[4] == 'hotel':
            # print(sheet2.row_values(i)[6])
            ticketTotal += sheet2.row_values(i)[6]
        if sheet2.row_values(i)[4] == 'meeting':
            # print(sheet2.row_values(i)[6])
            cateringTotal += sheet2.row_values(i)[6]
        if sheet2.row_values(i)[4] == 'other':
            # print(sheet2.row_values(i)[6])
            ticketTotal += sheet2.row_values(i)[6]
'''

'''
 print(sheet2.row_values(i)[4]) # 第4行
 sheet2_nrows4 = sheet2.row_values(4)  # 获得第4行数据
    sheet2_cols2 = sheet2.col_values(2)  # 获得第2列数据
'''


   # a = [1, 2, 3, 1, 1, 2]
    # result = Counter(a)
    # print(result)

    # print("--------------------------------------")
    # nrows = sheet2.nrows
    # print(nrows)
    #
    # print("--------------------------------------")
    # list = [1, 1, 2, 2, 3]
    # print(list)
    # set1 = set(list)
    # print(set1)
    # print(len(set1))  # len(set1)即为列表中不同元素的数量


