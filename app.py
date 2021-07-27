from flask import Flask,render_template
import xlrd
import xlwt
from collections import Counter
import pandas as pd

app = Flask(__name__)


@app.route('/')
def index():
    return render_template("index.html")


@app.route('/index')
def home():
    #return render_template("index.html")
    return index()


@app.route('/score')
def score():
    # 打开文件
    workBook = xlrd.open_workbook('D:\\PycharmProjects\\flaskProject1\\templates\\xls\\豆瓣电影Top250.xls');

    score = []  # 评分
    num = []  # 每个评分所统计出的电影数量
    sheet1 = workBook.sheets()[0]  # 获得第1张sheet，索引从0开始
    score1 = sheet1.col_values(4)  # 获取评分信息
    score2 = Counter(score1)
    score3 = sorted(score2.items(), key=lambda dict1: score2[0])
    for item in score3:
        score.append(item[0])
        num.append(item[1])



    workBook1 = xlrd.open_workbook('D:\\ProgramFiles\\docTest\excel\\TeamSettlementDetails.xls')
    sheet1 = workBook1.sheets()[0]

    aa = Counter(sheet1.col_values(4))
    moduleName = []
    # Counter({'other': 7862, 'catering': 2605, 'ticket': 2486, 'hotel': 1343, 'meeting': 979, 'training': 617, 'guid': 407, 'party': 84})
    moduleName = sorted(set(aa))

    otherTotal = 0
    cateringTotal = 0
    ticketTotal = 0
    hotelTotal = 0
    meetingTotal = 0
    trainingTotal = 0
    guidTotal = 0
    partyTotal = 0

    list = []
    sheet1_nrows = sheet1.nrows  # 获得行数
    for i in range(sheet1_nrows):  # 逐行打印sheet1数据
        if sheet1.row_values(i)[4] == 'catering':
            # print(sheet1.row_values(i)[6])
            cateringTotal += sheet1.row_values(i)[6]
        if sheet1.row_values(i)[4] == 'guid':
            # print(sheet1.row_values(i)[6])
            guidTotal += sheet1.row_values(i)[6]
        if sheet1.row_values(i)[4] == 'ticket':
            # print(sheet1.row_values(i)[6])
            ticketTotal += sheet1.row_values(i)[6]
        if sheet1.row_values(i)[4] == 'hotel':
            # print(sheet1.row_values(i)[6])
            hotelTotal += sheet1.row_values(i)[6]
        if sheet1.row_values(i)[4] == 'meeting':
            # print(sheet1.row_values(i)[6])
            meetingTotal += sheet1.row_values(i)[6]
        if sheet1.row_values(i)[4] == 'other':
            # print(sheet1.row_values(i)[6])
            otherTotal += sheet1.row_values(i)[6]
        if sheet1.row_values(i)[4] == 'party':
            # print(sheet1.row_values(i)[6])
            partyTotal += sheet1.row_values(i)[6]
        if sheet1.row_values(i)[4] == 'training':
            # print(sheet1.row_values(i)[6])
            trainingTotal += sheet1.row_values(i)[6]




    return render_template("score.html", score=score, num=num, moduleName=moduleName, cateringTotal=cateringTotal,
                           guidTotal=guidTotal,
                           ticketTotal=ticketTotal, hotelTotal=hotelTotal, meetingTotal=meetingTotal,
                           otherTotal=otherTotal, partyTotal=partyTotal, trainingTotal=trainingTotal)



if __name__ == '__main__':
    app.run()


