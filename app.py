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



    return render_template("score.html",score= score,num=num)



if __name__ == '__main__':
    app.run()


