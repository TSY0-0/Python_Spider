# -*- codeing = utf-8 -*-
# @Time : 2021/6/30 11:14
# @Author : Tancy
# @File : 选课.py
# @Software : PyCharm


import requests
from lxml import etree
import xlwt


def inXuanke(session,page,n):
    url = "http://jwcnew.nefu.edu.cn/dblydx_jsxsd/xkgl/jxkcList"
    data = {
        "xnxqid":"2021-2022-1",
        "kkyx":"",
        "kcmc":"",
        "skjs":"null",
        "skzy":"",
        "kcsx":4,
        "sknj":"",
        "skyx":"",
        "pageIndex":page
    }
    html = session.post(url,data=data)
    tree = etree.HTML(html.text)
    divs = tree.xpath('//*[@id="Form1"]/table/tr')
    i=1
    l=[]
    for div in divs:
        j = 0
        if i==1:
            i=i+1
        else:
            num = div.xpath('./td[1]/text()')[0]
            sheet.write(n, j, num)
            j = j + 1
            code = div.xpath('./td[2]/text()')[0]
            sheet.write(n, j, code)
            j = j + 1
            name = div.xpath('./td[3]/text()')[0]
            sheet.write(n, j, name)
            j = j + 1
            teacher = (div.xpath('./td[4]/a/text()'))
            if len(teacher)!=0:
                teacher = ",".join((div.xpath('./td[4]/a/text()')))#去逗号
                sheet.write(n, j, teacher)
            else:
                sheet.write(n, j, " ")
            # print(teacher)
            j = j + 1
            time = div.xpath('./td[5]/text()')
            if len(time)!=0:
                time = ",".join((div.xpath('./td[5]/text()')))#去逗号
                sheet.write(n, j, time)
            else:
                sheet.write(n, j, " ")
            # print(time)
            j = j + 1
            classroom = div.xpath('./td[6]/text()')
            if len(classroom)!=0:
                classroom = ",".join((div.xpath('./td[6]/text()')))#去逗号
                sheet.write(n, j, classroom)
            else:
                sheet.write(n, j, " ")
            j = j + 1
            week = div.xpath('./td[7]/text()')
            if len(week)!=0:
                week = (div.xpath('./td[7]/text()'))#去逗号
                sheet.write(n, j, week)
            else:
                sheet.write(n, j, " ")
            j = j + 1
            calss = div.xpath('./td[8]/text()')
            if len(calss)!=0:
                calss = (div.xpath('./td[8]/text()'))#去逗号
                sheet.write(n, j, calss)
            else:
                sheet.write(n, j, " ")
            j = j + 1
            examine = div.xpath('./td[9]/text()')[0]
            sheet.write(n, j, examine)
            j = j + 1
            requirements = div.xpath('./td[10]/text()')
            if len(requirements)!=0:
                requirements = (div.xpath('./td[10]/text()'))#去逗号
                sheet.write(n, j, requirements)
            else:
                sheet.write(n, j, " ")
            j = j + 1
            a11 = div.xpath('./td[11]/text()')[0]
            sheet.write(n, j, a11)
            j = j + 1
            i=i+1
            n=n+1
            # print(time)
    book.save("选课信息.xls")

if __name__ == '__main__':
    session = requests.session()
    data = {
        "USERNAME": "2019214384",
        "PASSWORD": "1329971050TSY"
    }
    url = "http://jwcnew.nefu.edu.cn/dblydx_jsxsd/xk/LoginToXk"
    resp = session.post(url,data=data,timeout=1000)
    book = xlwt.Workbook(encoding="utf-8",style_compression=0)  #创建对象
    sheet = book.add_sheet('选课信息',cell_overwrite_ok=True)  # 创建工作表
    col = ("序号", "课程编号", "课程名字", "上课教师", "上课时间", "上课地点", "开课周次", "上课班级", "考核方式", "考试要求", "学分")
    for i in range(0, 11):
        sheet.write(0, i, col[i])  # 列名
    n=1
    for page in range(1,24):
        inXuanke(session,page,n)
        n=n+20
        print("给爷爬！")
    print("完事！")


