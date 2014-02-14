#coding=utf8
#!/usr/bin/env python3.3
__author__ = 'Yoga'


import urllib
import urllib.request
import urllib.parse
import time
import re
import xlwt


#创建头部
def create_header(cookie, url):
    COOKIE = cookie
    HEADERS = {
    # "Host": "jwgl.avceit.cn:58000",
    "Referer": url + "wsxk/stu_zxjg_rpt.aspx",
    "User-Agent": "Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/32.0.1700.107 Safari/537.36",
    "Cookie": COOKIE}
    return HEADERS, url


#写入Excel文件数据
def write_execl(rows, columns, text, sheet):
    sheet.write(rows, columns, text)


#登陆并收集数据
def test_login(TeacherID, Rows, Columns, Sheet, HEADERS):
    try:
        print("链接ID：" + TeacherID)
        url = HEADERS[1] + 'JXZY/INFO_Teacher.aspx?id=' + TeacherID
        req = urllib.request.Request(url, headers=HEADERS[0])
        text = urllib.request.urlopen(req).read()
        text = text.decode('gbk')
        text = text.replace("&nbsp;", "")
        text = text.replace(" ", "")
        text = text.replace("\r", "")
        text = text.replace("\n", "")
        #正则采集信息
        #工号
        Id_reg = re.compile('<tdwidth=\"100\"class=B>(.*?)</td>')
        Id = Id_reg.search(text)
        #姓名
        neme_reg = re.compile('<tdwidth=\"100\"class=T>(.*?)</td>')
        name = neme_reg.search(text)
        #性别
        sex_reg = re.compile("<tdclass=T>(.*?)<br></td>")
        sex = sex_reg.search(text)
        #出生日期
        bob_reg = re.compile('<tdclass=Talign=topheight=\"24\">(.*?)年(.*?)月(.*?)日<br></td></tr>')
        bob = bob_reg.search(text)
        #学历
        edu_bg_reg = re.compile('<tdclass=Talign=top>(.*?)<br></td>')
        edu_bg = edu_bg_reg.search(text)
        #学位
        degree_reg = re.compile('<tdclass=Talign=top>(.*?)<br></td>')
        degree = degree_reg.findall(text)
        #职称
        title_reg = re.compile('<tdclass=T>(.*?)<br></td>')
        title = title_reg.findall(text)
        #入校年份
        enter_reg = re.compile('<tdclass=B>(.*?)年(.*?)月(.*?)日<br></td></tr>')
        enter = enter_reg.search(text)
        #民族
        nation_reg = re.compile("<tdclass=Talign=top>(.*?)<br></td>")
        nation = nation_reg.findall(text)
        #身份证号码
        Id_card_reg = re.compile('<tdclass=Talign=leftcolspan=3>(.*?)<br></td></tr>')
        Id_card = Id_card_reg.search(text)
        #籍贯
        hometown_reg = re.compile('<tdclass=Balign=top>(.*?)<br></td>')
        hometown = hometown_reg.search(text)
        #岗位
        job_reg = re.compile('<tdclass=B>(.*?)<br>')
        job = job_reg.search(text)
        #是否在岗
        Is_job_reg = re.compile('<tdclass=Balign=center>(.*?)</td></tr>')
        Is_job = Is_job_reg.search(text)
        #联系电话
        num_reg = re.compile('<tdclass=BTalign=top>(.*?)<br></td>')
        num = num_reg.search(text)
        #手机号码
        tel_reg = re.compile('<tdclass=BTalign=topcolspan=5>(.*?)<br></td></tr>')
        tel = tel_reg.search(text)
        #电子邮箱
        mail_reg = re.compile('<tdclass=BTalign=topcolspan=5>(.*?)<br></td></tr>')
        mail = mail_reg.findall(text)
        #简历
        resume_reg = re.compile('<textarea(.*)>(.+?)</textarea>')
        resume = resume_reg.findall(text)

        #原始网页数据存储到TXT文件
        # T_file = open(r'F:\Teacher2.txt', 'a')
        # print(text, file=T_file)
        # T_file.close()

        #判断整个页面是否为空，若遇到工号为空页面停止采集
        if Id:
            print("链接ID：" + TeacherID)
            write_execl(Rows, Columns, TeacherID, Sheet)
            Columns = Columns + 1
            print("工号：" + Id.group(1))
            write_execl(Rows, Columns, Id.group(1), Sheet)
            Columns = Columns + 1
            print("姓名：" + name.group(1))
            write_execl(Rows, Columns, name.group(1), Sheet)
            Columns = Columns + 1
            print("性别：" + sex.group(1))
            write_execl(Rows, Columns, sex.group(1), Sheet)
            Columns = Columns + 1
            if bob:
                print("出身日期：" + bob.group(1) + "/" + bob.group(2) + "/" + bob.group(3))
                write_execl(Rows, Columns, bob.group(1) + "/" + bob.group(2) + "/" + bob.group(3), Sheet)
            else:
                print("出生日期：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if edu_bg.group(1):
                print("学历：" + edu_bg.group(1))
                write_execl(Rows, Columns, edu_bg.group(1), Sheet)
            else:
                print("学历：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if degree[1]:
                print("学位：" + degree[1])
                write_execl(Rows, Columns, degree[1], Sheet)
            else:
                print("学位：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if title[1]:
                print("职称：" + title[1])
            else:
                print("职称：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if enter:
                print("入校年份：" + enter.group(1) + "/" + enter.group(2) + "/" + enter.group(3))
                write_execl(Rows, Columns, enter.group(1) + "/" + enter.group(2) + "/" + enter.group(3), Sheet)
            else:
                print("入校年份：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if nation[2]:
                print("民族：" + nation[2])
                write_execl(Rows, Columns, nation[2], Sheet)
            else:
                print("民族：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if Id_card.group(1):
                print("身份证号码：" + Id_card.group(1))
                write_execl(Rows, Columns, Id_card.group(1), Sheet)
            else:
                print("身份证号码：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if hometown.group(1):
                print("籍贯：" + hometown.group(1))
                write_execl(Rows, Columns, hometown.group(1), Sheet)
            else:
                print("籍贯：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if job.group(1):
                print("岗位：" + job.group(1))
                write_execl(Rows, Columns, job.group(1), Sheet)
            else:
                print("岗位：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if Is_job.group(1):
                print("是否在岗：" + Is_job.group(1))
                write_execl(Rows, Columns, Is_job.group(1), Sheet)
            else:
                print("是否在岗：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if num.group(1):
                print("联系电话：" + num.group(1))
                write_execl(Rows, Columns, num.group(1), Sheet)
            else:
                print("联系电话：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if tel.group(1):
                print("手机号码：" + tel.group(1))
                write_execl(Rows, Columns, tel.group(1), Sheet)
            else:
                print("手机号码：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if mail[1]:
                print("电子邮箱：" + mail[1])
                write_execl(Rows, Columns, mail[1], Sheet)
            else:
                print("电子邮箱：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            if resume:
                print("简历：" + str(resume[0][1]))
                write_execl(Rows, Columns, str(resume[0][1]), Sheet)
            else:
                print("简历：无")
                write_execl(Rows, Columns, "无", Sheet)
            Columns = Columns + 1
            # print(text)
            print("\n")
            Rows = Rows + 1
            return Rows
        else:
            return Rows
    except:
        print("error")


#写入Excel表格头部
def create_excel(Rows, Columns, sheet):
    write_execl(Rows, Columns, "链接ID", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "工号", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "姓名", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "性别", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "出生日期", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "学历", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "学位", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "职称", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "入校年份", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "民族", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "身份证号码", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "籍贯", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "岗位", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "是否在岗", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "联系电话", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "手机号码", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "电子邮箱", sheet)
    Columns = Columns + 1
    write_execl(Rows, Columns, "简历", sheet)
    Columns = Columns + 1


if __name__ == '__main__':
    url = input("请输入教育管理系统界面登陆地址，例如：http://jwgl.avceit.cn:58000/jwweb/ （一般情况下路径是jwweb，切记地址最后要带有“/”）：")
    cookie = input("请输入您的cookie\n"
                   "（注意请使用浏览器查看登陆之后的cookie，\n"
                   "例如：Chrome浏览器，按下F12键，点击Network，"
                   "就拿登陆后的第一页为例，选择MAINFRM.aspx，即可在Headers中查看到Cookie值\n"
                   "如Cookie:ASP.NET_SessionId=qmwlpi45mpazhe5*********）：")
    path = input("请输入数据文件保存目录（文件夹地址需要真实存在，如：F:\  切记地址最后要带有“\”)")
    #创建头部信息
    HEADERS = create_header(cookie, url)
    #初始化链接ID
    TeacherID = "0000001"
    #初始化教师记录计数器
    Check = 0
    #初始化行号列号
    Rows = 0
    Columns = 0
    #ID计数器
    TID = 0
    #xlvt模块创建excel工作薄
    wbk = xlwt.Workbook()
    #创建excel表
    sheet = wbk.add_sheet("Sheet 1", cell_overwrite_ok=True)
    #初始化Excel头部
    create_excel(Rows, Columns, sheet)
    #增加行数
    Rows = 1
    Columns = 0
    #文件名计数器
    file_number = 0
    while TID <= 900:
        LID = TID
        if TID < 10:
            TeacherID = '000000' + str(LID)
        elif TID >= 10 and TID < 100:
            TeacherID = '00000' + str(LID)
        elif TID >= 100:
            TeacherID = '0000' + str(LID)
        Rows = test_login(TeacherID, Rows, Columns, sheet, HEADERS)
        TID = TID + 1
        time.sleep(0.5)
        if Rows == 50:
            Check = Check + Rows
            wbk.save(path + "Teacher" + str(file_number) + ".xls")
            wbk = xlwt.Workbook()
            Rows = 0
            sheet = wbk.add_sheet("Sheet 1", cell_overwrite_ok=True)
            create_excel(Rows, Columns, sheet)
            Rows = 1
            Columns = 0
            file_number = file_number + 1
    Check = Check + Rows
    print("总共记录了" + str(Check) + "名教师")
    wbk.save(path + "Teacher" + str(file_number) + ".xls")


