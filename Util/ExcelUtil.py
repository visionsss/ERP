# -*- coding:utf-8 -*-
"""
@author: Darcy
@time: 2018/9/25 12:42
"""
import Util.FileUitl

"""
    复制模板excel表
    Parm:
        modebook 模板文件
        Usebook 复制文件
    Return:
        modebook 复制完成文件
"""
def CopyFromModeExcel(modeBook, UseBook):
    sheets = UseBook.sheetnames
    for i in range(1,len(sheets)):
        ws1 = modeBook[modeBook.sheetnames[0]]
        ws2 = modeBook.copy_worksheet(ws1)
        ws2.title = sheets[i]
    return modeBook

"""
    复制模板excel表 标准养护混凝土抗压强度计算表特供版 少于10
    Parm:
        modebook 模板文件
        Usebook 复制文件
    Return:
        modebook 复制完成文件
"""
def CopyFromModeExcelFor4LessTen(modeBook, UseBook, i):
    sheets = UseBook.sheetnames
    ws1 = modeBook[modeBook.sheetnames[0]]
    ws2 = modeBook.copy_worksheet(ws1)
    ws2.title = sheets[i]
    return modeBook

"""
    复制模板excel表 标准养护混凝土抗压强度计算表特供版 少于多于
    Parm:
        modebook 模板文件
        Usebook 复制文件
    Return:
        modebook 复制完成文件
"""
def CopyFromModeExcelFor4MoreTen(modeBook, UseBook, i):
    sheets = UseBook.sheetnames
    ws1 = modeBook[modeBook.sheetnames[1]]
    ws2 = modeBook.copy_worksheet(ws1)
    ws2.title = sheets[i]
    return modeBook

"""
    判断安定性是否合格
    Parm:
        Stablity 
    Return:
        1 合格
        0 不合格
"""
def CemAtrIsStablility(Stablility):
    if(Stablility != 1):
        return '不合格'
    else:
        return '合格'


"""
    获取混凝土试块强度试验序号
    Parm:
        ConStrengReportSum Sheet
    Return:
        IdNum 序号个数
"""
def GetSheetNum(sheet):
    rows = sheet.max_row
    IdNum = 0
    for k in range(7, rows + 1):
        Id = sheet.cell(row=k, column=2).value
        if(Id == None):
            return IdNum
        else:
            IdNum = IdNum + 1;
    return IdNum

"""
    计算列表平均数
    Parm:
        num 列表
    Return:
        AvgNum 平均数
"""
def AverageNum(num):
    nsum = 0
    for i in range(len(num)):
        nsum += num[i]
    return nsum / len(num)

"""
    为抗渗等级，膨胀为None值加入/
    Parm:
        cell
    Return：
        None /
        NotNone cell.Value
"""
def IsLevelNone(value):
    if(value == None):
        return '/'
    else:
        return value

"""
    为抗渗等级，膨胀为0值加入/
    Parm:
        cell
    Return：
        None 0
        NotNone cell.Value
"""
def IsSwellingLevelNone(value):
    if(value == 0):
        return '/'
    else:
        return value

"""
    数值转中文大写
"""

def digital_to_chinese(digital):
    str_digital = str(digital)
    chinese = {'1': '壹', '2': '贰', '3': '叁', '4': '肆', '5': '伍', '6': '陆', '7': '柒', '8': '捌', '9': '玖', '0': '零'}
    chinese2 = ['拾', '佰', '仟', '万']
    jiao = ''
    bs = str_digital.split('.')
    yuan = bs[0]
    if len(bs) > 1:
        jiao = bs[1]
    r_yuan = [i for i in reversed(yuan)]
    count = 0
    for i in range(len(yuan)):
        if i == 0:
            continue
        r_yuan[i] += chinese2[count]
        count += 1
        if count == 4:
            count = 0
            chinese2[3] = '亿'
    s_jiao = [i for i in jiao][:2]# 去掉小于厘之后的
    j_count = -1
    for i in range(len(s_jiao)):
        # s_jiao[i] += chinese2[j_count]
        j_count -= 1
    if(len(s_jiao) == 0):
        s_jiao = ['']
    else:
        s_jiao = ['点'] + s_jiao
    last = [i for i in reversed(r_yuan)] + s_jiao

    last_str = ''.join(last)
    for i in range(len(last_str)):
        digital = last_str[i]
        if digital in chinese:
            last_str = last_str.replace(digital, chinese[digital])

    return last_str

def to_currency(number):
    if not isinstance(number, float) and not isinstance(number, int):
        return 'non number'
    if number < 0 or number > 9999999999999.99:
        return 'wrong number'
    if number == 0:
        return '零'
    c_d = {'0': '零', '1': '壹', '2': '贰', '3': '叁', '4': '肆', '5': '伍', '6': '陆', '7': '柒', '8': '捌', '9': '玖'}
    d_d = {0: '', 1: '', 2: '点', 3: '拾', 4: '佰', 5: '仟', 6: '万', 7: '拾', 8: '佰', 9: '仟', 10: '亿', 11: '拾', 12: '佰',
           13: '仟',
           14: '万'}
    L = []
    pre = '0'
    s = str(int(number * 100))[::-1].replace('.', '')
    index = -1

    for c in s:
        index += 1
        if c == '0' and pre == '0':
            if index == 2:
                L.insert(0, '')
        elif c == '0':
            if index == 2:
                L.insert(0, '')
            else:
                L.insert(0, '零')
            pre = c
        else:
            if index == 2:
                if pre == '0':
                    L.insert(0, c_d[c] + "")
                    pre = c
                else:
                    L.insert(0, c_d[c] + "" + d_d[index])
                    pre = c
            else:
                L.insert(0, c_d[c] + "" + d_d[index])
                pre = c
    return ''.join(L)

"""
    字符串换行 部位
"""
def StringInLine(inputString, n):
    str = inputString.strip('\n')
    strLen = len(str)
    addLineNum = int(strLen / n)
    flag = 0
    outputString = ""
    if(addLineNum == 0):
        return inputString
    while(addLineNum >= 0):
        outputString = outputString + inputString[0 + flag * n: (flag + 1) * n] + '\n'
        flag = flag + 1
        addLineNum = addLineNum - 1
    return outputString


if __name__ == '__main__':
    print(to_currency(600.1))

