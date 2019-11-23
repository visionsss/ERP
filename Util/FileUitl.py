# -*- coding:utf-8 -*-
"""
@author: Darcy
@time: 2018/9/13 18:34
"""
import xlrd
import string
from datetime import datetime
from xlrd import xldate_as_tuple
import openpyxl
from sqlalchemy.sql.expression import ClauseElement
import win32com.client as win32

"""
    判断文件名是否包含xls
    Parm:
        filename 文件路径
    Return:
        True 包含xls
        False 不包含xls
"""
def FileIsXlsx(filename):
    try:
        xls = '.xlsx'
        result = xls in filename
        return result
    except(ValueError):
        return False

"""
    打开文件Xlsx
    Parm:
        filename 文件路径
    Return:
        workbook 工作簿
"""
def openFileXlsx(filename):
    workbook = openpyxl.load_workbook(filename, data_only=True)
    return workbook

"""
    打开文件Xls
    Parm:
        filename 文件路径
    Return:
        workbook 工作簿
"""
def openFileXls(filename):
    workbook = xlrd.open_workbook(filename, formatting_info=1)
    return workbook

"""
    处理表格时间类数据方法 y/m/d
    Parm:
        cellValue 单元格时间数据 (float)
    Return:
        返回格式 year month day
        
"""
def CellToDatatimeYDM(cellValue):
    date = datetime(*xldate_as_tuple(cellValue, 0))
    cellValue = date.strftime('%Y/%m/%d %H:%M')
    return cellValue

def CellToDatatimeYDMXlsx(cellValue):
    if(type(cellValue) == datetime):
        return cellValue
    else:
        return CellToDatatimeYDM(cellValue)

"""
    处理表格时间类数据方法 H/M
    Parm:
        cellValue 单元格时间数据 (float)
    Return:
        返回格式 hour minutes

"""
def CellToDatatimeHM(cellValue):
    date = datetime(*xldate_as_tuple(cellValue, 0))
    cellValue = date.strftime('%H:%M')
    return cellValue

"""
    处理水泥资料安定性
    Parm:
        cellValue 安定性 
    Return:
        0 不合格
        1 合格
"""
def CemAtrIsStablility(cellValue):
    notStablility = '不合格'
    result = notStablility in cellValue
    if(result == True):
        return 0
    else:
        return 1

"""
    处理混凝土抗渗等级
    Parm:
        cellValue 抗渗等级 
    Return:
        False None
        True 返回cellValue
"""
def ConImperLevel(cellValue):
    NoneStr = '/'
    result = NoneStr in cellValue
    if(result == True):
        return None
    else:
        return cellValue

def ConSweelLevel(cellValue):
    if(type(cellValue) == float):
        return cellValue;
    else:
        return None

"""
    xlrd获得单位格背景颜色
    Parm:
        book 工作簿
        sheet 工作表
        row 列
        col 行
    Return:
        color
"""
def CemAtriIsPriority(book, sheet, row, col):
    xfs = sheet.cell_xf_index(row, col)
    xf = book.xf_list[xfs]
    bgx = xf.background.pattern_colour_index
    if (bgx == 64):
        return 0
    else:
        return 1

def CemAtriIsPriorityXlsx(cell):
    if(cell.fill.start_color.index == "00000000"):
        return 0
    else:
        return 1
"""
    判断获取单元格是否为空值
    Parm:
        cell 单元格
    Return:
        空值 None
        不为空 cell.value
"""
def IsDateNone(cell):
    if(cell.ctype == 6):
        return None
    else:
        return cell.value

def IsDateTimeNone(cell, value):
    if(cell.ctype == 6):
        return None
    else:
        value = value + cell.value
        return CellToDatatimeYDM(value)
"""
    判断获取单元格datatime类型是否为空
    Parm:
        cell 单元格
        timeFlag 标记时间
    Return:
        None 空值
        CellToDatatimeYDM 转化为datatime类型
"""
def IsDataHMTimeNone(cell, timeFlag):
    if(cell.ctype == 6):
        return None
    else:
        CellToDatatimeYDM(cell.value + timeFlag)

"""
判断获取单元格datatime类型是否为空
    Parm:
        cell 单元格
    Return:
        None 空值
        CellToDatatimeYDM 转化为datatime类型
"""
def IsDataYMDTimeNone(cell):
    if(cell.ctype == 6):
        return None
    else:
        CellToDatatimeYDM(cell.value)

"""
    数据库判断是否重复数据
    Parm:
        session 数据库连接池
        model 数据库对象
        **kwargs 判断语句
    Return:
        instance sql语句
        False 重复
        True 不重复
"""
def get_or_create(session, model, defaults=None, **kwargs):
    instance = session.query(model).filter_by(**kwargs).first()
    if instance:
        return instance, False
    else:
        params = dict((k, v) for k, v in kwargs.items() if not isinstance(v, ClauseElement))
        params.update(defaults or {})
        instance = model(**params)
        session.add(instance)
        return instance, True




