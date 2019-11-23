# -*- coding:utf-8 -*-
"""
@author: Darcy
@time: 2018/10/16 16:29
"""
import Util.FileUitl
import Util.ExcelUtil
"""
    生成工地预拌混凝土证明书空白模板
    Parm:
        UsePath 工地混凝土使用记录路径
    Return:
        modeBook 生成工作簿
"""
def CreteConUseProveMode(UseBook):
    modeName = "../../Mode/8.xlsx"
    modeBook = Util.FileUitl.openFileXlsx(modeName)
    modeBook = Util.ExcelUtil.CopyFromModeExcel(modeBook, UseBook)
    return modeBook

if __name__ == '__main__':
    UsePath = "..\..\工地混凝土使用记录.xlsx"
    UseBook = Util.FileUitl.openFileXlsx(UsePath)
    modeBook = CreteConUseProveMode(UseBook)
    modeBook.save('工地预拌混凝土证明书空白模板8.xlsx')
    modeBook.save('工地购进使用记录空白模板8.xlsx')
    pass