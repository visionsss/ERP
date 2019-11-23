"""
@author: Darcy
@time: 2018/10/2 12:16
"""
import Util.ExcelUtil
import Util.FileUitl
from Util.ExcelInitUtil import *
from datetime import timedelta
from DateBase.connect_db import session
"""
    生成工地预拌混凝土证明书
    Parm:
        UseBook 工地混凝土使用记录表
    Return:
        modeBook 生成工作簿
"""
def CreteConUseProve(UseBook):
    modeName = "../../Mode/8.xlsx"
    from Util.get_parm import get_parm
    parm = get_parm()
    modeBook = Util.FileUitl.openFileXlsx(modeName)
    modeBook = Util.ExcelUtil.CopyFromModeExcel(modeBook, UseBook)
    useSheets = UseBook.sheetnames
    for sheetNum in range(len(useSheets)):
        """
            获取使用记录数据
        """
        DataList = []
        sheet = UseBook.worksheets[sheetNum]
        ProjectName = sheet['A2'].value.replace('工程名称：', '').strip()
        BuildUnit = sheet['A3'].value.replace('建设单位：', '').strip()
        ConstrUnit = sheet['A4'].value.replace('施工单位：', '').strip()
        ContractId = sheet['A5'].value.replace('合同编号：', '').strip()
        ProjectManager = sheet['A6'].value.replace('项目经理：', '').strip()
        ProveAdress = sheet['A7'].value.replace('证明书地址：','').strip()
        ProveSite = sheet['A8'].value.replace('证明书工程部位：','').strip()
        # print(ProveSite)
        CementSumNum = 0
        rows = sheet.max_row
        for rowNum in range(10, rows+1):
            CementSumNum = CementSumNum + sheet.cell(row=rowNum, column=8).value
            DataList.append(sheet.cell(row=rowNum, column=6).value +  timedelta(days=28))
        DataList.sort(reverse=False)
        """
            插入数据
        """
        resultSheet = modeBook.worksheets[sheetNum]
        iniexcel8(resultSheet)  # 初始化表格
        resultSheet['A3'] = "合 同 号 ：" + ContractId
        resultSheet['A4'] = "建设单位：" + BuildUnit
        resultSheet['A5'] = "施工单位：" + ConstrUnit
        resultSheet['A6'] = "施工负责人：" + ProjectManager
        resultSheet['A7'] = "位于：" + ProveAdress
        resultSheet['A8'] = "工程名称：" + ProjectName
        resultSheet['A9'] = "工程部位：" + ProveSite
        # 3.28
        if(ProveSite.find("主体") > 0):
            resultSheet['A13'] = "           持 证 范 围：竣工验收"
        CuringNumString = "实际使用我站混凝土为：{CementSumNum}㎥（{CementSumCN}立方米)"
        resultSheet['A10'] = CuringNumString.format(CementSumNum="{0:.1f}".format(CementSumNum),
                                                    CementSumCN=Util.ExcelUtil.to_currency(CementSumNum))
        resultSheet['D21'] = DataList[-1].strftime('%Y$%#m/%#d').replace('$','年').replace( '/', '月') + '日'
        # resultSheet['B15'] = parm[0].ConUseProveUtil
    return modeBook

if __name__ == '__main__':
    UsePath = "..\..\工地混凝土使用记录.xlsx"
    # UsePath = r"C:\Users\62458\Desktop\文件\竣工验收资料表格整理\8、工地使用预拌混凝土证明书\输入\工地混凝土使用记录.xlsx"
    UseBook = Util.FileUitl.openFileXlsx(UsePath)
    session = session
    modeBook = CreteConUseProve(UseBook)
    modeBook.save('工地使用预拌混凝土证明书8.xlsx')
    pass