from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from Util.PictureUtil.GetPicPath import getpicpath
'''
insertpic（）#插入图片
ws      是一个sheet表
picname 是参数表得到的内容（是人的名字）
width   是图片的宽
heigh   是图片的高
position是图片的位置

return ws 返回一个sheet表
'''


def insertpic(ws, picname='a留空', position='C22', width=70, heigh=25):
    width = 85
    heigh = 30
    path = getpicpath(picname)
    img = Image(path)
    img.width = width
    img.height = heigh
    ws.add_image(img, position)
    return ws


if __name__ == '__main__':
    wb = load_workbook('../Mode/1.xlsx')
    ws = wb.worksheets[0]
    import Dao.SQLUtil
    import Util.ParmUtil
    session, engine = Dao.SQLUtil.connectSQL()
    parm = Util.ParmUtil.GetParm(session)[0]
    name = parm.Project1Manager
    # 函数调用
    ws = insertpic(ws, name, 60, 30, 'A22')
    wb.save('test.xlsx')
