import os


def getpiname():
    PicName = []
    s = os.getcwd()
    s = os.path.split(s)[0]
    s = os.path.split(s)[0]
    s = s + r'\pic'
    for i in os.listdir(s):
        PicName.append(i.replace('.jpg', ''))
    return PicName
