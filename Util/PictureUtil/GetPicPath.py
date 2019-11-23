import os
def getpicpath(picname):
    s = os.getcwd()
    s = os.path.split(s)[0]
    s = os.path.split(s)[0]
    s = s + "\pic\\" + picname + ".jpg"
    return s
