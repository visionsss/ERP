# coding=utf-8
"""
Author: vision
date: 2019/4/8 10:53
"""
from os import mkdir, getcwd
from os.path import split
from time import localtime, strftime


def his():
    add_time = strftime("%Y-%m-%d  %H %M %S", localtime())
    mkdir('../../history/' + str(add_time))
    s = getcwd()
    s = split(s)[0]
    s = split(s)[0]
    s = s + '\\history\\'
    # print(s + str(add_time))
    return s + str(add_time)


if __name__ == "__main__":
    his()
