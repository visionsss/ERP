# coding=utf-8
"""
Author: vision
date: 2019/4/15 10:35
"""

# 精度转换


def get_float(num, count):
    num = float(num)
    if num == 0:
        return 0
    num = num * (10**(count + 1))
    ys = num % 10
    if ys < 5:
        num = num - ys
    else:
        num = num + (10 - ys)
    num = num / (10 ** (count + 1))
    if count <= 0:
        return int(num)
    return num


if __name__ == "__main__":
    n = get_float(12.3, 1)
    print(n)
    n = get_float(37.624, 2)
    print(n)
    n = get_float(37.123456789, 5)
    print(n)
    n = get_float(37.123456789, 0)
    print(n)
    n = get_float(37.123456789, -1)
    print(n)
