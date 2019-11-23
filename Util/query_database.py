# coding=utf-8
"""
Author: vision
date: 2019/4/7 14:44
"""
from DateBase.connect_db import session
from DateBase.creat import ConcreteMix, CementAttributeDatum

# query_mix 用法
# mix1 = query_mix('普通砼', '15', None, None)
# mix1 = query_mix('泵送砼', '30', 'P6', '0.08')


def query_mix(ConcreteName, StrengthLevel, ImperLevel, SwellLevel):
    mix = session.query(ConcreteMix).filter(
        ConcreteMix.ConcreteName == ConcreteName,
        ConcreteMix.StrengthLevel == StrengthLevel,
        ConcreteMix.ImperLevel == ImperLevel,
        ConcreteMix.SwellLevel == SwellLevel,
    )
    return mix[0]


if __name__ == "__main__":
    mix1 = query_mix('普通砼', '15', None, None)
    mix1 = query_mix('泵送砼', '30', 'P6', None)
    print(mix1.MixRatioID)
