# coding=utf-8
"""
Author: vision
date: 2019/4/4 19:59
"""

from DateBase.creat import session
from DateBase.creat import Parameter


def get_parm():
    parm = session.query(Parameter).filter()
    return parm[-1]


if __name__ == "__main__":
    parm = get_parm()
    print(parm.Id, end=' ')
    print(parm.MinC_Strength, end=' ')
    print(parm.MaxC_Strength, end=' ')
    print(parm.MinS_FinenessDensity, end=' ')
    print(parm.MaxS_FinenessDensity, end=' ')
    print(parm.MinS_SurfaceDensity, end=' ')
    print(parm.MaxS_SurfaceDensity, end=' ')
    print(parm.MinS_Density, end=' ')
    print(parm.MaxS_Density, end=' ')
    print(parm.MinS_SlitContent, end=' ')
    print(parm.MaxS_SlitContent, end=' ')
    print(parm.MinS_WaterContent, end=' ')
    print(parm.MaxS_WaterContent, end=' ')
    print(parm.MinG_GrainContent, end=' ')
    print(parm.MaxG_GrainContent, end=' ')
    print(parm.MinG_CrushLevel, end=' ')
    print(parm.MaxG_CrushLevel, end=' ')
    print(parm.MinG_Density, end=' ')
    print(parm.MaxG_Density, end=' ')
    print(parm.MinG_SlitContent, end=' ')
    print(parm.MaxG_SlitContent, end=' ')
    print(parm.MinG_WaterContent, end=' ')
    print(parm.MaxG_WaterContent, end=' ')
    print(parm.MinA_Density, end=' ')
    print(parm.MaxA_Density, end=' ')
    print(parm.MinR7_Compression, end=' ')
    print(parm.MaxR7_Compression, end=' ')
    print(parm.MinR28_Compression, end=' ')
    print(parm.MaxR28_Compression, end=' ')
    print(parm.Project1Manager, end=' ')
    print(parm.Project1FillSheeter, end=' ')
    print(parm.Project2InspectCodeEdit, end=' ')
    print(parm.Project2Manager, end=' ')
    print(parm.Project2Checker, end=' ')
    print(parm.Project2Try, end=' ')
    print(parm.Project3MakeSheet, end=' ')
    print(parm.Project3InspectCodeEdit, end=' ')
    print(parm.Project4Manager, end=' ')
    print(parm.Project4Checker, end=' ')
    print(parm.Project4Calculate, end=' ')
    print(parm.Project4InspectCodeEdit, end=' ')
    print(parm.Project5Manager, end=' ')
    print(parm.Project5Filler, end=' ')
    print(parm.Project7ConDesignSpeciEdit, end=' ')
    print(parm.Project7CodeEdit, end=' ')
    print(parm.Project7Manager, end=' ')
    print(parm.Project7Checker, end=' ')
    print(parm.Project7try, end=' ')
    print(parm.Project9Manager, end=' ')
    print(parm.Project9Checker, end=' ')
    print(parm.Project9Record, end=' ')
    print(parm.Project10ConTestReportTestBasisEdit, end=' ')
    print(parm.Project10InspectCodeEdit, end=' ')
    print(parm.Project10TimeEdit, end=' ')
    print(parm.Project10MaxCreepEdit, end=' ')
    print(parm.Project10MinCreepEdit, end=' ')
    print(parm.Project10Manager, end=' ')
    print(parm.Project10Examine, end=' ')
    print(parm.Project10Checker, end=' ')
    print(parm.Project11Manager, end=' ')
    print(parm.Project11Checker, end=' ')
