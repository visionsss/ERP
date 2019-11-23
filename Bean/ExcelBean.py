# -*- coding:utf-8 -*-
"""
@author: Darcy
@time: 2018/9/26 16:18
"""

"""
    2.混凝土试件抗压强度检验报告头类
"""
class ConStrengRepoBean():
    def __init__(self, ProjectName, BuildUnit, RepoTime):
        self.ProjectName = ProjectName
        self.BuildUnit = BuildUnit
        # self.InspectionName = InspectionName
        self.RepoTime = RepoTime

"""
    2.混凝土试件使用记录测试单体类
"""
class ConStrengUseBean():
    def __init__(self, TestName, ProjectSite, ConcreteStrength, CuringDate, MoldingTime, LongStrength, WideStrength,
                 HigthStrength, AvgStrength,CuringNum,SortTime, ImperLevel, SwellLevel, StrengthString):
        self.TestName = TestName
        self.ProjectSite = ProjectSite
        self.ConcreteStrength = ConcreteStrength
        self.CuringDate = CuringDate
        self.MoldingTime = MoldingTime
        self.LongStrength = LongStrength
        self.WideStrength = WideStrength
        self.HigthStrength = HigthStrength
        self.AvgStrength = AvgStrength
        self.CuringNum = CuringNum
        self.SortTime = SortTime
        self.ImperLevel = ImperLevel
        self.SwellLevel = SwellLevel
        self.StrengthString = StrengthString
"""
    3.混凝土强度试验单体类
"""
class ConRepoSumBean():
    def __init__(self, ProjectSite, ConcreteStrength, CuringDate, TypicalStrength, ToStrength, SumSortTime):
        self.ProjectSite = ProjectSite
        self.ConcreteStrength = ConcreteStrength
        self.CuringDate = CuringDate
        self.TypicalStrength = TypicalStrength
        self.ToStrength = ToStrength
        self.SumSortTime = SumSortTime

"""

"""
class ConRepoCompBean():
    def __init__(self, CompStrength, CompAvgStrength):
        self.CompStrength = CompStrength
        self.CompAvvgStrength = CompAvgStrength

"""
    5.混个凝土出厂合格证 工地使用数据类
"""
class FactoryConUseBean():
    def __init__(self, ProjectName, BuildUnit, ConstrUnit, ContractId, ProjectManager, ProjectSite, ConcreteName,
                 ConcreteStrength, ImperLevel, SwellLevel, CuringDate, StrengthString):
        self.ProjectName = ProjectName
        self.BuildUnit = BuildUnit
        self.ConstrUnit = ConstrUnit
        self.ContractId = ContractId
        self.ProjectManager = ProjectManager
        self.ProjectSite = ProjectSite
        self.ConcreteName = ConcreteName
        self.ConcreteStrength = ConcreteStrength
        self.ImperLevel = ImperLevel
        self.SwellLevel = SwellLevel
        self.CuringDate = CuringDate
        self.StrengthString = StrengthString

"""
    7.混凝土配合比设计报 使用记录单体类
"""
class ConMixUseBean():
    def __init__(self, ProjectName, BuildUnit, ConcreteName, ConcreteStrength, ImperLevel, SwellLevel, CuringDate):
        self.ProjectName = ProjectName
        self.BuildUnit = BuildUnit
        self.ConcreteName = ConcreteName
        self.ConcreteStrength = ConcreteStrength
        self.ImperLevel = ImperLevel
        self.SwellLevel = SwellLevel
        self.CuringDate = CuringDate
        # self.SortTime = SortTime

"""
    7.混凝土配合比设计报 水泥使用记录单体 
"""
class ConMixCementAtr():
    def __init__(self, ArrivalTime, CementId, CementName, ConcreteStrength, ImperLevel, SwellLevel, R3_Bending, R28_Bending, R3_Compression, R28_Compression):
        self.ArrivalTime = ArrivalTime
        self.CementId = CementId
        self.CementName = CementName
        self.ConcreteStrength = ConcreteStrength
        self.ImperLevel = ImperLevel
        self.SwellLevel = SwellLevel
        self.R3_Bending = R3_Bending
        self.R28_Bending = R28_Bending
        self.R3_Compression = R3_Compression
        self.R28_Compression = R28_Compression