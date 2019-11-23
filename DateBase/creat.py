# coding=utf-8
"""
Author: vision
date: 2019/4/3 16:00
"""
from sqlalchemy import Column, Date, Integer, String, DateTime, ForeignKey
from sqlalchemy.ext.declarative import declarative_base
from DateBase.connect_db import session, engine
from sqlalchemy.orm import relationship


# 创建table基类
Base = declarative_base()


class CementAttributeDatum(Base):
    __tablename__ = 'cement_attribute_data'
    Id = Column(Integer, primary_key=True)
    ArrivalTime = Column(Date, nullable=False)
    CementVariety = Column(String(20), nullable=False)
    Manufacturer = Column(String(20), nullable=False)
    ProductionDate = Column(Date, nullable=False)
    CementId = Column(String(20), nullable=False)
    CementNumber = Column(Integer)
    IsStability = Column(Integer)
    InitialTime = Column(String(20))
    FinalTime = Column(String(20))
    R3_Compression = Column(String(20))
    R28_Compression = Column(String(20))
    R3_Bending = Column(String(20))
    R28_Bending = Column(String(20))
    PriorityLevel = Column(Integer, nullable=False)


class ConcreteMix(Base):
    __tablename__ = 'concrete_mix'
    ConcreteName = Column(String(20), nullable=False)
    MixRatioID = Column(String(20), primary_key=True, nullable=False)
    StrengthLevel = Column(String(20))
    ImperLevel = Column(String(20))
    SwellLevel = Column(String(20))
    MixRatioName = Column(String(50))
    SlumpNum = Column(String(20), nullable=False)
    StandardDeviation = Column(String(20), nullable=False)
    ConcreteStrengh = Column(String(20), nullable=False)
    WaterNum = Column(String(20), nullable=False)
    CementNum = Column(String(20), nullable=False)
    FlyashNum = Column(String(20), nullable=False)
    SandNum = Column(String(20), nullable=False)
    GravelNum = Column(String(20), nullable=False)
    CementRatio = Column(String(20), nullable=False)
    SandRatio = Column(String(20), nullable=False)
    AdmixtureAmount = Column(String(20), nullable=False)
    AdmixtureNum = Column(String(20), nullable=False)
    SwellingNum = Column(Integer, nullable=False)
    MassDensity = Column(String(20), nullable=False)
    InitialTime = Column(String(50), nullable=False)
    FinalTime = Column(String(50), nullable=False)


class Parameter(Base):
    __tablename__ = 'parameter'
    Id = Column(Integer, primary_key=True)
    MinC_Strength = Column(String(20), nullable=False)
    MaxC_Strength = Column(String(20), nullable=False)
    MinS_FinenessDensity = Column(String(20), nullable=False)
    MaxS_FinenessDensity = Column(String(20), nullable=False)
    MinS_SurfaceDensity = Column(String(20), nullable=False)
    MaxS_SurfaceDensity = Column(String(20), nullable=False)
    MinS_Density = Column(String(20), nullable=False)
    MaxS_Density = Column(String(20), nullable=False)
    MinS_SlitContent = Column(String(20), nullable=False)
    MaxS_SlitContent = Column(String(20), nullable=False)
    MinS_WaterContent = Column(String(20), nullable=False)
    MaxS_WaterContent = Column(String(20), nullable=False)
    MinG_GrainContent = Column(String(20), nullable=False)
    MaxG_GrainContent = Column(String(20), nullable=False)
    MinG_CrushLevel = Column(String(20), nullable=False)
    MaxG_CrushLevel = Column(String(20), nullable=False)
    MinG_Density = Column(String(20), nullable=False)
    MaxG_Density = Column(String(20), nullable=False)
    MinG_SlitContent = Column(String(20), nullable=False)
    MaxG_SlitContent = Column(String(20), nullable=False)
    MinG_WaterContent = Column(String(20), nullable=False)
    MaxG_WaterContent = Column(String(20), nullable=False)
    MinA_Density = Column(String(20), nullable=False)
    MaxA_Density = Column(String(20), nullable=False)
    MinR7_Compression = Column(String(20), nullable=False)
    MaxR7_Compression = Column(String(20), nullable=False)
    MinR28_Compression = Column(String(20), nullable=False)
    MaxR28_Compression = Column(String(20), nullable=False)
    Project1Manager = Column(String(40), nullable=False)  # 表1工程负责人
    Project1FillSheeter = Column(String(40), nullable=False)  # 表1填表人
    Project2InspectCodeEdit = Column(String(40), nullable=False)  # 表2检验码
    Project2Manager = Column(String(40), nullable=False)  # 表2单位检验工程负责人
    Project2Checker = Column(String(40), nullable=False)  # 表2校核
    Project2Try = Column(String(40), nullable=False)  # 表2检验
    Project3MakeSheet = Column(String(40), nullable=False)  # 表3制表
    Project3InspectCodeEdit = Column(String(40), nullable=False)  # 表3检验码
    Project4Manager = Column(String(40), nullable=False)  # 表4单位工程负责人
    Project4Checker = Column(String(40), nullable=False)  # 表4复核
    Project4Calculate = Column(String(40), nullable=False)  # 表4计算
    Project4InspectCodeEdit = Column(String(40), nullable=False)  # 表4检验码
    Project5Manager = Column(String(40), nullable=False)  # 表5技术负责人
    Project5Filler = Column(
        String(40),
        nullable=False)                         # 表5技术填表人
    Project7ConDesignSpeciEdit = Column(String(40), nullable=False)  # 表7实验规程
    Project7CodeEdit = Column(String(40),
                              nullable=False)                       # 表7实验规程：
    Project7Manager = Column(String(40), nullable=False)  # 表7技术负责人
    Project7Checker = Column(String(40), nullable=False)  # 表7校核
    Project7try = Column(String(40), nullable=False)  # 表7试验
    Project9Manager = Column(String(40), nullable=False)  # 表9工程负责人
    Project9Checker = Column(String(40), nullable=False)  # 表9质检员
    Project9Record = Column(String(40), nullable=False)  # 表9记录
    Project10ConTestReportTestBasisEdit = Column(
        String(40), nullable=False)  # 表10检验依据
    Project10InspectCodeEdit = Column(String(40), nullable=False)  # 表10检验码
    Project10TimeEdit = Column(String(40), nullable=False)  # 加水压日期的时分
    Project10MaxCreepEdit = Column(String(20), nullable=False)  # 最高渗水值最大区间
    Project10MinCreepEdit = Column(String(20), nullable=False)  # 最低渗水值最小区间
    Project10Manager = Column(String(40), nullable=False)  # 表10技术负责人
    Project10Examine = Column(String(40), nullable=False)  # 表10审核
    Project10Checker = Column(String(40), nullable=False)  # 表10检验人
    Project11Manager = Column(String(40), nullable=False)  # 表11技术负责人
    Project11Checker = Column(String(40), nullable=False)  # 表11试验


class ConcreteUseRecord(Base):
    __tablename__ = 'concrete_use_record'
    Id = Column(Integer, primary_key=True)
    UseRecordId = Column(
        ForeignKey('concrete_use_record_head.Id'),
        nullable=False,
        index=True)
    ProjectSite = Column(String(100), nullable=False)
    ConcreteName = Column(String(100), nullable=False)
    ConcreteStrength = Column(String(100), nullable=False)
    ImperLevel = Column(String(100))
    SwellLevel = Column(String(100))
    CuringDate = Column(DateTime, nullable=False)
    CuringTime = Column(String(100), nullable=False)
    CuringNum = Column(Integer, nullable=False)
    demo = Column(String(255))

    concrete_use_record_head = relationship(
        'ConcreteUseRecordHead',
        primaryjoin='ConcreteUseRecord.UseRecordId == ConcreteUseRecordHead.Id',
        backref='concrete_use_records')


class ConcreteUseRecordHead(Base):
    __tablename__ = 'concrete_use_record_head'
    Id = Column(Integer, primary_key=True)
    ProjectName = Column(String(80), nullable=False)
    BuildUnit = Column(String(50), nullable=False)
    ConstrUnit = Column(String(50), nullable=False)
    ContractId = Column(String(50), nullable=False)
    ProjectManager = Column(String(20), nullable=False)
    demo = Column(String(255))


def creat_table():
    Base.metadata.create_all(engine)
    session.commit()


if __name__ == '__main__':
    creat_table()
