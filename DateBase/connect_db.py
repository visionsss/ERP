# coding=utf-8
"""
Author: vision
date: 2019/4/3 19:00
"""
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
# db_path = '../DateBase/work_space.db'
db_path = '../../DateBase/work_space.db'
engine = create_engine('sqlite:///'+db_path, echo=False)
DBsession = sessionmaker(bind=engine)
session = DBsession()



if __name__ == "__main__":
    pass
