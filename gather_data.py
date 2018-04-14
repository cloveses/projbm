import os
import csv
import xlrd
from db_mod import *


# 导入初三在校生、转学表、关键信息变更表（excel格式）到数据库中
# 三类表分别存放子目录：chg、gradey18、keyinfo之中

ZH_KS = ('gsrid','dsrid','idcode','name','sex','birth',
    'sch','zhtype','optdate','zhsrc','zhdes')
GRADE_KS = ('sch','grade','sclass','gsrid','ssrid',
    'dsrid','name','idcode','sex')
kEYINFO_KS = ('ssrid','oname','name','osex','sex','obirth',
    'birth','oidcode','idcode','sch','grade','sclass')

def get_files(directory):
    files = []
    files = os.listdir(directory)
    files = [f for f in files if f.endswith('.xls') or f.endswith('.xlsx')]
    files = [os.path.join(directory,f) for f in files]
    return files

@db_session
def gath_data(tab_obj,ks,chg_dir,grid_end=1,start_row=1):
    """start_row＝1 有一行标题行；gred_end=1 末尾行不导入"""
    files = get_files(chg_dir)
    for file in files:
        wb = xlrd.open_workbook(file)
        ws = wb.sheets()[0]
        nrows = ws.nrows
        for i in range(start_row,nrows-grid_end):
            datas = ws.row_values(i)
            datas = {k:v for k,v in zip(ks,datas) if v}
            tab_obj(**datas)

if __name__ == '__main__':
    db.bind(**DB_PARAMS)
    db.generate_mapping(create_tables=True)

    gath_data(StudZhAll,ZH_KS,'chg')
    gath_data(GradeY18,GRADE_KS,'gradey18',0) # 末尾行无多余数据
    gath_data(KeyInfoChg,kEYINFO_KS,'keyinfo')