import os
import csv
import xlrd
from db_mod import *


# 导入初三中考报名所有学生
# 增加届别字段
# 增加班级字段
SIGN_KS = ('signid','name','sex','idcode','sch',
    'schcode','zhtype','graduation_year','classcode')

# 转学代码zhtype说明：
# 1 县外转入，2 县外转入，县内转
# 3 县内转学，4 无转学记录 0 历届


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

    gath_data(SignAll,SIGN_KS,'signall',0) # 末尾行无多余数据
