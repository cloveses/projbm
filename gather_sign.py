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
            if 'graduation_year' in datas:
                datas['graduation_year'] = str(int(datas['graduation_year']))
            tab_obj(**datas)

def check_idcode(stud):
    '''身份证号验证程序'''
    idcode = stud.idcode
    checkcodes = ['1', '0', 'X','9', '8', '7', '6', '5', '4', '3', '2']
    wi = [7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2]
    s = sum((i*int(j) for i,j in zip(wi,idcode[:-1])))
    checkcode = checkcodes[s % 11]
    ck = idcode[-1].upper()
    if checkcode != ck:
        return '校验出错！'
    if int(idcode[-2]) % 2 == 0 and stud.sex == '男':
        return '性别码出错！'
    if int(idcode[-2]) % 2 == 1 and stud.sex == '女':
        return '性别码出错！'

@db_session
def check_stud_idcode():
    print('身份证号码校验错误信息：')
    for s in select(s for s in SignAll):
        if s.idcode and s.idcode[:-1].isdigit():
            ret = check_idcode(s)
            if ret:
                print(ret,':',s.sch,s.name,s.idcode,s.sex)

if __name__ == '__main__':
    db.bind(**DB_PARAMS)
    db.generate_mapping(create_tables=True)

    gath_data(SignAll,SIGN_KS,'signall',0) # 末尾行无多余数据
    check_stud_idcode()
