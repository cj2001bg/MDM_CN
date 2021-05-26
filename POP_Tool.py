# POP 整理

import pandas as pd
from Field_Tools import *
import sys

# copy值
def copy_value(df11,df1):
    # 字段关系
    f_list = [
                'Material Number',
                'Material Description (English)',
                'Material Description (Chinese)',
                'Material Group',
                'Base Unit of Measure',
                'Purchasing Group',
                'Product Hierarchy',
                'Material Group: Packaging Materials'
            ]
    for f in f_list:
        df1[f] = df11[f]

    print('Copy value Done.')

# 固定值
def fix_value(df1):
    # 字段关系
    f_dict = {
            'Industry Sector':'S',
            'Material Type':'POCH',
            'Plant':'0078',
            'Storage Location':'0007',
            'Sales Organization':'0078',
            'Distribution Channel':'01',
            'Gross Weight':'1',
            'Weight Unit':'KG',
            'Net Weight':'1',
            'Volume':'1',
            'Volume Unit':'DM3',
            'Size/Dimensions':'1*1*1',
            'Material Group: Packaging Materials':'314',
            'Country of Origin':'CN',
            'Country Key':'CN',
            'Tax Data - Tax Category':'MWST',
            'Tax Data - Tax Class 1':'1',
            'Cash Discount Indicator':'X',
            'Material Statistics Group':'1',
            'Account Assignment Group':'82',
            'General Item Category Group':'NORM',
            'Item Category Group':'ZPOP',
            'Material Group 2':'POP',
            'Material Group 3':'POP',
            'Material Group 4':'POP',
            'Material Group 5':'POP',
            'Availability Check':'02',
            'Delivering Plant':'0078',
            'Material Pricing Group':'01',
            'Purchasing Value Key':'3',
            'MRP Type':'ND',
            'Procurement Type':'F',
            'Issue Storage Location':'0007',
            'Default Storage Location':'0007',
            'Period Indicator':'W',
            'BWD Consumption Period':'000',
            'Period indicator for SLED':'D',
            'Price Control Indicator':'S',
            'Price Determination':'3'
                }
    for k,v in f_dict.items():
        df1[k]= v

    print('Fixed value Done.')

# 拆分值
def split_value(df1):
    # 字段列表
    f_list = [
                'Base Unit of Measure',
                'Material Group',
                'Product Hierarchy'
            ]
    # 拆分结果放在临时df0中，取左边值赋值
    for f in f_list:
        df0 = df1[f].str.split('-',expand=True)
        df1[f] = df0[0]

    print('Split value Done.')

# RDC多行处理
def rdc_value(df1):
    # copy df1 到临时表df0,作为数据源
    df0 = df1.copy(deep=True)
    # RDC plant列表
    rdc_list = ['7805','7810','7820','7830','7835','7845','7855','7865','7895']
    
    for r in rdc_list:
        # 每次重新 copy df0的数据源到临时表df01
        df01 = df0.copy()
        # 需要根据RDC plant赋值值的字段列表
        f_list = ['Plant',
                'Storage Location',
                'Issue Storage Location',
                'Default Storage Location']
        for f in f_list:
            df01[f] = r
        # 原表中'Distribution Channel'已经是‘01’了，RDC中就用‘04’了
        df01['Distribution Channel'] = '04'
        df1 = df1.append(df01)
    
    print('RDC value Done.')
    return df1

# 生成上传模板
def GenaTemp():
    writer = pd.ExcelWriter('POP_Upload.xlsx')

    # # 生成Z1UMM24上传模板
    # 获取Z1UMM24所有字段名
    f1 = FieldSAPALL()
    # 生成带标题的Z1UMM24的上传模板
    df1 = pd.DataFrame(columns=f1)
    # 后面有对应的writer.save()
    df1.to_excel(writer,'Z1UMM24')

    writer.save()
    print('GenaTemp Done')

# 数据整理
def GetSheet_pop():
    # 提示做预处理===========================
    text = '''预处理：
                    1. 分配code
                    如果OK，请输入1继续。。。'''
    print(text)
    # 不OK就退出
    res = input('请输入：')
    if res != '1':
        print('程序退出。')
        sys.exit()

    # 生成模板=================================
    GenaTemp()
    # 获取数据源================================
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='POP',index_col=0,dtype=str)

    # 获取模板================================
    # Z1UMM24模板
    df1 = pd.read_excel('POP_Upload.xlsx',sheet_name='Z1UMM24',index_col=0,dtype=str)

    # 生成Z1UMM24数据================================
    # copy值
    copy_value(df11,df1)
    # 固定值
    fix_value(df1)
    # 拆分值
    split_value(df1)
    # RDC多行处理
    df1 = rdc_value(df1)

    # 生成Rrmakd工作表======================================
    remark = {'0':'NA'}
    df44 = pd.Series(remark)

    # 保存生成的模板数据================================
    # 保存数据源
    writer = pd.ExcelWriter('POP_Upload.xlsx')
    # 保存数据
    df11.to_excel(writer,'POP')
    df44.to_excel(writer,'Remark')
    df1.to_excel(writer,'Z1UMM24')

    writer.save()
    print('Please check POP_Upload.xls')

# GenaTemp()
GetSheet_pop()