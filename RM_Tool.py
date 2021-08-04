# NPD 整理
# 5-20,SZ,OK
# 7-15,SH,OK
# 20210715,正式使用

import pandas as pd
from Field_Tools import *
import sys

# copy值
def copy_value(df11,df1):
    # 字段关系
    f_dict = {
            'Material Number':'原材料编号',
            'Material Description (English)':'原材料名称（英文）',
            'Material Description (Chinese)':'原材料名称（中文）',
            'Material Group':'物料类型',
            'Old Material Number':'替换哪个原有材料',
            'Basic Material':'物料类型',
            'Country of Origin':'原产地',
            'Purchasing Group':'采购组',
            'Purchasing Value Key':'采购价值代码',
            'MRP Controller':'MRP控制者',
            'Minimum Lot Size':'最小批量',
            'Maximum Lot Size':'最大批次',
            'Maximum Stock Level':'最大库存',
            'Rounding Value for Purc Order Qty':'成倍滚动量',
            'Planned Delivery Time in Days':'计划到货时间(days)',
            'Goods Receipt Processing Time in Days':'收货处理时间（天）',
            'Safety Stock':'安全库存',
            'Minimum Remaining Shelf Life':'最小保质期（天）',
            'Standard Price':'标准采购定单含税价格',
            'Total Shelf Life':'保质期(days)'
                }
    for k,v in f_dict.items():
        df1[k] = df11[v]
    
    print('Copy value Done.')

# 固定值
def fix_value(df1):
    # 字段关系
    f_dict = {
            'Industry Sector':'S',
            'Base Unit of Measure':'KG',
            'Weight Unit':'KG',
            'Material Group: Packaging Materials':'000',
            'Class Type':'022',
            'Class':'FIFO_BATCH',
            'General Item Category Group':'NORM',
            'Availability Check':'02',
            'Batch Management Requirement Ind':'X',
            'Source List Requirement':'X',
            'MRP Group':'ZVR4',
            'MRP Type':'PD',
            'Lot Size':'ZP',
            'Procurement Type':'F',
            'Backflush':'1',
            'Default Storage Location':'0002',
            'Scheduling Margin Key for Floats':'000',
            'Period Indicator':'P',
            'Fiscal Year Variant':'Z1',
            'Planning Strategy Group':'ZX',
            'Period indicator for SLED':'D',
            'Negative Stocks Allowed in Plant':'X',
            'Price Control Indicator':'S',
            'Price Unit':'1',
            'Lot Size for Product Costing':'1',
            'Price Determination':'3',
            'Material Ledger Indicator':'X',
            'Material Origin':'X',
            'Country Key':'CN',
            'Tax Data - Tax Category':'MWST',
            'Variance Key':'000001'
                }
    for k,v in f_dict.items():
        df1[k] = v

    print('Fixed value Done.')

# 拆分值
def split_value(df1):
    # 字段列表
    f_list = [
            'Material Group',
            'Country of Origin',
            'Purchasing Group',
            'Purchasing Value Key',
            'MRP Controller'
            ]
    # 拆分后放在临时df0中，取'-'左边值，重新赋值
    for f in f_list:
        df00 = df1[f].str.split('-',expand=True)
        df1[f] = df00[0]
    
    # Basic Material,临时放在df01拆分后取'-'右边值
    df01 = df1['Basic Material'].str.split('-',n=1,expand=True)
    df1['Basic Material'] = df01[1]

    # Minimum Lot Size,Rounding Value for Purc Order Qty转整数,临时放在df02,取‘.’左边值
    f_list2 = ['Minimum Lot Size','Rounding Value for Purc Order Qty']
    for f2 in f_list2:
        df02 = df1[f2].str.split('.',expand=True)
        df1[f2] = df02[0]

    print('Split value Done.')

# 转换
def conver_value(df11,df1):
    # 获取工厂对应字典
    plant_dict = PlantDict()
    # 根据df11判断，赋值给df1
    for i in df11.index:
        # Plant, 根据NDP工厂字段判断
        pl = df11.at[i,'工厂']
        df1.at[i,'Plant'] = plant_dict[pl]
        # Issue Storage Location，Storage Location（之一）,根据工厂
        if df1.at[i,'Plant'] == '0023':
            df1.at[i,'Issue Storage Location'] = df1.at[i,'Storage Location'] = '0015'
        else:
            df1.at[i,'Issue Storage Location'] = df1.at[i,'Storage Location'] = '0006'
        
        # Material Type,根据Material Number判断
        mat = df1.at[i,'Material Number']
        if 200000 <= int(mat) <= 209999:
            df1.at[i,'Material Type'] = 'RMSZ'
        elif 210000 <= int(mat) <= 239999:
            df1.at[i,'Material Type'] ='ROHC'
        elif 240000 <= int(mat) <= 249999:
            df1.at[i,'Material Type'] = 'RMCH'
        else:
            df1.at[i,'Material Type'] = 'WRONG number range'
    
        # In-house Production Time,根据Country of Origin
        if df1.at[i,'Country of Origin'] != 'CN':
            df1.at[i,'In-house Production Time'] = '1'
        
        # 根据Vendor Type判断Valuation Class
        vt = df11.at[i,'Vendor Type']
        if vt == '3rd party':
            df1.at[i,'Valuation Class'] = '110'
        elif vt == 'ICY':
            df1.at[i,'Valuation Class'] = '120'
        else:
            df1.at[i,'Valuation Class'] = '111'

        # Standard Price保留数字,列操作
        df1['Standard Price'] = df1['Standard Price'].str.replace('RMB/KG','')

    print('Convert value Done.')

# 多库位处理
def copy_loc(df1):
    # copy数据源df1到临时df00
    df00 = df1.copy(deep=True)
    # Storage Location(之二)，将临时结果存储到df01
    # 每个库位调整后，将df01追加到df1
    for a in ['0002','0021','0099']:
        df01 = df00.copy()
        df01['Storage Location'] = a
        df1 = df1.append(df01)

    df1 = df1.reset_index(drop=True)

    print('Location value Done.')
    return df1

# ACC copy值
def copy_value2(df1,df2):
    # 字段关系
    f_dict = {
            'Material Code':'Material Number',
            'English Matl Desc':'Material Description (English)',
            'Chinese Matl Desc':'Material Description (Chinese)',
            'Plant':'Plant',
            'Valuation Key':'Plant',
            'Valuation Class':'Valuation Class',
            'Price Unit':'Price Unit',
            'Standard Price':'Standard Price',
            'Lot Size for Product Costing':'Lot Size for Product Costing'
            }
    for k,v in f_dict.items():
        df2[k] = df1[v]

    print('ACC copy value Done.')

# ACC 固定值
def fix_value2(df2):
    # 对应关系
    f_dict = {
                'Price Determination':'3',
                'Material Ledger Indicator':'X',
                'Price Control Indicator':'S',
                'Material Origin':'X',
                'Variance Key':'000001'
            }
    for k,v in f_dict.items():
        df2[k] = v

    # 删除重复项
    df2.drop_duplicates(inplace=True)
    df2 = df2.reset_index(drop=True)

    print('ACC fixed value Done.')
    return df2

# 生成模板
def GenaTemp():
    writer = pd.ExcelWriter('RM_Upload.xlsx')
    # 生成Z1UMM24上传模板
    # 获取Z1UMM24所有字段名
    f1 = FieldSAPALL()
    # 生成带标题的Z1UMM24的上传模板
    df1 = pd.DataFrame(columns=f1)
    # 后面有对应的writer.save()
    df1.to_excel(writer,'Z1UMM24')

    # 生成ACC&CO上传模板
    # 获取Z1UMM18的ACC&CO所有字段名
    f2 = FieldACCList()
    df2 = pd.DataFrame(columns=f2)
    # 后面有对应的writer.save()
    df2.to_excel(writer,'ACC&CO')
    
    writer.save()
    print('GenaTemp Done')

# 数据整理
def GetSheet_rm():
    # 提示做预处理===========================
    text = '''预处理：
                1. NPD描述是否需调整
                2. 是否有双价格，保证‘标准采购定单含税价格’里是高的那个
                3. ICY/3RD,110- 3rd party，120- ICY，111-cocoa slurry
                4. 多工厂创建请手工拆分工厂成多条，没做功能
                如果OK，请输入1234继续。。。'''
    print(text)
    # 不OK就退出
    res = input('请输入：')
    if res != '1234':
        print('程序退出。')
        sys.exit()
    
    # 生成模板=================================
    GenaTemp()

    # 获取数据源================================
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='RM',index_col=0,dtype=str)

    # 获取模板================================
    # Z1UMM24模板
    df1 = pd.read_excel('RM_Upload.xlsx',sheet_name='Z1UMM24',index_col=0,dtype=str)
    # ACC&CO模板
    df2 = pd.read_excel('RM_Upload.xlsx',sheet_name='ACC&CO',index_col=0,dtype=str)
    
    # 生成Z1UMM24数据================================
    # copy值
    copy_value(df11,df1)
    # 固定值
    fix_value(df1)
    # 拆分值
    split_value(df1)
    # 转换
    conver_value(df11,df1)
    # 多库位处理
    df1 = copy_loc(df1)

    # 生成ACC&CO数据================================
    # ACC copy值
    copy_value2(df1,df2)
    # ACC 固定值
    df2 = fix_value2(df2)

    # 生成Rrmakd工作表======================================
    remark = {
            '1':'手工拆分成多工厂'
            }
    df44 = pd.Series(remark)

    # 保存生成的模板数据================================
    writer = pd.ExcelWriter('RM_Upload.xlsx')

    # 保存数据源
    df11.to_excel(writer,'RM')
    # 保存数据
    df44.to_excel(writer,'Remark')
    df1.to_excel(writer,'Z1UMM24')
    df2.to_excel(writer,'ACC&CO')

    writer.save()
    print('Pls check RM_Upload.xlsx')

# GenaTemp()
# GetSheet_rm()