# SFG 创建
# 测试记录：
# 5-10，SH2,OK
# 5-18,SZ,OK

import pandas as pd
from Field_Tools import *
import sys

# copy值
def copy_value(df11,df1):
    # 字段对应关系
    f_dict = {
                'Material Number':'半成品编号',
                'Material Description (English)':'半成品名称（英文）',
                'Material Description (Chinese)':'半成品名称（中文）',
                'Old Material Number':'被替换旧物料号',
                'Base Unit of Measure':'基本单位',
                'Net Weight':'净重',
                'Minimum Lot Size':'最小批次',
                'Maximum Lot Size':'最大批次',
                'Maximum Stock Level':'最大库存',
                'Rounding Value for Purc Order Qty':'成倍滚动值',
                'Planned Delivery Time in Days':'计划发货时间',
                'Goods Receipt Processing Time in Days':'收货时间',
                'Safety Stock':'安全库存',
                'Minimum Remaining Shelf Life':'最小货架生命周期',
                'Total Shelf Life':'保质期',
                'Material Group':'半成品产品组',
                'MRP Type':'MRP类型',
                'MRP Controller':'MRP 控制者',
                'Lot Size':'批次',
                'Procurement Type':'产出方式',
                'Period Indicator':'周期标识',
                'Fiscal Year Variant':'财务年度',
                'Planning Strategy Group':'计划策略组'
                }    
    for k,v in f_dict.items():
        df1[k] = df11[v]
    print('Copy value Done.')

# 固定值
def fix_value(df1):
    # 字段固定值
    f_dict = {
            'Industry Sector':'S',
            'Cross-Plant Material Status':'03',
            'Weight Unit':'KG',
            'Country Key':'CN',
            'Tax Data - Tax Category':'MWST',
            'General Item Category Group':'NORM',
            'Availability Check':'02',
            'MRP Group':'ZVR2',
            'Plant-Specific Material Status':'08',
            'Backflush':'1',
            'Scheduling Margin Key for Floats':'000',
            'BWD Consumption Period':'000',
            'Period indicator for SLED':'D',
            'Negative Stocks Allowed in Plant':'X',
            'Price Control Indicator':'S',
            'Price Unit':'1',
            'Lot Size for Product Costing':'1000',
            'Price Determination':'3',
            'Material Ledger Indicator':'X',
            'Material Origin':'X',
            'Variance Key':'000001'
                }
    for k,v in f_dict.items():
        df1[k] = v
    
    print('Fixed value Done.')

# 拆分值
def split_value(df1):
    # 需拆分字段列表
    f_list = [
                'Material Group',
                # 'Manufacturing Type',
                'MRP Type',
                'MRP Controller',
                'Lot Size',
                'Procurement Type',
                'Period Indicator',
                'Fiscal Year Variant',
                'Planning Strategy Group'
            ]
    # 逐列拆分，按‘-’取左边的
    for a in f_list:
        df1[a] = df1[a].str.split('-',expand=True)

    print('Split value Done.')

# 取值转换
def convert_value1(df11,df1):
    # 获取Field_Tools工厂对应字典
    plant_dict = PlantDict()

    # 逐行遍历df11，将值给df1
    for a in df11.index:
        # plant,根据NPD工厂字段转换
        pl = df11.at[a,'工厂']
        df1.at[a,'Plant'] = plant_dict[pl]

        # Batch Management，判断
        ba = df11.at[a,'是否进行批次管理']
        if ba == 'Yes':
            df1.at[a,'Batch Management Requirement Ind'] = 'X'

        # Special Procurement Type,判断
        sp = df11.at[a,'是否虚拟半成品']
        if sp == 'Yes':
            df1.at[a,'Special Procurement Type'] = '50'

    print('Convert value 1 Done.')

# 条件判断转换
def convert_value2(df11,df1):
    # BOM Usage，Alternative BOM，With Quantity Structure
    f_dict = {'BOM Usage':'5','Alternative BOM':'1','With Quantity Structure':'X'}
    # Purchasing Group,Purchasing Value Key,数据源拆分值，后面可以直接复制
    df11['采购组'] = df11['采购组'].str.split('-',expand=True)
    df11['采购价值代码'] = df11['采购价值代码'].str.split('-',expand=True)
    
    # 逐行判断df1后，将转换结果给df1
    for a in df1.index:
        mat = df1.at[a,'Material Number']
        pr_type = df1.at[a,'Procurement Type']
        # 根据Procurement Type判断：Material Type，Purchasing Group,Purchasing Value Key
        if pr_type == 'F':
            # Material Type,外购半成品
            if 511000 <= int(mat) <= 519999:
                df1.at[a,'Material Type'] = 'ZCFH'
            else:
                df1.at[a,'Material Type'] = 'WRONG code range'
            # Purchasing Group,Purchasing Value Key,copy值
            df1.at[a,'Purchasing Group'] = df11.at[a,'采购组']
            df1.at[a,'Purchasing Value Key'] = df11.at[a,'采购价值代码']
        else:
            # Material Type，自制半成品
            if 500000 <= int(mat) <= 510999:
                df1.at[a,'Material Type'] = 'HALC'
            elif 511000 <= int(mat) <= 519999:
                df1.at[a,'Material Type'] = 'SFCH'
            elif 550001 <= int(mat) <= 559999:
                df1.at[a,'Material Type'] = 'SHMX'
            elif 560001 <= int(mat) <= 569999:
                df1.at[a,'Material Type'] = 'SZMX'
            else:
                df1.at[a,'Material Type'] = 'WRONG code range'

            # BOM Usage，Alternative BOM，With Quantity Structure
            for k,v in f_dict.items():
                df1.at[a,k] = v
    
    # 利用df11和df1的顺序一致特性
    # Valuation Class
        if df11.at[a,'产出方式'] == 'F-3rd-3rd party procurement':
            df1.at[a,'Valuation Class'] = '310'
        elif df11.at[a,'产出方式'] == 'F-ICY-ICY procurement':
            df1.at[a,'Valuation Class'] = '330'
        else:
            df1.at[a,'Valuation Class'] = '300'
    # Class Type，Class，Country of Origin，根据Material Type判断
        mat_type = df1.at[a,'Material Type']
        if mat_type in ['ZCFH','SHMX','SZMX']:
            df1.at[a,'Class Type'] = '022'
            df1.at[a,'Class'] = 'FIFO_BATCH'
            df1.at[a,'Country of Origin'] = 'CN'
    
    print('Convert value 2 Done.')    

# 多工厂处理
def convert_value3(df1):
    # 遍历已整理的df1,根据工厂判断:Storage Location,Issue Storage Location,Default Storage Location
    for a in df1.index:
        pl = df1.at[a,'Plant']
        # Storage Location,Issue Storage Location,Default Storage Location
        if pl == '0023':
            df1.at[a,'Storage Location'] = df1.at[a,'Issue Storage Location'] = '0015'
            df1.at[a,'Default Storage Location'] = '0001'
        else:
            df1.at[a,'Storage Location'] = df1.at[a,'Issue Storage Location'] = df1.at[a,'Default Storage Location'] = '0006' 
    
    # 多工厂判断处理
    # df1后面会变化，df1初版存一下df00
    df00 = df1.copy(deep=True)
    # 筛选出0078&0079双工厂的行，临时存放筛选结果的存到df01
    df01 = df00.loc[df00['Plant'] == '0078,0079']
    
    # 分二次赋值工厂给双工厂物料行df01，追加到df1中
    for p in ['0078','0079']:
        df01['Plant'] = p
        df1 = df1.append(df01)
    
     # 追加后重置df1的index,才能做遍历删除原来多值的物料
    df1 = df1.reset_index(drop=True)
    # 删除原来是0078,0079的物料行
    for c in df1.index:
        pl = df1.at[c,'Plant']
        if pl == '0078,0079':
            df1 = df1.drop(index=[c])
        
    # 做过删除后，重置df1的index
    df1 = df1.reset_index(drop=True)

    print('Convert value 3 Done')
    return df1

# 生成ACC&CO数据
def acc_value(df1,df2):
    # 可以完全copy df1已整理好的数据
    # copy字段关系
    f1_dict = {
            'Material Code':'Material Number',
            'English Matl Desc':'Material Description (English)',
            'Chinese Matl Desc':'Material Description (Chinese)',
            'Plant':'Plant',
            'Valuation Key':'Plant',
            'Valuation Class':'Valuation Class',
            'With Quantity Structure':'With Quantity Structure',
            'Alternative BOM':'Alternative BOM',
            'BOM Usage':'BOM Usage'
            }
    for k1,v1 in f1_dict.items():
        df2[k1] = df1[v1]
    
    # 固定值字段关系
    f2_dict = {
                'Price Determination':'3',
                'Material Ledger Indicator':'X',
                'Price Control Indicator':'S',
                'Price Unit':'1',
                'Material Origin':'X',
                'Variance Key':'000001',
                'Lot Size for Product Costing':'1000'
            }
    for k2,v2 in f2_dict.items():
        df2[k2] = v2

    print('ACC&CO Done')

# 生成Z_GL_MATERIAL数据
def class_value(df11,df3):
    # copy值
    df3['Material'] = df11['半成品编号']
    df3['Manufacturing Type'] = df11['Manufacturing Type']

    # 固定值
    df3['Class'] = 'Z_GL_MATERIAL'

    # 拆分
    df1 = df3['Manufacturing Type'].str.split('-',expand=True)
    df3['Manufacturing Type'] = df1[0]

    # 删除重复行
    df3.drop_duplicates(inplace=True)
    df3 = df3.reset_index(drop=True)

    print('Class value Done.')
    return df3

# 生成模板
def GenaTemp():
    writer = pd.ExcelWriter('SFG_Upload.xlsx')

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

    # 生成Class=Z_GL_MATERIAL的上传模板
    fgGLFileds = FGHanaFields()
    df3 = pd.DataFrame(columns=fgGLFileds)
    # 后面有对应的writer.save()
    df3.to_excel(writer,'Z_GL_MATERIAL')

    writer.save()
    print('GenaTemp Done.')

# 数据整理
def GetSheet_sfg():
    # 提示做预处理===========================
    text1 = '''
            1. 若为SH双工厂，无需拆分
            如果OK，请输入1继续。。。'''
    print(text1)
    # 不OK就退出
    res = input('请输入：')
    if res != '1':
        print('程序退出。')
        sys.exit()

    # 生成模板=================================
    GenaTemp()

    # 获取数据源================================
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='SFG',index_col=0,dtype=str)

    # 获取模板================================
    # Z1UMM24模板
    df1 = pd.read_excel('SFG_Upload.xlsx',sheet_name='Z1UMM24',index_col=0,dtype=str)
    # ACC&CO模板
    df2 = pd.read_excel('SFG_Upload.xlsx',sheet_name='ACC&CO',index_col=0,dtype=str)
    # 获得Z_GL_MATERIAL模板
    df3 = pd.read_excel('SFG_Upload.xlsx',sheet_name='Z_GL_MATERIAL',index_col=0,dtype=str)

    # 生成Z1UMM24数据================================
    # copy值
    copy_value(df11,df1)
    # 固定值
    fix_value(df1)
    # 拆分值
    split_value(df1)
    # 取值转换
    convert_value1(df11,df1)
    # 条件判断转换
    convert_value2(df11,df1)
    # 多工厂处理
    df1 = convert_value3(df1)
    
    # 生成ACC&CO数据================================
    # 从df1 copy过来
    acc_value(df1,df2)

    # 生成Z_GL_MATERIAL数据======================================
    df3 = class_value(df11,df3)

    # 生成Rrmakd工作表======================================
    remark = {'1':'MM02,勾选Unlimited',
            '2':'维护Class=001,Z_GL_MATERIAL，MT'}
    df44 = pd.Series(remark)

    # 保存生成的模板数据================================
    writer = pd.ExcelWriter('SFG_Upload.xlsx')
    # 保存数据源
    df11.to_excel(writer,'SFG')
    # 保存数据
    df44.to_excel(writer,'Remark')
    df1.to_excel(writer,'Z1UMM24')
    df2.to_excel(writer,'ACC&CO')
    df3.to_excel(writer,'class')
    
    writer.save()
    print('Pls check SFG_Upload.xlsx')

# GenaTemp()
GetSheet_sfg()