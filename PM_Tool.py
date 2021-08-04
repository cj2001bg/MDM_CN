# NPD PM整理
# 7-12,新增Gross Weight,Net Weitht,Spec code, Recyclable Plastic Material整理
# 5-18,SZ,OK     5-20,SH1,SH2,OK        7-13,新增字段测试OK
# 7-13，正式使用

import pandas as pd
from Field_Tools import *
import sys

# Z1UMM24,copy值
def copy_value(df11,df1):
    # 字段对应关系
    f_dict = {
                'Material Number':'PM Code',
                'Material Description (English)':'包材描述（英文）',
                'Material Description (Chinese)':'包材描述（中文）',
                'Base Unit of Measure':'基本单位',
                'Material Group':'物料类型',
                'Old Material Number':'替换哪个原有材料',
                'Material Group: Packaging Materials':'包材物料组',
                'Gross Weight':'Packaging Material Total Weight(g)',
                'Net Weight':'Plastic Material Weight(g)',
                'Basic Material':'物料类型',
                'Country of Origin':'原产地',
                'Purchasing Group':'采购组',
                'Purchasing Value Key':'采购价值代码',
                'MRP Controller':'MRP控制者',
                'Lot Size':'批量类型',
                'Minimum Lot Size':'最小批量',
                'Maximum Stock Level':'最大库存',
                'Rounding Value for Purc Order Qty':'Rounding Volume',
                'Planned Delivery Time in Days':'计划到货时间',
                'Goods Receipt Processing Time in Days':'收货处理时间1',
                'Safety Stock':'安全库存',
                'Component Scrap (%)':'Component Scrap (%)',
                'Minimum Remaining Shelf Life':'最短保质期',
                'Total Shelf Life':'保质期',
                'Valuation Class':'Vendor type',
                'Price Unit':'标准采购定单不含税价格',
                'Standard Price':'标准采购定单不含税价格',
                'Future Price':'future price',
                'Future Price - Valid From (YYYYMMDD)':'valid from date'
                }
    for k,v in f_dict.items():
        df1[k] = df11[v]

    print('Copy value Done.')

# Z1UMM24,固定值
def fix_value(df1):
    # 字段关系
    f_dict = {
                   'Industry Sector':'S',
                    'Weight Unit':'G',
                    'Class Type':'022',
                    'Class':'FIFO_BATCH',
                    'General Item Category Group':'NORM',
                    'Availability Check':'02',
                    'Batch Management Requirement Ind':'X',
                    'Source List Requirement':'X',
                    'MRP Group':'ZVR4',
                    'Plant-Specific Material Status':'06',
                    'MRP Type':'PD',
                    'Procurement Type':'F',
                    'Backflush':'1',
                    'Scheduling Margin Key for Floats':'000',
                    'Period Indicator':'P',
                    'Fiscal Year Variant':'Z1',
                    'Planning Strategy Group':'10',
                    'Under Tolerance':'0',
                    'Over Tolerance':'0',
                    'Period indicator for SLED':'D',
                    'Negative Stocks Allowed in Plant':'X',
                    'Price Control Indicator':'S',
                    'Price Determination':'3',
                    'Material Ledger Indicator':'X',
                    'Material Origin':'X',
                    'Variance Key':'000001'
                    }
    for k,v in f_dict.items():
        df1[k] = v

    print('Fix Value Done.')

# Z1UMM24,拆分
def split_value(df11,df1):
    # '-'，需拆分字段列表,取左边值
    f_list = [
            'Base Unit of Measure',
            'Material Group',
            'Material Group: Packaging Materials',
            'Country of Origin',
            'Purchasing Group',
            'Purchasing Value Key',
            'MRP Controller',
            'Lot Size'
            ]
    # 按‘-’拆分，重新赋值
    for a in f_list:
        # 将拆分的值临时放在df0中，将左边的值赋值到列
        df0 = df1[a].str.split('-',expand=True)
        df1[a] = df0[0]
    
    # Basic Material，'-'，需拆分,取右边值
    df00 = df1['Basic Material'].str.split('-',n=1,expand=True)
    df1['Basic Material'] = df00[1]

    # '/'需拆分字段列表
    df01 = df11['标准采购定单含税价格'].str.split('/',expand=True)
    # print(df01)
    # Standard Price,取左边值
    df1['Standard Price'] = df01[0]
    # Price Unit,取右边值
    df1['Price Unit'] = df01[1]
    # Price Unit,检查为1的，赋值为1
    for i in df1.index:
        if df1.at[i,'Price Unit'] == None:
            df1.at[i,'Price Unit'] = 1  
    
    # Lot Size for Product Costing
    df1['Lot Size for Product Costing'] = df1['Price Unit']

    print('Split Value Done.')

# Z1UMM24,转换数据
def convert_value1(df11,df1):
    # 根据数据源判断df11
    # Order Unit,Unit of Issue
    for i in df11.index:
        bu = df11.at[i,'基本单位']
        ou = df11.at[i,'采购单位']
        # 当2个单位不一样时，为Order Unit赋值
        if bu != ou:
            df1.at[i,'Order Unit'] = ou
        
        # Plant
        # 调用工厂对应关系
        pl_dict = PlantDict()
        pl = df11.at[i,'工厂']
        df1.at[i,'Plant'] = pl_dict[pl]

    # Order Unit,Unit of Issue,拆分df1[Order Unit],取左边值
    df0 = df1['Order Unit'].str.split('-',expand=True)
    df1['Order Unit'] = df1['Unit of Issue'] = df0[0]

    # 根据已整理的表判断df1
    for j in df1.index:
        # Var. Oun,根据Order Unit判断
        if df1.at[j,'Order Unit'] == 'ROL':
            df1.at[j,'Var. Oun'] = 1

        # Material Type，根据Material Number（PM Code）
        mat_code = int(df1.at[j,'Material Number'])
        if 250000 <= mat_code <= 299999:
            df1.at[j,'Material Type'] = 'PMSZ'
        elif 210000 <= mat_code <= 239999:
            df1.at[j,'Material Type'] = 'PAPC'
        elif 300000 <= mat_code <= 399999:
            df1.at[j,'Material Type'] = 'PMCH'
        else:
            df1.at[j,'Material Type'] = 'WRONG'    

        # Storage Location，Issue Storage Location,Default Storage Location根据Plant判断
        plant = df1.at[j,'Plant']
        if plant == '0023':
            df1.at[j,'Storage Location'] = df1.at[j,'Issue Storage Location'] = '0015'
            df1.at[j,'Default Storage Location'] = '0002'
        else:
            df1.at[j,'Storage Location'] = df1.at[j,'Issue Storage Location'] = '0006'
            df1.at[j,'Default Storage Location'] = '0003'
        
        # Minimum Remaining Shelf Life，判断是否为0
        if df1.at[j,'Total Shelf Life'] == '0':
            df1.at[j,'Minimum Remaining Shelf Life'] = '0'
        
        # Valuation Class,判断
        if df1.at[j,'Valuation Class'] == '3rd party':
            df1.at[j,'Valuation Class'] = '210'
        elif df1.at[j,'Valuation Class'] == 'ICY':
            df1.at[j,'Valuation Class'] = '220'
        else:
            df1.at[j,'Valuation Class'] = '211'
        
    # 多库位处理
    df00 = df1.copy(deep=True)
    # 临时存储df01,工厂为0078 & 0079的
    df01 = df00[df00['Plant'].isin(['0078','0079'])]   
    # 为SH工厂的物料，Storage Location，赋值0003,追加到df1
    df01['Storage Location'] = '0003'
    df1 = df1.append(df01)

    # 临时存储df02,工厂为0023的
    df02 = df00[df00['Plant'].isin(['0023'])]
    # 为SZ工厂的物料，Storage Location，赋值0002,追加到df1
    df02['Storage Location'] = '0002'
    df1 = df1.append(df02).reset_index(drop=True)

    print('Convert Value Done.')
    return df1

# ACC&CO,copy值
def copy_value2(df1,df2):
    # ACC&CO的值，直接从Z1UMM24中copy已整理好的值
    # 字段对应关系
    f_dict ={
               'Material Code':'Material Number',
                'English Matl Desc':'Material Description (English)',
                'Chinese Matl Desc':'Material Description (Chinese)',
                'Plant':'Plant',
                'Valuation Key':'Plant',
                'Valuation Class':'Valuation Class',
                'Price Unit':'Price Unit',
                'Standard Price':'Standard Price',
                'Future Price':'Future Price',
                # 有效期仅在有Future Price时启用
                # 'Future Price - Valid From (YYYYMMDD)':'Future Price - Valid From (YYYYMMDD)',
                'Lot Size for Product Costing':'Lot Size for Product Costing'
                }
    for k,v in f_dict.items():
        df2[k] = df1[v]

    print('ACC&CO copy value Done.')

# ACC&CO,固定值
def fix_value2(df2):
    # 字段关系
    f_dict = {
                   'Price Determination':'3',
                    'Material Ledger Indicator':'X',
                    'Price Control Indicator':'S',
                    'Material Origin':'X',
                    'Variance Key':'000001'
                    }
    for k,v in f_dict.items():
        df2[k] = v
    
    # 删除重复行
    df2.drop_duplicates(inplace=True)
    df2 = df2.reset_index(drop=True)

    print('ACC&CO fixed value Done.')
    return df2

# UOM,copy值
def copy_value3(df11,df3):
    # 字段关系
    f_dict ={
                'Material Code':'PM Code',
                'English Matl Desc':'包材描述（英文）',
                'Chinese Matl Desc':'包材描述（中文）',
                'Conversion Value (Y)':'单位转换'
                }
    
    # 筛选出有单位转换的，基本单位和采购单位不一致的
    for i in df11.index:
        bu = df11.at[i,'基本单位']
        ou = df11.at[i,'采购单位']
        if bu != ou:
            for k,v in f_dict.items():
                df3.at[i,k] = df11.at[i,v]

    print('UOM copy value Done.')

# UOM,固定值
def fix_value3(df3):
    # 继续整理df3
    # 复制df3到临时df0
    df0 = df3.copy(deep=True)

    # df3带单位转换的行，把固定字段赋值
    # 字段关系
    f_dict = {
                'Conversion Value (X)':'1',
                'Alternative UOM':'ROL',
                'Base UOM':'M'
                    }
    for k,v in f_dict.items():
        df3[k] = v

    # 处理df0,整理固定行后，追加到df3中
    # 字段关系
    f_dict2 = {
                'Conversion Value (X)':'1',
                'Alternative UOM':'M',
                'Conversion Value (Y)':'1',
                'Base UOM':'M'
                    }
    for k2,v2 in f_dict2.items():
        df0[k2] = v2
    # 追加到df3中
    df3 = df3.append(df0).reset_index(drop=True)

    # 删除重复行
    df3.drop_duplicates(inplace=True)
    df3 = df3.reset_index(drop=True)

    print('UOM fixed value Done.')
    return df3

# Class,copy值
def copy_value4(df11,df4):
    # 字段关系
    f_dict = {
                'Material':'PM Code',
                'Category 1':'Procurement group1',
                'Category 2':'Procurement group2',
                'Spec Code':'包装规格标准号',
                'Recyclable Plastic Material':'Recyclable Plastic Material'
                }
    for k,v in f_dict.items():
        df4[k] = df11[v]

    print('Class copy value Done.')

# Class,固定值
def fix_value4(df4):
    # 字段关系
    f_dict = {
                'Class':'Z_GP_PROC_CAT',
                'Material type':'PACKAGING',
                'Category 3':'00000087'
                    }
    for k,v in f_dict.items():
        df4[k] = v

    print('Class fixed value Done.')

# Class,拆分/合并值
def convert_value(df11,df4):
    # 按‘-’拆分
    f_list = ['Category 1','Category 2']
    for a in f_list:
        # df00历史存放拆分的结果
        df00 = df4[a].str.split('-',expand=True)
        # 取左边值，给原来的列
        df4[a] = df00[0]

    # Material structure
    df4['Material structure'] = df11['Material Structure'].str.slice(0,30)
    df4['Material structure 2'] = df11['Material Structure'].str.slice(31,60)

    # 删除重复行
    df4.drop_duplicates(inplace=True)
    df4 = df4.reset_index(drop=True)

    print('Class convert value Done.')
    return df4

# 生成模板
def GenaTemp():
    writer = pd.ExcelWriter('PM_Upload.xlsx')

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

    # 生成UOM上传模板
    # 获取UOM所有字段名
    f3 = FieldUOMList()
    df3 = pd.DataFrame(columns=f3)
    # 后面有对应的writer.save()
    df3.to_excel(writer,'UOM')

    # 生成Class上传模板
    # 获取PM的Classification所有字段名
    f4 = rpm_class_list()
    df4 = pd.DataFrame(columns=f4)
    # 后面有对应的writer.save()
    df4.to_excel(writer,'Class')

    writer.save()
    print('GenaTemp Done.')

# 数据整理
def GetSheet_pm():
    # 提示做预处理===========================
    text1 = '''
            1. 检查，物料号是否已在SAP中创建
            2. '工厂1'复制到‘工厂’，双工厂手工拆分
            3. 检查，价格是否为1/1000,若都没有‘/’时，必须给其中1行加‘/’
            如果OK，请输入123继续。。。'''
    print(text1)
    # 不OK就退出
    res = input('请输入：')
    if res != '123':
        print('程序退出。')
        sys.exit()
    
    # 生成模板=================================
    GenaTemp()

    # 获取数据源================================
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='PM',index_col=0,dtype=str)

    # 获取模板================================
    # Z1UMM24模板
    df1 = pd.read_excel('PM_Upload.xlsx',sheet_name='Z1UMM24',index_col=0,dtype=str)
    # ACC&CO模板
    df2 = pd.read_excel('PM_Upload.xlsx',sheet_name='ACC&CO',index_col=0,dtype=str)
    # UOM模板
    df3 = pd.read_excel('PM_Upload.xlsx',sheet_name='UOM',index_col=0,dtype=str)
    # Class模板
    df4 = pd.read_excel('PM_Upload.xlsx',sheet_name='Class',index_col=0,dtype=str)

    # 生成Z1UMM24数据================================
    # copy值
    copy_value(df11,df1)
    # 固定值
    fix_value(df1)
    # 拆分
    split_value(df11,df1)
    # 转换数据
    df1 = convert_value1(df11,df1)

    # 生成ACC&CO数据================================
    # ACC&CO,copy值
    copy_value2(df1,df2)
    # ACC&CO,固定值
    df2 = fix_value2(df2)

    # 生成UOM数据======================================
    # UOM,copy值
    copy_value3(df11,df3)
    # UOM,固定值
    df3 = fix_value3(df3)

    # 生成Z_GP_MATERIAL数据======================================
    # Class,copy值
    copy_value4(df11,df4)
    # Class,固定值
    fix_value4(df4)
    # Class,拆分/合并值
    df4 = convert_value(df11,df4)

    # 生成Rrmakd工作表======================================
    remark = {
            '0':'检查，物料号是否已在SAP中创建',
            '1':'UOM确认转换关系没问题，删除Delete?内容，删除无需转换的行',
            '2':'MM17补充,Order Unit,Unit of Issue',
            '3':'若有Future Price，填写Future Price - Valid From。',
            '4':'Classification 视图,HANA数据,检查Material structure 2有的，分2次上传',
            '5':'Classification 视图,FIFO_BATCH'
            }
    df44 = pd.Series(remark)

    # 保存生成的模板数据================================
    writer = pd.ExcelWriter('PM_Upload.xlsx')
    # 保存数据源
    df11.to_excel(writer,'PM')
    # 保存数据   
    df44.to_excel(writer,'Remark')
    df1.to_excel(writer,'Z1UMM24')
    df2.to_excel(writer,'ACCC&CO')
    df3.to_excel(writer,'UOM')
    df4.to_excel(writer,'Class')
    



    writer.save()
    print('Please check PM_Upload.xlsx')

# GenaTemp()
# GetSheet_pm()