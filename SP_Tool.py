# ID之SP物料整理
# 5-18,SH & SZ,OK

import pandas as pd
from Field_Tools import *
import sys

# Z1UMM24, copy值
def copy_value(df11,df1):
    # 字段关系，字段名相同
    f_list = [
            'Material Description (English)',
            'Material Description (Chinese)',
            'Material Type',
            'Plant',
            'Storage Location',
            'Base Unit of Measure',
            'Material Group',
            'Industry Std Desc',
            'Availability Check',
            'Country of Origin',
            'Order Unit',
            'ABC Indicator',
            'MRP Type',
            'MRP Controller',
            'Lot Size',
            'Rounding Value for Purc Order Qty',
            'Backflush',
            'Planned Delivery Time in Days',
            'Safety Stock',
            'Unit of Issue',
            'Storage Bin'
            ]
    for f in f_list:
        df1[f] = df11[f]
    
    # 字段关系，字段名有差异
    f_dict = {
        'Size/Dimensions':'Size+ Specification',
        'Purchasing Group':'Purchasing Group ',
        'Maximum Stock Level':'Maximum Stock',
        'Issue Storage Location':'Prod.stor.location',
        'Default Storage Location':'Storage Loc.for EP'
        }
    for k,v in f_dict.items():
        df1[k] = df11[v]

    print('Copy value Done.')

# Z1UMM24, 固定值
def fix_value(df1):
    # 字段关系
    f_dict = {
            'Industry Sector':'S',
            'Weight Unit':'KG',
            'General Item Category Group':'NORM',
            'Procurement Type':'F',
            'Scheduling Margin Key for Floats':'000',
            'Period Indicator':'P',
            'Fiscal Year Variant':'Z1'
                }
    for k,v in f_dict.items():
        df1[k] = v

    print('Fixed value Done.')

# Z1UMM24,拆分值
def split_value(df1):
    # 字段列表
    f_list = [
            'Plant',
            'Base Unit of Measure',
            'Material Group',
            'Country of Origin',
            'Order Unit',
            'Unit of Issue'
            ]
    for f in f_list:
        # 按‘-’拆分，临时存放到df0中
        df0 = df1[f].str.split('-',expand=True)
        # 取左边值，赋值给原列
        df1[f] = df0[0]

    print('Split value Done.')

# Z1UMM24,转换数据
def convert_value(df11,df1):
    # 继续整理df1
    # 填充空白单元格，放入临时df0中，映射完毕后，用于判断
    df0 = df1.copy(deep=True)
    df0 = df0.fillna(value='empty')
    # print(df0.at[3,'Rounding Value for Purc Order Qty'])

    # 逐行判断处理
    for i in df1.index:
        # Order Unit,Unit of Issue,根据Base Unit of Measure判断
        f_list1 = ['Order Unit','Unit of Issue']
        for f1 in f_list1:
            if df1.at[i,f1] == df1.at[i,'Base Unit of Measure']:
                df1.at[i,f1] = None
        # Purchasing Value Key,根据Plant判断
        if df1.at[i,'Plant'] == '0023':
            df1.at[i,'Purchasing Value Key'] = '13'
        else:
            df1.at[i,'Purchasing Value Key'] = '3'

        # 根据Lot Size 判断，用df0的填充空值判断后，将结果写入df1中
        if df1.at[i,'Lot Size'] == 'HB':
            # 用df0的填充空值判断后，将结果写入df1中
            if(df0.at[i,'Maximum Stock Level'] == 'empty') or (df0.at[i,'Safety Stock'] == 'empty'):
                df1.at[i,'Lot Size'] = 'WRONG:Lot Size = HB, stock level cannot be empty'
            # 用df0的填充空值判断后，将结果写入df1中
            elif (df0.at[i,'Maximum Stock Level'] == 'empty')  and (df0.at[i,'Safety Stock'] == 'empty'):
                df1.at[i,'Lot Size'] = 'WRONG:Stock leve = empty,Lot Size = EX'
        elif df1.at[i,'Lot Size'] == 'ZW':
            ro = df0.at[i,'Rounding Value for Purc Order Qty']
            if ro == 'empty':
                # 用df0的填充空值判断后，将结果写入df1中
                df1.at[i,'Rounding Value for Purc Order Qty'] = 'WRONG:Lot Size = ZW，Rounding cannot be empty'
        
        # 根据Material Type判断
        mt = df1.at[i,'Material Type']
        # MRP Group
        if mt in ['ZS01','ZS02','ZC03']:
            df1.at[i,'MRP Group'] = 'ZVR5'
        elif mt in ['ZC05','ZC06','ZC07']:
            df1.at[i,'MRP Group'] = 'ZVR6'

        # Plant-Specific Material Status
        if mt in ['ZC06','ZC07']:
            if df11.at[i,'Used in to BOM'] == 'No':
                df1.at[i,'Plant-Specific Material Status'] = '02'

        # Valuation Class
        v_dict = {'ZC06':'230','ZC07':'230','ZSPV':'900','ZS02':'902'}
        if mt in v_dict:
            df1.at[i,'Valuation Class'] = v_dict[mt]

        # Price Determination
        if mt in ['ZSPV','ZS02']:
            df1.at[i,'Price Determination'] = '2'
            df1.at[i,'Price Control Indicator'] = 'V'
        else:
            df1.at[i,'Price Determination'] = '3'

        # Price Control Indicator,Price Unit,Standard Price
        # Material Ledger Indicator，Material Origin，Variance Key
        if mt in ['ZC06','ZC07','ZSPV','ZS02']:
            df1.at[i,'Price Unit'] = '1'
            df1.at[i,'Standard Price'] = df11.at[i,'Standard Price']
            df1.at[i,'Material Ledger Indicator'] = df1.at[i,'Material Origin'] = 'X'
            df1.at[i,'Variance Key'] = '000001'
         
    print('Convert value Done.')
    return df1

# POTXT,copy数据
def copy_value2(df11,df2):
    # 字段关系
    f_dict = {
            'Chinese Matl Desc':'Material Description (Chinese)',
            'English Matl Desc':'Material Description (English)'
            }
    for k,v in f_dict.items():
        df2[k] = df11[v]

    print('POTXT copy value Done.')

# POTXT,固定值
def fix_value2(df2):
    # 设置一个从1开始的行号
    j = 1
    for i in df2.index:
        j = j + 1
        df2.at[i,'Material'] = '=VLOOKUP(A' + str(j) + ',Z1UMM24!A:B,2,0)'
    df2['Text Language'] = 'E/1'

    print('POTXT fixed value Done.')

# POTXT,转换值
def convert_value2(df11,df2):
    # Purchase Order Text
    # 将df11的空单元格预填充，映射到df0中，否则会出现合并后为空的列
    df0 = df11.fillna(value='NA')
    df2['Purchase Order Text'] = df0['Size+ Specification']
    # 要合并字段列表
    f_list = ['Industry Std Desc',
                'Drawing Number',
                'Imported Equ. Description',
                'Imported Equ. Spec.',
                'Imported Equ.  Serial number']
    for f in f_list:
        if f != 'Drawing Number':
            df2['Purchase Order Text'] = df2['Purchase Order Text'] + ';' + df0[f]
        else:
            df2['Purchase Order Text'] = df2['Purchase Order Text'] + ';图号：' + df0[f]

    print('POTXT convert value Done.')
    return df2

# 生成上传模板
def GenaTemp():
    writer = pd.ExcelWriter('SP_Upload.xlsx')

    # 生成Z1UMM24上传模板
    # 获取Z1UMM24所有字段名
    f1 = FieldSAPALL()
    # 生成带标题的Z1UMM24的上传模板
    df1 = pd.DataFrame(columns=f1)
    # 后面有对应的writer.save()
    df1.to_excel(writer,'Z1UMM24')

    # 生成POTXT上传模板
    f2 = PoTxtList()
    df2 = pd.DataFrame(columns=f2)
    # 后面有对应的writer.save()
    df2.to_excel(writer,'POTXT')

    writer.save()
    print('GenaTemp Done.')

# 数据整理
def GetSheet_sp():
    # 提示做预处理===========================
    text = '''预处理：
                    1. 填写序号为MRT编号
                    2. 检查申请表逻辑
                    如果OK，请输入12继续。。。'''
    print(text)
    # 不OK就退出
    res = input('请输入：')
    if res != '12':
        sys.exit()

    # 生成模板=================================
    GenaTemp()

    # 获取数据源================================
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='SP',index_col=0,dtype=str)
    # 获取模板================================
    # Z1UMM24模板
    df1 = pd.read_excel('SP_upload.xlsx',sheet_name='Z1UMM24',index_col=0,dtype=str)
    # POTXT模板
    df2 = pd.read_excel('SP_upload.xlsx',sheet_name='POTXT',index_col=0,dtype=str)

    # 生成Z1UMM24数据================================
    # copy值
    copy_value(df11,df1)
    # 固定值
    fix_value(df1)
    # 拆分值
    split_value(df1)
    # 转换数据
    df1 = convert_value(df11,df1)

    # 生成POTXT数据================================
    # copy数据
    copy_value2(df11,df2)
    # 固定值
    fix_value2(df2)
    # 转换值
    df2 = convert_value2(df11,df2)

    # 生成Rrmakd工作表======================================
    remark = {
                '1':'备注是否需要加入POTXT',
                '2':'用于BOM的AM要维护:Reorder Point / Backflush = 1 / MRP Type = VB',
                '3':'ZC06/ZC07/ZSPV/ZS02,要做财务视图',
                '4':'如果有UOM，手动维护',
                '5':'查找WRONG信息',
                '6':'POTXT，要手工转成2种语言',
                '7':'POTXT，要删除NA,加上备注中信息'}
    df44 = pd.Series(remark)

    # 保存生成的模板数据================================
    writer = pd.ExcelWriter('SP_upload.xlsx')
    # 保存数据源
    df11.to_excel(writer,'SP')
    # 保存数据
    df44.to_excel(writer,'Remark')
    df1.to_excel(writer,'Z1UMM24')
    df2.to_excel(writer,'POTXT')
    
    writer.save()
    print('Please check SP_upload.xlsx')

# GenaTemp()
GetSheet_sp()