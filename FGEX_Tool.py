# 出口FG工具
# BGT,SZ&SH,8-4,OK
# NPD08,SZ&SH,8-4,OK
# NPD15,SZ&SH,8-4,OK

import pandas as pd
from Field_Tools import *
import sys

# 阶段BGT，copy值
def copy_value1(df11,df1):
    f_dict = {
                'Material Number':'fg code',
                'Material Description (English)':'产品描述（英文）',
                'Material Description (Chinese)':'产品描述（中文）',
                'Base Unit of Measure':'出厂单位',
                'Material Group':'产品组',
                'External Material Group':'GPM品牌',
                'Net Weight':'出厂单位标识净重（公斤/箱）',
                'Material Group: Packaging Materials':'品系1名称',
                'Material Group 1':'品系2名称',
                'Material Group 2':'产品线',
                'Material Group 3':'包装类别名称',
                'Material Group 4':'分装方式名称',
                'Material Group 5':'口味名称',
                'Product Hierarchy':'Product Hierarchy编号'
                    }
    # 遍历本次要创建的物料行
    for i in df11.index:
        for k,v in f_dict.items():
            df1.at[i,k] = df11.at[i,v]    
    
    print('PI,Copy value Done!')

# 阶段BGT，固定值
def fix_value1(df1):
    f_dict =  {
                'Industry Sector':'S',
                'Distribution Channel':'03',
                'Cross-Plant Material Status':'03',
                'Weight Unit':'KG',
                'Volume Unit':'DM3',
                'Class Type':'022',
                'Class':'FIFO_BATCH',
                'Country Key':'CN',
                'Tax Data - Tax Category':'MWST',
                'Tax Data - Tax Class 1':'0',
                'Cash Discount Indicator':'X',
                'Material Statistics Group':'1',
                'Account Assignment Group':'81',
                'General Item Category Group':'NORM',
                'Availability Check':'02',
                'Material Pricing Group':'01',
                'Country of Origin':'CN',
                'Batch Management Requirement Ind':'X',
                'MRP Group':'ZVR1',
                'Plant-Specific Material Status':'08',
                'MRP Type':'X0',
                'MRP Controller':'000',
                'Lot Size':'ZW',
                'Rounding Value for Purc Order Qty':'1',
                'Scheduling Margin Key for Floats':'000',
                'Period Indicator':'P',
                'Fiscal Year Variant':'Z1',
                'Consumption Mode':'1',
                'BWD Consumption Period':'030',
                'Procurement Type':'E',
                'Period indicator for SLED':'D',
                'MRP Relevancy for Dependent Requirements':'1',
                'Valuation Class':'420',
                'Price Control Indicator':'S',
                'Price Unit':'1',
                'BOM Usage':'5',
                'Alternative BOM':'1',
                'Lot Size for Product Costing':'100',
                'Price Determination':'3',
                'Material Ledger Indicator':'X',
                'With Quantity Structure':'X',
                'Material Origin':'X',
                'Variance Key':'000001'
                    }

    for k,v in f_dict.items():
        df1[k] = v

    print('PI,fix value Done')
    
# 阶段BGT，取值转换
def convert_value1(df11,df1):
    # 工厂字典
    plantDict = PlantDict()
    for i in df11.index:
        k = df11.at[i,'生产工厂']
        df1.at[i,'Plant'] = plantDict[k]

    # 需要拆分字段字典
    f_dict1 = {
                'Base Unit of Measure':	'出厂单位',
                'Material Group: Packaging Materials':'品系1名称',
                'Material Group 1':'品系2名称',
                'Material Group 3':'包装类别名称',
                'Material Group 4':'分装方式名称',
                'Material Group 5':'口味名称'
                }
    # 拆分字段值，列为单位
    for k1,v1 in f_dict1.items():
        df00 = df11[v1].str.split('-',expand=True)
        df1[k1] = df00[0]

    print('PI, convert value Done.')

# 阶段BGT，条件判断值
def option_value1(df1):
    # TBD值转换成BUG
    f_list = ['Material Group: Packaging Materials',
            'Material Group 1',
            'Material Group 2',
            'Material Group 3',
            'Material Group 4',
            'Material Group 5'
            ]
    for f in f_list:
        for j in df1.index:
            if df1.at[j,f] == 'TBD':
                df1.at[j,f] = 'BUG'
    
    # 格式化Material Group 2,Material Group 4
    df1['Material Group 2'] = df1['Material Group 2'].str.zfill(3)
    df1['Material Group 4'] = df1['Material Group 4'].str.zfill(3)

    # 根据Material Number判断Material Type
    for i in df1.index:
        m = df1.at[i,'Material Number']
        if 410000 <= int(m) <= 499999:
            df1.at[i,'Material Type'] = 'FERC'
        elif 630000 <= int(m) <= 649999:
            df1.at[i,'Material Type'] = 'FERF'
        else:
            df1.at[i,'Material Type'] = 'NA'

        # 根据Plant判断，Storage Location，Sales Organization,Item Category Group,Delivering Plant,Issue Storage Location,Default Storage Location
        plant1 = df1.at[i,'Plant']
        # SH工厂
        if plant1 != '0023':
            df1.at[i,'Storage Location'] =  '0006'
            df1.at[i,'Sales Organization'] = '0078'
            df1.at[i,'Item Category Group'] = 'NORM'
            df1.at[i,'Delivering Plant'] = '0078'
            df1.at[i,'Issue Storage Location'] = '0006'
            if plant1 == '0079':
                df1.at[i,'Default Storage Location'] = None
            else:
                df1.at[i,'Default Storage Location'] = '0001'
        # SZ工厂
        else:
            df1.at[i,'Storage Location'] = '0001'
            df1.at[i,'Sales Organization'] = '0023'     
            df1.at[i,'Item Category Group'] = 'ZORM'     
            df1.at[i,'Delivering Plant'] = '0023'  
            df1.at[i,'Issue Storage Location'] = '0015'
            df1.at[i,'Default Storage Location'] = '0001'

        # 规则检查
        vol = df1.at[i,'Volume']
        df1.at[i,'Volume']= round(vol,3)

        mg = df1.at[i,'Material Group']
        if int(mg) > 71:
            df1.at[i,'Material Group'] = 'WRONG'
        
        
        pl = df1.at[i,'Plant']
        if pl == '0079':
            df1.at[i,'Default Storage Location'] = None

        old = df1.at[i,'Old Material Number']
        if (old == 'None') or (old == 'NONE'):
            df1.at[i,'Old Material Number'] = None

    print('PI, option value done.')

# 阶段BGT，SH工厂多行处理
def copy_line(df1):
    # 复制第一阶段整理结果，做操作后再append到df1
    # df1原来的工厂替换掉
    df01 = df1.copy(deep=True)
    df00 = df1.copy(deep=True)
    df02 =  df00.loc[df00['Plant'].apply(lambda  x:x!='0023')]
    # 处理双工厂的数据
    for i in df1.index:
        pl = df1.at[i,'Plant']
        if pl == '0078,0079':
            df1.at[i,'Plant'] = '0079'
            df1.at[i,'Default Storage Location'] = None

    # 处理追加的行df01
    for i1 in df01.index:
        pl1 = df01.at[i1,'Plant']
        if pl1 == '0023':
            df01.at[i1,'Storage Location'] = '0015'
        elif pl1 == '0079':
            df01.at[i1,'Plant'] = '0078'
            df01.at[i1,'Storage Location'] = '0006'
            df01.at[i1,'Procurement Type'] = 'F'
            df01.at[i1,'Default Storage Location'] = '0001'
            df01.at[i1,'BOM Usage'] = df01.at[i1,'Alternative BOM'] = df01.at[i1,'With Quantity Structure'] = None
        elif pl1 == '0078,0079':
            df01.at[i1,'Plant'] = '0078'
            df01.at[i1,'Storage Location'] = '0006'
            df01.at[i1,'Default Storage Location'] = '0001'
        # pl1 == '0078'
        else:
            df01.at[i1,'Storage Location'] = '0001'

    # 处理追加的行df02,SH第3行数据
    for i2 in df02.index:
        pl2 = df02.at[i2,'Plant']
        if pl2 == '0079':
            df02.at[i2,'Plant'] = '0078'
            df02.at[i2,'Storage Location'] = df02.at[i2,'Default Storage Location'] =  '0001'
            df02.at[i2,'Procurement Type'] = 'F'
            df02.at[i2,'BOM Usage'] = df02.at[i2,'Alternative BOM'] = df02.at[i2,'With Quantity Structure'] = None
        elif pl2 == '0078,0079':
            df02.at[i2,'Plant'] = '0078'
            df02.at[i2,'Storage Location'] = df02.at[i2,'Default Storage Location'] =  '0001'
        # pl2 == '0078'
        else:
            df02.at[i2,'Plant'] = '0079'
            df02.at[i2,'Storage Location'] = '0006'
            df02.at[i2,'Procurement Type'] = 'F'
            df02.at[i2,'Default Storage Location'] = df02.at[i2,'BOM Usage'] = df02.at[i2,'Alternative BOM'] = df02.at[i2,'With Quantity Structure'] = None

    df1 = df1.append(df01,ignore_index=True).append(df02,ignore_index=True)
    print('PI, copy line done.')
    return df1

# ACC&CO上传数据
# ACC,copy字段
def copy_value2(df11,df2):
    f_dict = {
            'Plant':	'生产工厂',
            'Valuation Key':	'生产工厂',
            'Material Code':	'fg code',
            'English Matl Desc':	'产品描述（英文）',
            'Chinese Matl Desc':	'产品描述（中文）'
            }
    for k,v in f_dict.items():
        df2[k] = df11[v]

    print('Acc copy value Done')    

# ACC,固定字段
def fix_value2(df2):
    f_dict = {
                'Price Determination':	'3',
                'Material Ledger Indicator':	'X',
                'Valuation Class':	'420',
                'VC:Sales Order Stk':	'420T',
                'Price Control Indicator':	'S',
                'Price Unit':	'1',
                'Material Origin':	'X',
                'Variance Key':	'000001',
                'BOM Usage':	'5',
                'Alternative BOM':	'1',
                'With Quantity Structure':	'X',
                'Lot Size for Product Costing':	'100'
                }
    for i in df2.index:
        for k,v in f_dict.items():
            df2.at[i,k] = v

    print('ACC fixvalue Done')

# ACC,转换值
def convert_value2(df2):
    # 工厂字典
    plantDict = PlantDict()
    for i in df2.index:
        k = df2.at[i,'Plant']
        df2.at[i,'Plant'] = df2.at[i,'Valuation Key'] = plantDict[k]
    
    # 缓存一下，处理SH数据
    df00 = df2.copy(deep=True)

    # 处理双工厂生产的
    for i2 in df2.index:
        pl2 = df2.at[i2,'Plant']
        if pl2 == '0078,0079':
            df2.at[i2,'Plant'] = df2.at[i2,'Valuation Key'] = '0078'

    df01 = df00.loc[df00['Plant'].apply(lambda x: x != '0023')]
    for i3 in df01.index:
        pl3 = df01.at[i3,'Plant']
        if pl3 == '0078':
            df01.at[i3,'Plant'] = df01.at[i3,'Valuation Key'] ='0079'
            df01.at[i3,'BOM Usage'] = df01.at[i3,'Alternative BOM'] = df01.at[i3,'With Quantity Structure'] = None
        elif pl3 == '0079':
            df01.at[i3,'Plant'] = df01.at[i3,'Valuation Key'] ='0078'
            df01.at[i3,'BOM Usage'] = df01.at[i3,'Alternative BOM'] = df01.at[i3,'With Quantity Structure'] = None     
        # pl3 == '0078,0079'
        else:
            df01.at[i3,'Plant'] = df01.at[i3,'Valuation Key'] ='0079'

    df2 = df2.append(df01,ignore_index=0)
    print('ACC convert value Done')
    return df2

# Z_GL_MATERIAL,阶段BGT
def copy_value3(df11,df3):
    # copy值
    df3['Material'] = df11['fg code']
    df3['Manufacturing Type'] = df11['Manufacturing Type']
    # 固定值
    df3['Class'] = 'Z_GL_MATERIAL'
    # 拆分代码
    df1 = df11['Manufacturing Type'].str.split('-',expand=True)
    df3['Manufacturing Type'] = df1[0]

    # 删除重复行
    df3.drop_duplicates(inplace=True)
    df3 = df3.reset_index(drop=True)

    print('Class copy value Done')
    return df3   

# 阶段I&II,直接copy
def copy_value4(df11,df1):
    f_dict = {
                'Material Number':'fg code',
                'Old Material Number':	'此新品将替代哪个现有产品',
                'Material Description (English)':	'产品描述（英文）',
                'Material Description (Chinese)':	'产品描述（中文）',
                'Industry Std Desc':	'产品简称（中文）',
                'Material Group 5':	'口味',
                'Material Group 2':	'产品线',
                'Material Group 3':	'包装类别',
                'Material Group':	'产品组',
                'External Material Group':	'GPM品牌',
                'Material Group: Packaging Materials':	'品系1',
                'Material Group 1':	'品系2',
                'Material Group 4':	'分装方式',
                'Product Hierarchy':	'Product Hierarchy编号',
                'Total Shelf Life':	'总保质期（天）',
                'Net Weight':	'出厂单位标识净重（公斤/箱）',
                'Volume':	'外箱体积（立方分米）',
                'Gross Weight':	' 毛重（公斤/箱）'
                    }
    for k,v in f_dict.items():
        df1[k] = df11[v]
    
    print('PI&II, copy value done.')

# 阶段I&II,固定值
def fix_value3(df1):
    f_dict = {
                'Industry Sector':	'S',
                'Distribution Channel':	'03',
                'Cross-Plant Material Status':	'03',
                'Weight Unit':	'KG',
                'Volume Unit':	'DM3',
                'Class Type':	'022',
                'Class':	'FIFO_BATCH',
                'Country Key':	'CN',
                'Tax Data - Tax Category':	'MWST',
                'Tax Data - Tax Class 1':	'0',
                'Cash Discount Indicator':	'X',
                'Material Statistics Group':	'1',
                'Account Assignment Group':	'81',
                'General Item Category Group':	'NORM',
                'Availability Check':	'02',
                'Material Pricing Group':	'01',
                'Country of Origin':	'CN',
                'Batch Management Requirement Ind':	'X',
                'MRP Group':	'ZVR1',
                'Plant-Specific Material Status':	'08',
                'MRP Type':	'X0',
                'Lot Size':	'ZW',
                'Rounding Value for Purc Order Qty':	'1',
                'Scheduling Margin Key for Floats':	'000',
                'Period Indicator':	'P',
                'Fiscal Year Variant':	'Z1',
                'Consumption Mode':	'1',
                'BWD Consumption Period':	'030',
                'Period indicator for SLED':	'D',
                'Negative Stocks Allowed in Plant':	'',
                'MRP Relevancy for Dependent Requirements':	'1',
                'Valuation Class':	'420',
                'Valuation Category':'420T',
                'Price Control Indicator':	'S',
                'Price Unit':	'1',
                'Lot Size for Product Costing':	'100',
                'Price Determination':	'3',
                'Material Ledger Indicator':	'X',
                'Material Origin':	'X',
                'Procurement Type':'E',
                'BOM Usage':	'5',
                'Alternative BOM':	'1',
                'With Quantity Structure':	'X',
                'Variance Key':	'000001'
                    }
    for i in df1.index:
        for k,v in f_dict.items():
            df1.at[i,k] = v 
    print('PI&II, fix value done.')

# 阶段I&II,转换值
def convert_value3(df11,df1):
    # 工厂字典
    plantDict = PlantDict()
    for i in df11.index:
        k = df11.at[i,'生产工厂']
        df1.at[i,'Plant'] = plantDict[k]

        # Size/Dimensions转换
        size1 = df11.at[i,'外箱尺寸（长*宽*高，厘米）']
        size2 = df11.at[i,'分装方式1']
        df1.at[i,'Size/Dimensions'] = size1 + '(' + size2 + ')'

    # Base Unit of Measure,MRP Controller拆分
    # 需要拆分字段字典
    f_dict1 = {
                    'Base Unit of Measure':	'出厂单位',
                    'MRP Controller':	'MRP 控制者'
                }

    # 拆分字段值，列为单位
    for k1,v1 in f_dict1.items():
        df00 = df11[v1].str.split('-',expand=True)
        df1[k1] = df00[0]
    
    print('PI&II, convert value done.')

# 阶段I&II,条件判断值
def option_value2(df1):
    # 格式化Material Group 2,Material Group 4
    df1['Material Group 2'] = df1['Material Group 2'].str.zfill(3)
    df1['Material Group 4'] = df1['Material Group 4'].str.zfill(3)

    # 根据Material Number判断Material Type
    for i in df1.index:
        m = df1.at[i,'Material Number']
        if 410000 <= int(m) <= 499999:
            df1.at[i,'Material Type'] = 'FERC'
        elif 630000 <= int(m) <= 649999:
            df1.at[i,'Material Type'] = 'FERF'
        else:
            df1.at[i,'Material Type'] = 'NA'

        # 根据Plant判断，Storage Location，Sales Organization,Item Category Group,Delivering Plant,Issue Storage Location,Default Storage Location
        plant1 = df1.at[i,'Plant']
        # SH工厂
        if plant1 != '0023':
            df1.at[i,'Storage Location'] =  '0006'
            df1.at[i,'Sales Organization'] = '0078'
            df1.at[i,'Item Category Group'] = 'NORM'
            df1.at[i,'Delivering Plant'] = '0078'
            df1.at[i,'Issue Storage Location'] = '0006'
            if plant1 == '0079':
                df1.at[i,'Default Storage Location'] = None
            else:
                df1.at[i,'Default Storage Location'] = '0001'

        # SZ工厂
        else:
            df1.at[i,'Storage Location'] = '0001'
            df1.at[i,'Sales Organization'] = '0023'     
            df1.at[i,'Item Category Group'] = 'ZORM'     
            df1.at[i,'Delivering Plant'] = '0023'  
            df1.at[i,'Issue Storage Location'] = '0015'
            df1.at[i,'Default Storage Location'] = '0001'

        # Minimum Remaining Shelf Life
        s = df1.at[i,'Total Shelf Life']
        df1.at[i,'Minimum Remaining Shelf Life'] = float(s) // 2

    # 规则检查
        vol = df1.at[i,'Volume']
        df1.at[i,'Volume']= round(float(vol),3)

        mg = df1.at[i,'Material Group']
        if int(mg) > 71:
            df1.at[i,'Material Group'] = 'WRONG'
        
        mgpm = df1.at[i,'Material Group: Packaging Materials']
        if  400 <= int(mgpm) <= 499:
            pass
        else:
            df1.at[i,'Material Group: Packaging Materials'] = 'WRONG'
        
        mg1 = df1.at[i,'Material Group 1']
        if 300 <= int(mg1) <= 399:
            pass
        else:
            df1.at[i,'Material Group 1'] = 'WRONG'

        mg3 = df1.at[i,'Material Group 3']
        if mg3 == '00':
            df1.at[i,'Material Group 3'] = 'WRONG'

        mg4 = df1.at[i,'Material Group 4']
        if mg4 == '00':
            df1.at[i,'Material Group 4'] = 'WRONG'
        
        pl = df1.at[i,'Plant']
        if pl == '0079':
            df1.at[i,'Default Storage Location'] = None

        old = df1.at[i,'Old Material Number']
        if (old == 'None') or (old == 'NONE'):
            df1.at[i,'Old Material Number'] = None

    print('PI&II, option value done.')

# 阶段I&II,UOM值
def uom_value(df11,df4):
    df = df4.copy(deep=True)
    df['Material Code'] = df11['fg code']
    df['English Matl Desc'] = df11['产品描述（英文）']
    df['Chinese Matl Desc'] = df11['产品描述（中文）']
    df['Conversion Value (X)'] = df11['分装方式1']
    df['Delete?'] = df11['零售单位']
    df['Base UOM'] = 'CTN'
    
    # 要固定的值做成字段
    valueDic = { 
                # 'Conversion Value (X)':[1,cv1,cv2],
                'Alternative UOM' : ['CTN','PDW','COW','NNW','TBD','TBD','TBD'],
                'Conversion Value (Y)':[1,1000,1000,1000,1,1,1]
                }

    # 如果要带PDW,COW这里要变成7
    for j in range(7):
        for i in df.index:
            df00 = df.copy(deep=True)
            for key,value in valueDic.items():
                df00[key] = value[j]
        for i in df00.index:
            if df00.at[i,'Alternative UOM'] == 'CTN':
                df00.at[i,'Conversion Value (X)'] = 1
                # 赋值Barcode相关字段
                df00.at[i,'EAN/UPC'] = df11.at[i,'Carton package barcode']
                df00.at[i,'EAN CAT'] = 'C7'
            elif df00.at[i,'Alternative UOM'] == 'PDW':
                w1 = df11.at[i,'实际净重（公斤/箱）']

                df00.at[i,'Conversion Value (X)'] = float(w1) * 1000
            elif df00.at[i,'Alternative UOM'] == 'COW':
                w3 = df11.at[i,'实际净重（公斤/箱）']
                df00.at[i,'Conversion Value (X)'] = float(w3) * 1000
            elif df00.at[i,'Alternative UOM'] == 'NNW':
                w2 = df11.at[i,'出厂单位标识净重（公斤/箱）']
                df00.at[i,'Conversion Value (X)'] = float(w2) * 1000
            else:
                pass

        df4 = df4.append(df00,ignore_index=True)

    print('PI&II, uom value done.')
    return df4

# 生成模板
def GenaTemp():
    writer = pd.ExcelWriter('FGEX_Upload.xlsx')

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

    # 生成FG UOM上传模板
    f3 = FieldUOMList()   
    df3 = pd.DataFrame(columns=f3)
    df3.to_excel(writer,'UOM')

    # 生成Z_GL_MATERIAL上传模板
    f4 = FGHanaFields()
    df4 = pd.DataFrame(columns=f4)
    df4.to_excel(writer,'Z_GL_MATERIAL')

    writer.save()
    print('GenaTemp Done.')

# 阶段BGT，BGT整理，调用模板，调用生成数据函数，存档为可上传格式
def GetSheet_bgt():
    # 提示做预处理===========================
    text1 = '''预处理：
                1.分配code注意SH/SZ
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
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='BGT',index_col=0,dtype=str)
    
    # 获取模板================================
    # Z1UMM24模板
    df1 = pd.read_excel('FGEX_Upload.xlsx',sheet_name='Z1UMM24',index_col=0,dtype=str)
    # 获得ACC&CO模板
    df2 = pd.read_excel('FGEX_Upload.xlsx',sheet_name='ACC&CO',index_col=0,dtype=str)
    # 获得Z_GL_MATERIAL模板
    df3 = pd.read_excel('FGEX_Upload.xlsx',sheet_name='Z_GL_MATERIAL',index_col=0,dtype=str)

    # 生成Z1UMM24数据================================
    # copy值
    copy_value1(df11,df1)
    # 固定值
    fix_value1(df1)
    # 取值转换
    convert_value1(df11,df1)
    # 条件判断
    option_value1(df1)
    # SH工厂多行处理
    df1 = copy_line(df1)
    # 按物料号排序
    df1.sort_values(by='Material Number',inplace=True)

    # 生成ACC&CO数据================================
    # copy值
    copy_value2(df11,df2)
    # 固定值
    fix_value2(df2)  
    # 转换值
    df2 = convert_value2(df2)
    df2.sort_values(by='Material Code',inplace=True)

    # 生成Z_GL_MATERIAL数据================================
    df3 = copy_value3(df11,df3)

    # 生成Rrmakd工作表================================
    remark = {
            '1':'调整Size',
            '2':'ACC&CO,分工厂上传',
            '3':'维护Manu.Type',
            '4':'BGT的物料，创建后直接Dele.Flag',
            '5':'分配code注意SH/SZ'
            }
    df44 = pd.Series(remark)
    
    # 保存生成的模板数据================================
    writer = pd.ExcelWriter('FGEX_Upload.xlsx')
    df11.to_excel(writer,'FGEX')
    df44.to_excel(writer,'Remark')

    df1.to_excel(writer,'Z1UMM24')
    df2.to_excel(writer,'ACC&CO')
    df3.to_excel(writer,'Z_GL_MATERIAL')
    
    writer.save()
    print('Please check FGEX_Upload.xlsx')

# 阶段I&II,NPD08和NPD15整理
def GetSheet_fg():
    # 提示做预处理 =================================
    text1 = '''预处理：
                1. 分配code注意SH/SZ
                2. NPD08，外箱尺寸（长*宽*高，厘米）用 **
                3. NPD08，MRP 控制者 用 000-
                如果OK，请输入123继续。。。'''
    print(text1)
    # 不OK就退出
    res = input('请输入：')
    if res != '123':
        print('程序退出。')
        sys.exit()  

    # 生成模板 =================================
    GenaTemp()

    # 获取数据源================================
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='FG',index_col=0,dtype=str)

    # 获取模板================================
    # Z1UMM24模板
    df1 = pd.read_excel('FGEX_Upload.xlsx',sheet_name='Z1UMM24',index_col=0,dtype=str)
    # 获得ACC&CO模板
    df2 = pd.read_excel('FGEX_Upload.xlsx',sheet_name='ACC&CO',index_col=0,dtype=str)
    # 获得Z_GL_MATERIAL模板
    df3 = pd.read_excel('FGEX_Upload.xlsx',sheet_name='Z_GL_MATERIAL',index_col=0,dtype=str)
    # 获得UOM模板
    df4 = pd.read_excel('FGEX_Upload.xlsx',sheet_name='UOM',index_col=0,dtype=str)

    # 生成Z1UMM24数据================================
    # copy值
    copy_value4(df11,df1)
    # 固定值
    fix_value3(df1)
    # 转换值
    convert_value3(df11,df1)
    # 条件判断值
    option_value2(df1)

    # SH工厂多行处理
    df1 = copy_line(df1)
    # 按物料号排序
    df1.sort_values(by='Material Number',inplace=True)

    # 生成ACC&CO数据================================
    # copy值
    copy_value2(df11,df2)
    # 固定值
    fix_value2(df2)  
    # 转换值
    df2 = convert_value2(df2)
    df2.sort_values(by='Material Code',inplace=True)

    # 生成Z_GL_MATERIAL数据================================
    df3 = copy_value3(df11,df3)

    # 生成UOM数据================================
    # UOM值
    df4 = uom_value(df11,df4)

    # 生成Rrmakd工作表================================
    remark = {
            '1':'调整Size',
            '2':'检查是否有WRONG的',
            '3':'NPD15,调整UOM,number per pack',
            '4':'NPD15,清除UOM中Delet?的内容',
            '5':'NPD08新增/NPD15移除Dele.Flag',
            '6':'维护Manu.Type',
            '7':'分配code注意SH/SZ'
            }
    df44 = pd.Series(remark)

    # 保存生成的模板数据================================
    writer = pd.ExcelWriter('FGEX_Upload.xlsx')
    df11.to_excel(writer,'FGEX')
    df44.to_excel(writer,'Remark')

    df1.to_excel(writer,'Z1UMM24')
    df2.to_excel(writer,'ACC&CO')
    df3.to_excel(writer,'Z_GL_MATERIAL')
    df4.to_excel(writer,'UOM')
    
    writer.save()
    print('Please check FGEX_Upload.xlsx')

# 选择阶段
def FG_main():
    text = '''
            FG 出口创建：
            1.阶段BGT，BGT预创建
            2.阶段I&II，NPD08/NPD15
            请输入相应序号'''
    
    print(text)
    res = input('请输入相应序号：')
    # 阶段BGT，BGT预创建
    if res == '1':
        GetSheet_bgt()

    # 阶段I&II，NPD08/NPD15
    elif res == '2':
        GetSheet_fg()

    else:
        print('程序退出。')
        sys.exit()


# GenaTemp()
# GetSheet_bgt()
# GetSheet_fg()
# FG_main()