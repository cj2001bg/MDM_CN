import pandas as pd
from Field_Tools import *
import sys

# I阶段，Z1UMM24,copy值
def copy_field(df11,df1):
    f_list = [
        'Material Number',
        'Material Description (English)',
        'Material Description (Chinese)',
        'Material Type',
        'Plant',
        'Base Unit of Measure',
        'Material Group',
        'Old Material Number',
        'External Material Group',
        'Gross Weight',
        'Net Weight',
        'Volume',
        'Size/Dimensions',
        'Material Group: Packaging Materials',
        'Basic Material',
        'Industry Std Desc',
        'Material Group 1',
        'Material Group 2',
        'Material Group 3',
        'Material Group 4',
        'Material Group 5',
        'Product Hierarchy',
        'MRP Controller',
        'Procurement Type',
        'Minimum Remaining Shelf Life',
        'Total Shelf Life',
        'Valuation Class'
            ]
    # 遍历本次要创建的物料行
    for f in f_list:
        df1[f] = df11[f]

    print('Copy value Done')

# I阶段，Z1UMM24,固定值
def fix_value(df1):
    f_dict = {
            # 'Sales Organization':'0078',
            # 'Distribution Channel':'01',
            'Industry Sector':'S',
            'Cross-Plant Material Status':'03',
            'Weight Unit':'KG',
            'Volume Unit':'DM3',
            'Class Type':'022',
            'Class':'FIFO_BATCH',
            'Country Key':'CN',
            'Tax Data - Tax Category':'MWST',
            'Tax Data - Tax Class 1':'1',
            'Cash Discount Indicator':'X',
            'Material Statistics Group':'1',
            'Account Assignment Group':'80',
            'Item Category Group':'NORM',
            'General Item Category Group':'NORM',
            'Availability Check':'02',
            'Material Pricing Group':'01',
            'Country of Origin':'CN',
            'Purchasing Value Key':'3',
            'Batch Management Requirement Ind':'X',
            'MRP Group':'ZVR1',
            'Plant-Specific Material Status':'08',
            'MRP Type':'X0',
            'Lot Size':'ZW',
            'Rounding Value for Purc Order Qty':'1',
            'Scheduling Margin Key for Floats':'000',
            'Period Indicator':'P',
            'Fiscal Year Variant':'Z1',
            'Consumption Mode':'1',
            'BWD Consumption Period':'030',
            'Period indicator for SLED':'D',
            'MRP Relevancy for Dependent Requirements':'1',
            'Price Determination':'3'
                }
    for k,v in f_dict.items():
        df1[k] = v

    print('Fixed value Done.')    

# I阶段，Z1UMM24,转换
def convert_value1(df1):
    # 替换代码
    code_dict = {
                '637014':'637167'
                    }
    
    for i in df1.index:
        # 替换代码
        code = df1.at[i,'Material Number']
        df1.at[i,'Material Number'] = code_dict[code]

        # 根据Plant判断，Storage Location，Sales Organization,Account Assignment Group,Item Category Group,Delivering Plant,Procurement Type,Issue Storage Location,Default Storage Location
        plant1 = df1.at[i,'Plant']
        # SH工厂
        if plant1 != '0023':
            df1.at[i,'Storage Location'] =  '0006'
            df1.at[i,'Sales Organization'] = df1.at[i,'Delivering Plant'] = '0078'   
            df1.at[i,'Distribution Channel'] = '01'
            df1.at[i,'Account Assignment Group'] = '80'
            df1.at[i,'Item Category Group'] = 'NORM'
            df1.at[i,'Delivering Plant'] = '0078'
            df1.at[i,'Procurement Type'] = 'E'
            df1.at[i,'Issue Storage Location'] = '0006'
            if plant1 == '0079':
                df1.at[i,'Default Storage Location'] = None
            else:
                df1.at[i,'Default Storage Location'] = '0001'
        # SZ工厂
        else:
            df1.at[i,'Storage Location'] = '0015'
            df1.at[i,'Sales Organization'] = df1.at[i,'Delivering Plant'] = '0078'   
            df1.at[i,'Distribution Channel'] = '01'  
            df1.at[i,'Account Assignment Group'] = '80'
            df1.at[i,'Item Category Group'] = 'NORM'     
            df1.at[i,'Procurement Type'] = 'E'
            df1.at[i,'Issue Storage Location'] = '0015'
            df1.at[i,'Default Storage Location'] = '0001'

    print('Optional value Done')

# ACC&CO,copy值
def copy_value2(df1,df2):
    f_dict= {
            'Material Code':'Material Number',
            'English Matl Desc':'Material Description (Chinese)',
            'Chinese Matl Desc':'Material Description (English)',
            'Plant':'Plant',
            'Valuation Key':'Plant',
            'Valuation Class':'Valuation Class'
                }
    for k,v in f_dict.items():
        df2[k] = df1[v]
    
    print('Acc copy value Done')

# ACC&CO,固定值
def fix_value2(df2):
        f_dict = {
                'Price Determination':	'3',
                'Material Ledger Indicator':	'X',
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

# ACC&CO,转换值
def convert_value2(df2):
    for i in df2.index:
        vc = df2.at[i,'Valuation Class']
        df2.at[i,'VC:Sales Order Stk'] = vc +'T'

    print('ACC convert value Done')

# 阶段II，Z1UMM24,条件判断字段值
def option_value_II(df1):
    for i in df1.index:
        # 根据Plant判断，Storage Location，Sales Organization,Account Assignment Group,Item Category Group,Delivering Plant,Procurement Type,Issue Storage Location,Default Storage Location
        plant1 = df1.at[i,'Plant']
        # SH工厂
        if plant1 != '0023':
            df1.at[i,'Storage Location'] =  '0006'
            df1.at[i,'Sales Organization'] = df1.at[i,'Delivering Plant'] = '0078'   
            df1.at[i,'Distribution Channel'] = '01'
            df1.at[i,'Account Assignment Group'] = '80'
            df1.at[i,'Item Category Group'] = 'NORM'
            df1.at[i,'Delivering Plant'] = '0078'
            df1.at[i,'Procurement Type'] = 'E'
            df1.at[i,'Issue Storage Location'] = '0006'
            if plant1 == '0079':
                df1.at[i,'Default Storage Location'] = None
            else:
                df1.at[i,'Default Storage Location'] = '0001'
        # SZ工厂
        else:
            df1.at[i,'Storage Location'] = '0015'
            df1.at[i,'Sales Organization'] = df1.at[i,'Delivering Plant'] = '0023'   
            df1.at[i,'Distribution Channel'] = '04'  
            df1.at[i,'Account Assignment Group'] = '83'
            df1.at[i,'Item Category Group'] = 'ZORM'     
            df1.at[i,'Procurement Type'] = 'E'
            df1.at[i,'Issue Storage Location'] = '0015'
            df1.at[i,'Default Storage Location'] = '0001'   

    print('II, optional value done.')

# 阶段II，生产工厂 + 非生产工厂 (w/o.财务视图)
def copyline(df1):
    # 复制阶段I，做操作后再append到df1
    df21 = df1.copy(deep=True)   #存放初始生产工厂数据，不变
    # 处理生产工厂是0078的
    df01 = df21.loc[df21['Plant'] == '0078']    # 找到SH生产工厂的值，存储到临时df30中
    pl_list1 = ['0078',
                '0079',
                '7805',
                '7810',
                '7820',
                '7830',
                '7835',
                '7840',
                '7845',
                '7850',
                '7855',
                '7860',
                '7865',
                '7895']
    for pl_sh in pl_list1:
        df31 = df01.copy()      #将符合条件的物料复制出来，df31可变
        df31['Plant'] = pl_sh
        # 调整费生产工厂&RDC参数
        if pl_sh != '0078':
            # 调整生产工厂&RDC共同参数
            df31['Purchasing Group'] = '780'
            df31['Procurement Type'] = 'F'
            df31['BOM Usage'] = df31['Alternative BOM'] = df31['With Quantity Structure'] = None

            # 专门调整一下0079工厂参数
            if pl_sh == '0079':     
                df31['Storage Location'] = '0006'
                # df31['MRP Controller'] = '000'
                df31['Default Storage Location'] = None

            # 专门调整一下RDC参数
            else:
                df31['Storage Location'] = df31['Issue Storage Location'] = df31['Default Storage Location'] =pl_sh
                df31['MRP Controller'] = '01'

        else:
            df31['Storage Location'] = '0001'
            df31['Distribution Channel'] = '04'
    
        df1 = df1.append(df31,ignore_index=True)

    # 处理生产工厂是0079的
    df02 = df21.loc[df21['Plant'] == '0079']    # 找到SH生产工厂的值，存储到临时df30中
    pl_list2 = ['0078',
                '0078',
                '7805',
                '7810',
                '7820',
                '7830',
                '7835',
                '7840',
                '7845',
                '7850',
                '7855',
                '7860',
                '7865',
                '7895']
    # 0078计数器，用来判断是否为第二次copy行，调整库位的
    num = 0
    for pl_sh in pl_list2:
        df32 = df02.copy()      #将符合条件的物料复制出来，df31可变
        df32['Plant'] = pl_sh
        # 调整非生产工厂&RDC参数
        if pl_sh != '0079':
            # 调整生产工厂&RDC共同参数
            df32['Purchasing Group'] = '780'
            df32['Procurement Type'] = 'F'
            df32['BOM Usage'] = df32['Alternative BOM'] = df32['With Quantity Structure'] = None
            
            # 专门调整一下0078工厂参数
            if pl_sh == '0078':     
                # df32['MRP Controller'] = '000'
                df32['Default Storage Location'] = '0001'
                df32['Distribution Channel'] = '04'
                num +=1
                # 判断是否0078第二行，若是,则调整库位
                if num == 2:
                    df32['Storage Location'] = '0001'
            
            # 专门调整一下RDC参数
            else:
                df32['Storage Location'] = df32['Issue Storage Location'] = df32['Default Storage Location'] =pl_sh
                df32['MRP Controller'] = '01'
            
        df1 = df1.append(df32,ignore_index=True)

    # 处理生产工厂是0023的
    df03 = df21.loc[df21['Plant'] == '0023'] 
    pl_list3 = ['0023',
                '0078',
                '0078',
                '7805',
                '7810',
                '7820',
                '7830',
                '7835',
                '7840',
                '7845',
                '7850',
                '7855',
                '7860',
                '7865',
                '7895']
    # 0078计数器，用来判断是否为第二次copy行，调整库位的
    num3 = 0
    for pl_sz in pl_list3:
        df33 = df03.copy()      #将符合条件的物料复制出来，df31可变
        df33['Plant'] = pl_sz
        # 调整非生产工厂&RDC参数
        if pl_sz != '0023':
                # 调整生产工厂&RDC共同参数
                df33['Sales Organization'] = df33['Delivering Plant'] = '0078'
                df33['Account Assignment Group'] = '80'
                df33['Item Category Group'] = 'NORM'
                df33['Purchasing Group'] = '780'  
                df33['Procurement Type'] = 'F'
                df33['BOM Usage'] = df33['Alternative BOM'] = df33['With Quantity Structure'] = None

                # 专门调整一下0078工厂参数
                if pl_sz == '0078':     
                    df33['Distribution Channel'] = '01'
                    df33['Storage Location'] = df33['Issue Storage Location'] = '0006'
                    df33['MRP Controller'] = '000'
                    num3 +=1
                    # 判断是否0078第二行，若是,则调整库位
                    if num3 == 2:
                        df33['Storage Location'] = '0001'  

            # 专门调整一下RDC参数
                else:
                    df33['Storage Location'] = df33['Issue Storage Location'] = df33['Default Storage Location'] =pl_sz
                    df33['MRP Controller'] = '01'
        else:
            df33['Storage Location'] = '0001'

        df1 = df1.append(df33,ignore_index=True)

    print('Copy lines Done')
    return df1

# Z1UMM24,阶段II，RDC(w/o.财务视图)
def copyRDCLine(df1):
    # 复制阶段II的CopyLine结果，做多库位
    df21 = df1.copy(deep=True)   #存放初始生产工厂数据，不变
    
    # 追加RDC工厂多库位
    pl_dict = {'7805':['7806','7807','7808','7809','7905'],
                '7810':['7811','7812','7813','7910'],
                '7820':['7821','7822','7823','7920'],
                '7830':['7832','7831','7930'],
                '7835':['7836','7837','7838','7935'],
                '7840':['7841','7842','7843','7940'],
                '7845':['7846','7847','7848','7945'],
                '7850':['7851','7852','7853','7950'],
                '7855':['7856','7857','7858','7955'],
                '7860':['7861','7862','7863','7960'],
                '7865':['7866','7867','7868','7869','7965'],
                '7895':['7896','7897','7898','7899','7995']}
    # 遍历RDC工厂
    for k,v in pl_dict.items():
        df00 = df21.loc[df21['Plant'] == k]  #找到相应RDC行
        # 遍历每个RDC对应的库位
        for loc in v:        
            df01 = df00.copy()
            df01['Storage Location'] = loc
            df1 = df1.append(df01,ignore_index=True)

    print('Copy RDC Done')
    return df1

# UOM,替换code
def code_exchange(df3):
    # 替换代码
    code_dict = {
                '637014':'637167'
                    }
    
    for i in df3.index:
        # 替换代码
        code = df3.at[i,'Material Number']
        df3.at[i,'Material Number'] = code_dict[code]    
    print('UOM DONE')

# Check,HANAWeb
def check_value(df1):
    # 保存原版数据到临时df00，后面df1会变的
    df00 = df1.copy(deep=True)
    df1['Res'] = ''

    # 获取HANA依赖关系表================================
    # Check表获取
    c01 = pd.read_excel('HANA Reference Tables.xlsx',sheet_name='Check01',index_col=0,dtype=str)
    c02 = pd.read_excel('HANA Reference Tables.xlsx',sheet_name='Check02',index_col=0,dtype=str)
    c03 = pd.read_excel('HANA Reference Tables.xlsx',sheet_name='Check03',index_col=0,dtype=str)
    c04 = pd.read_excel('HANA Reference Tables.xlsx',sheet_name='Check04',index_col=0,dtype=str)
    c05 = pd.read_excel('HANA Reference Tables.xlsx',sheet_name='Check05',index_col=0,dtype=str)
    c06 = pd.read_excel('HANA Reference Tables.xlsx',sheet_name='Check06',index_col=0,dtype=str)
    c07 = pd.read_excel('HANA Reference Tables.xlsx',sheet_name='Check07',index_col=0,dtype=str)
    c08 = pd.read_excel('HANA Reference Tables.xlsx',sheet_name='Check08',index_col=0,dtype=str)

    for i in df1.index:
        category = df1.at[i,'Category']
        brand = df1.at[i,'Brand']
        subbrand = df1.at[i,'Sub Brand / Range']
        segment = df1.at[i,'Segment']
        subsegment = df1.at[i,'Sub Segment']
        manufacturing = df1.at[i,'Manuf. Type']
        primary = df1.at[i,'Primary Pack Type']
        secondary = df1.at[i,'Secondary Pack Type']
        single = df1.at[i,'Single Piece Wrapper']
        flavourGroup = df1.at[i,'Flavour Group']
        flavour = df1.at[i,'Flavour']
        fact = df1.at[i,'Factory Production']

        # Check01检查,Category,Brand,Subbrand
        res01 = c01[(c01.Category == category) & (c01.Brand == brand) & (c01.Subbrand == subbrand)].index.tolist()
        # 判断结果产生
        if res01 == []:
            res_txt01 = 'Category & Brand & Sub Brand不匹配/'
        else:
            res_txt01 = 'Category & Brand & Sub Brand没问题/'

        # Check02检查,Category,Segment
        res02 = c02[(c02.Category == category) & (c02.Segment == segment)].index.tolist()
        if res02 == []:
            res_txt02 = 'Category & Segment不匹配/'
        else:
            res_txt02 = 'Category & Segment没问题/'

        # Check03检查,Segment,Subsegment
        res03 = c03[(c03.Segment == segment) & (c03.SubSegment == subsegment)].index.tolist()
        if res03 == []:
            res_txt03 = 'Segment & Subsegment不匹配/'
        else:
            res_txt03 = 'Segment & Subsegment没问题/'

        # Check04检查,Category,Manufacturing Type
        res04 = c04[(c04.Category == category) & (c04.ManufacturingType == manufacturing)].index.tolist()
        if res04 == []:
            res_txt04 = 'Category & Manufacturing Type不匹配/'
        else:
            res_txt04 = 'Category & Manufacturing Type没问题/'
        
        # Check05检查,Primary pack type,Secondary Pack Type
        res05 = c05[(c05.PrimaryPackType == primary) & (c05.SecondaryPackType == secondary)].index.tolist()
        if res05 == []:
            res_txt05 = 'Primary pack type & Secondary Pack Type不匹配/'
        else:
            res_txt05 = 'Primary pack type & Secondary Pack Type没问题/'
        
        # Check06检查,Single Piece Wrapper，Primary pack type
        res06 = c06[(c06.SinglePieceWrapper == single) & (c06.PrimaryPackType == primary)].index.tolist()
        if res06 == []:
            res_txt06 = 'Single Piece Wrapper & Primary pack type不匹配/'
        else:
            res_txt06 = 'Single Piece Wrapper & Primary pack type没问题/'

        # Check07检查,Flavour Group，Flavour
        res07 = c07[(c07.FlavourGroup == flavourGroup) & (c07.Flavour == flavour)].index.tolist()
        if res07 == []:
            res_txt07 = 'Flavour Group & Flavour不匹配/'
        else:
            res_txt07 = 'Flavour Group & Flavour没问题/'

        # Check08检查,Manufacturing Type，Fact Prod
        res08 = c08[(c08.ManufacturingType == manufacturing) & (c08.FactoryProd == fact)].index.tolist()
        if res08 == []:
            res_txt08 = 'Manufacturing Type & Fact Prod不匹配/'
        else:
            res_txt08 = 'Manufacturing Type & Fact Prod没问题/'

        
        # 判断结果存储
        df1.at[i,'Res'] = res_txt01 + res_txt02 + res_txt03 + res_txt04 + res_txt05 + res_txt06 + res_txt07 + res_txt08 

    print('HANAWeb Check Done.')

# 生成模板
def GenaTemp():
    writer = pd.ExcelWriter('output.xlsx')

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
    print('Pls check output.xlsx.')

def GetSheet_4me():
    # 提示做预处理===========================
    text1 = '''预处理：
                1. 仅保留生产工厂
                2. 输入新旧code：code_dict,Z1UMM24和UOM中的
                '''
    print(text1)
    # 不OK就退出
    res = input('请输入：')
    if res != '12':
        print('程序退出。')
        sys.exit()

    # 生成模板=================================
    GenaTemp()

    # 获取数据源================================
    df11 = pd.read_excel('4ME.xlsx',sheet_name='Z1UMM24',index_col=0,dtype=str)
    df22 = pd.read_excel('4ME.xlsx',sheet_name='UOM',index_col=0,dtype=str)

    # 获取模板================================
    # Z1UMM24模板
    df1 = pd.read_excel('output.xlsx',sheet_name='Z1UMM24',index_col=0,dtype=str)
    # ACC&CO模板
    df2 = pd.read_excel('output.xlsx',sheet_name='ACC&CO',index_col=0,dtype=str)

    # I阶段，生成Z1UMM24数据================================
    # 生产工厂整理
    # copy值
    copy_field(df11,df1)
    # 固定值
    fix_value(df1)
    # 取值转换
    convert_value1(df1) 
    
    # 生成ACC&CO数据================================
    # copy ACC值
    copy_value2(df1,df2)   
    # 固定值 
    fix_value2(df2)
    # 取值转换
    convert_value2(df2) 

    # II阶段，生成Z1UMM24数据================================
    # 非生产工厂，RDC整理（w/o财务视图）
    # 阶段II，条件判断字段值
    option_value_II(df1)
    # 阶段II，生产工厂 + 非生产工厂 (w/o.财务视图)
    df1 = copyline(df1)
    # 阶段II，阶段II，RDC (w/o.财务视图)
    df1 = copyRDCLine(df1)
    df1.sort_values(by='Material Number',inplace=True)    

    # 生成UOM数据================================
    # 整体copy
    df3 = df22.copy(deep=True)
    # 替换代码
    code_exchange(df3)

    # 生成备忘录
    remark = {
            '1':'UOM的PDW/COW手工调整',
            '2':'如有多工厂生产的，手动修改参数',
            '3':'UOM的Material Number2删掉'
            }
    df44 = pd.Series(remark)    

    writer = pd.ExcelWriter('output.xlsx')
    # 保存生成的模板数据
    df44.to_excel(writer,'Remark')
    df1.to_excel(writer,'Z1UMM24')
    df2.to_excel(writer,'ACC&CO')
    df3.to_excel(writer,'UOM')
     
    writer.save()
    print('Please check output.xlsx')    

def check_HANA():
    # 获取数据源================================
    df11 = pd.read_excel('4ME.xlsx',sheet_name='HANA',index_col=0,dtype=str)
    df1 = df11.copy(deep=True)
    # 检查HANAWeb数据================================
    check_value(df1)    
    for i in df1.index:
        code = df1.at[i,'Material']
        df1.at[i,'Material'] = '000000000000' + code

    # 保存生成的模板数据================================
    writer = pd.ExcelWriter('output2.xlsx')

    # 保存数据源
    df11.to_excel(writer,'HANA')     
    df1.to_excel(writer,'Check')    

    writer.save()
    print('Pls check output2.xlsx')    

GetSheet_4me()
check_HANA()

    