# 内销FG整理
# 8-11,新增7860及其库位：7860,7861,7862,7863,7960

from SFG_Tool import convert_value1
import pandas as pd
from Field_Tools import *
import sys

# Z1UMM24直接copy值
def copy_field(df11,df1):
    f = {
                'Material Number':'fg code',
                'Material Description (English)':'产品描述（英文）',
                'Material Description (Chinese)':'产品描述（中文）',
                'Base Unit of Measure':'出厂单位',
                'Material Group':'产品组',
                'Old Material Number':'此新品将替代哪个现有产品',
                'External Material Group':'GPM品牌',
                'Gross Weight':' 毛重（公斤/箱）',
                'Net Weight':'出厂单位标识净重（公斤/箱）',
                'Volume':'外箱体积（立方分米）',
                'Material Group: Packaging Materials':'品系1',
                'Industry Std Desc':'产品简称（中文）',
                'Material Group 1':'品系2',
                'Material Group 2':'产品线',
                'Material Group 3':'包装类别',
                'Material Group 4':'分装方式',
                'Material Group 5':'口味',
                'Product Hierarchy':'Product Hierarchy编号',
                'MRP Controller':'MRP 控制者',
                'Total Shelf Life':'总保质期（天）'
                    }
    # 遍历本次要创建的物料行
    for i in df11.index:
        for k,v in f.items():
            df1.at[i,k] = df11.at[i,v]

    print('Copy value Done')

# Z1UMM24固定字段赋值
def fix_value(df1):
    f_dict = {
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
                'Valuation Class':'根据NPD’VC值',
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

    print('Fixed value Done.')

# Z1UMM24,转换后copy值
def convert_value1(df11,df1):
    # 1. 拆分，Base Unit of Measure,MRP Controller
    # 需要拆分字段字典
    f_dict1 = {
                    'Base Unit of Measure':	'出厂单位',
                    'MRP Controller':	'MRP 控制者'
                }
    # 拆分字段值，列为单位
    for k1,v1 in f_dict1.items():
        df00 = df11[v1].str.split('-',expand=True)
        df1[k1] = df00[0]

    # 2. 转换，遍历本次要创建的物料行
    # 工厂字典
    plantDict = PlantDict()

    for i in df1.index:
        # Plant转换
        k2 = df11.at[i,'生产工厂']
        df1.at[i,'Plant'] = plantDict[k2]

        # Size/Dimensions转换
        size1 = df11.at[i,'外箱尺寸（长*宽*高，厘米）']
        size2 = df11.at[i,'分装方式1']
        df1.at[i,'Size/Dimensions'] = size1 + '(' + size2 + ')'    

    print('Convert value1 Done')   

# Z1UMM24,阶段I，条件判断字段值
def option_value_I(df11,df1):
    # 根据Material Number判断Material Type
    for i in df1.index:
        m = int(df1.at[i,'Material Number'])
        if 410000 <= m <= 499999:
            df1.at[i,'Material Type'] = 'FERC'
        elif 630000 <= m <= 649999:
            df1.at[i,'Material Type'] = 'FERF'
        else:
            df1.at[i,'Material Type'] = 'NA'

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
            df1.at[i,'Sales Organization'] = df1.at[i,'Delivering Plant'] = '0078'   #NP08，只做0078/01
            df1.at[i,'Distribution Channel'] = '01'  
            df1.at[i,'Account Assignment Group'] = '80'
            df1.at[i,'Item Category Group'] = 'NORM'     
            df1.at[i,'Procurement Type'] = 'E'
            df1.at[i,'Issue Storage Location'] = '0015'
            df1.at[i,'Default Storage Location'] = '0001'
        
        # Minimum Remaining Shelf Life
        df1.at[i,'Minimum Remaining Shelf Life'] = df1.at[i,'Total Shelf Life'] // 2

        # 规则检查
        vol = df1.at[i,'Volume']
        df1.at[i,'Volume']= round(vol,3)

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

    print('Optional value Done')

# Z1UMM24,阶段II，条件判断字段值
def option_value_II(df11,df1):
    # 根据Material Number判断Material Type
    for i in df1.index:
        m = int(df1.at[i,'Material Number'])
        if 410000 <= m <= 499999:
            df1.at[i,'Material Type'] = 'FERC'
        elif 630000 <= m <= 649999:
            df1.at[i,'Material Type'] = 'FERF'
        else:
            df1.at[i,'Material Type'] = 'NA'

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

        # Minimum Remaining Shelf Life
        s = int(df1.at[i,'Total Shelf Life'])
        df1.at[i,'Minimum Remaining Shelf Life'] = s // 2

        # 规则检查
        vol = float(df1.at[i,'Volume'])
        df1.at[i,'Volume']= round(vol,3)

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

    print('Optional value Done')

# Z1UMM24,阶段II，生产工厂 + 非生产工厂 (w/o.财务视图)
def copyline(df1):
    # 复制阶段I，做操作后再append到fgmain
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

# ACC&CO,阶段I,上传数据
# ACC,阶段I,copy字段
def copy_value2(df11,df2):
    f = {
            'Plant':	'生产工厂',
            'Valuation Key':	'生产工厂',
            'Material Code':	'fg code',
            'English Matl Desc':	'产品描述（英文）',
            'Chinese Matl Desc':	'产品描述（中文）'
            }
    for k,v in f.items():
        df2[k] = df11[v]
    
    print('Acc copy value Done')

# ACC,阶段I，固定字段赋值
def fix_value2(df2):
    f_dict = {
                'Price Determination':	'3',
                'Material Ledger Indicator':	'X',
                'Valuation Class':	'根据NPD’VC值',
                'VC:Sales Order Stk':	'VC值+‘T’',
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

# ACC,阶段I&II，工厂条件判断字段
def convert_value2(df2):
    # Plant转换成代码
    plants = PlantDict()
    for i in df2.index:
        pl = df2.at[i,'Plant']
        df2.at[i,'Plant'] = df2.at[i,'Valuation Key'] = plants[pl]

    print('ACCI convert value Done')

# Z_NPD_MATERIAL,阶段I，上传数据
def copy_value3(df11,df3):
    # copy字段对应字典
    f = {
        'Material':	'fg code',
        'Mother SKU Code':	'Mother Code',
        'NPD':	'NPD',
        'Promotion':	'Promo'
        }
    for k,v in f.items():
        df3[k] = df11[v]
    
    # Class固定赋值
    df3['Class'] = 'Z_NPD_MATERIAL'

    # 删除重复行
    df3.drop_duplicates(inplace=True)
    df3 = df3.reset_index(drop=True)

    print('NPD copy value Done')
    return df3

# Z_GL_MATERIAL,阶段I，上传数据
def copy_value4(df11,df4):
    # copy值
    df4['Material'] = df11['fg code']
    df4['Manufacturing Type'] = df11['Manufacturing Type Short']
    # 固定值
    df4['Class'] = 'Z_GL_MATERIAL'
    # 拆分代码
    df1 = df11['Manufacturing Type'].str.split('-',expand=True)
    df4['Manufacturing Type'] = df1[0]

    # 删除重复行
    df4.drop_duplicates(inplace=True)
    df4 = df4.reset_index(drop=True)

    print('Class copy value Done')
    return df4

# UOM,阶段II，上传数据
def fix_value3(df11,df2):
    df = df2.copy(deep=True)
    df['Material Code'] = df11['fg code']
    df['English Matl Desc'] = df11['产品描述（英文）']
    df['Chinese Matl Desc'] = df11['产品描述（中文）']
    df['Conversion Value (X)'] = df11['分装方式1']
    df['Delete?'] = df11['零售单位']
    df['Base UOM'] = 'CTN'
    df['C6'] = df11['Consumer unit package barcode']

    # 要固定的值做成字段
    valueDic = { 
                # 'Conversion Value (X)':[1,cv1,cv2],
                'Alternative UOM' : ['CTN','PDW','COW','NNW','TBD','TBD','TBD'],
                'Conversion Value (Y)':[1,1000,1000,1000,1,1,1]
                }

    # 如果要带PDW,COW这里要变成7
    for j in range(7):
        for i in df.index:
            df1 = df.copy(deep=True)
            for key,value in valueDic.items():
                df1[key] = value[j]
        for i in df1.index:
            if df1.at[i,'Alternative UOM'] == 'CTN':
                df1.at[i,'Conversion Value (X)'] = 1
                # 赋值Barcode相关字段
                df1.at[i,'EAN/UPC'] = df11.at[i,'Carton package barcode']
                df1.at[i,'EAN CAT'] = 'C7'
            elif df1.at[i,'Alternative UOM'] == 'PDW':
                w1 = df11.at[i,'实际净重（公斤/箱）']
                df1.at[i,'Conversion Value (X)'] = w1 * 1000
            elif df1.at[i,'Alternative UOM'] == 'COW':
                w3 = df11.at[i,'实际净重（公斤/箱）']
                df1.at[i,'Conversion Value (X)'] = w3 * 1000
            elif df1.at[i,'Alternative UOM'] == 'NNW':
                w2 = df11.at[i,'出厂单位标识净重（公斤/箱）']
                df1.at[i,'Conversion Value (X)'] = w2 * 1000
            else:
                pass

        df2 = df2.append(df1,ignore_index=True)


    print('UOM fix value Done')
    return df2

# ACC,阶段III,上传数据
def copy_value5(df11,df1):
    # RDC整理
    pl_list = ['7805','7810','7820','7830','7835','7840','7845','7850','7855','7860',
                '7865','7895']
    for pl in pl_list:
        df00 = df11.copy(deep=True)
        df00['Plant'] = df00['Valuation Key'] = pl
        df1 = df1.append(df00,ignore_index=True) 
    
    # 非生产工厂整理
    pl_list2 = ['0023','0078','0079']
    for pl2 in pl_list2:
        df01 = df11.loc[df11['Plant'] == pl2]
        df02 = df01.copy()
        if (pl2 == '0023') or (pl2 == '0079'):
            df02['Plant'] = df02['Valuation Key'] = '0078'
        elif pl2 == '0078':
            # df02['Plant'] = df02['Valuation Key'] = '0079'
            continue
        else:
            df02['Plant'] = df02['Valuation Key'] = 'WRONG'
        df1 = df1.append(df02,ignore_index=True)
    
    # 非生产工厂&RDC无需填写的字段调整 
    df1['With Quantity Structure'] = df1['Alternative BOM'] = df1['BOM Usage'] = None
    df1['Valuation Class'] = df1['VC:Sales Order Stk'] = '根据NPD的VC,VC + "T"'

    print('AccIII copy value Done')
    return df1     

# 生成模板
def GenaTemp():
    writer = pd.ExcelWriter('FGHM_Upload.xlsx')

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

    # 生成Z_NPD_MATERIAL上传模板
    f4 = ['Material','Class','Mother SKU Code','NPD','Promotion']
    df4 = pd.DataFrame(columns=f4)
    df4.to_excel(writer,'NPD')

    # 生成Z_GL_MATERIAL上传模板
    f5 = FGHanaFields()
    df5 = pd.DataFrame(columns=f5)
    df5.to_excel(writer,'Z_GL_MATERIAL')

    writer.save()
    print('GenaTemp Done.')

# 调用模板，调用生成数据函数，存档为可上传格式
# 阶段I
    # NPD08预创建,创建生产工厂全部（包括财务视图），不分SH/SZ，必须创建销售组织0078/01，单库位
    # 注：阶段I不做多库位行
def GetSheet_fgI():
    # 提示做预处理===========================
    text1 = '''预处理：
                1. 检查PH,content是否有问题
                2. NPD08，外箱尺寸（长*宽*高，厘米）用 **
                3. NPD08，MRP 控制者 用 000-
                4. 双工厂生产，调整成2行
                5. 看一下Valuation Class
                6. 分配code注意SH/SZ
                如果OK，请输入123456继续。。。'''
    print(text1)
    # 不OK就退出
    res = input('请输入：')
    if res != '123456':
        print('程序退出。')
        sys.exit()

    # 生成模板=================================
    GenaTemp()

    # 获取数据源================================
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='FG',index_col=0,dtype=str)

    # 获取模板================================
    # Z1UMM24模板
    df1 = pd.read_excel('FGHM_Upload.xlsx',sheet_name='Z1UMM24',index_col=0,dtype=str)
    # ACC&CO模板
    df2 = pd.read_excel('FGHM_Upload.xlsx',sheet_name='ACC&CO',index_col=0,dtype=str)
    # 获得Z_NPD_MATERIAL模板
    df3 = pd.read_excel('FGHM_Upload.xlsx',sheet_name='NPD',index_col=0,dtype=str)
    # 获得Z_GL_MATERIAL模板
    df4 = pd.read_excel('FGHM_Upload.xlsx',sheet_name='Z_GL_MATERIAL',index_col=0,dtype=str)

    # 生成Z1UMM24数据================================
    # copy值
    copy_field(df11,df1)
    # 固定值
    fix_value(df1)
    # 取值转换
    convert_value1(df11,df1)
    # 阶段I，条件判断字段值
    option_value_I(df11,df1)

    # 生成ACC&CO数据================================
    # copy ACC值
    copy_value2(df11,df2)
    # 固定值
    fix_value2(df2)
    # 取值转换 
    convert_value2(df2)
    
    # 生成Z_NPD_MATERIAL数据================================
    # copy NPD值
    df3 = copy_value3(df11,df3)

    # 生成Z_GL_MATERIAL数据================================
    # copy Class值
    df4 = copy_value4(df11,df4)

    # 生成备忘录
    remark = {
            '1':'调整Size',
            '2':'检查是否有WRONG的',
            '3':'调整 Valuation Class',
            '4':'阶段I，将Sales Organization和Dist.Channel调整为0078/01',
            '5':'Old material number,不能有None',
            '6':'维护Manufacture Type',
            '7':'NPD08的物料，创建后直接Dele.Flag',
            '8':'分配code注意SH/SZ',
            }
    df44 = pd.Series(remark)

    writer = pd.ExcelWriter('FGHM_Upload.xlsx')
    # 保存生成的3个模板数据
    df11.to_excel(writer,'FGHM')
    df44.to_excel(writer,'Remark')
    df1.to_excel(writer,'Z1UMM24')
    df2.to_excel(writer,'ACC&CO')
    df3.to_excel(writer,'NPD')
    df4.to_excel(writer,'Z_GL_MATERIAL')
     
    writer.save()
    print('Please check FGHM_Upload.xlsx')

#阶段II
    # 生产工厂其他库位，非生产工厂&RDC（w/o.财务视图）
    #  NPD15创建，生产工厂其他库位，非生产工厂&RDC（w/o.财务视图）
def GetSheet_fgII():
    #提示做预处理 =================================
    text1 = '''预处理：
                1. 双工厂生产，调整成2行
                如果OK，请输入1继续。。。'''
    print(text1)
    # 不OK就退出
    res = input('请输入：')
    if res != '1':
        print('程序退出。')
        sys.exit()

    # 生成模板 =================================
    GenaTemp()


    # 获取数据源================================
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='FG',index_col=0,dtype=str)

    # 获取模板================================
    # Z1UMM24模板
    df1 = pd.read_excel('FGHM_Upload.xlsx',sheet_name='Z1UMM24',index_col=0,dtype=str)
    # UOM模板
    df2 = pd.read_excel('FGHM_Upload.xlsx',sheet_name='UOM',index_col=0,dtype=str)
    # 获得Z_NPD_MATERIAL模板
    df3 = pd.read_excel('FGHM_Upload.xlsx',sheet_name='NPD',index_col=0,dtype=str)
    # 获得Z_GL_MATERIAL模板
    df4 = pd.read_excel('FGHM_Upload.xlsx',sheet_name='Z_GL_MATERIAL',index_col=0,dtype=str)
    
    # 生成Z1UMM24数据================================
    # copy值
    copy_field(df11,df1)
    # 固定值
    fix_value(df1)
    # 取值转换
    convert_value1(df11,df1)
    # 阶段II，条件判断字段值
    option_value_II(df11,df1)
    # 阶段II，阶段II，生产工厂 + 非生产工厂 (w/o.财务视图)
    df1 = copyline(df1)
    # 阶段II，阶段II，RDC (w/o.财务视图)
    df1 = copyRDCLine(df1)
    df1.sort_values(by='Material Number',inplace=True)
    
    # 生成UOM数据================================
    df2 = fix_value3(df11,df2)

    # 生成Z_NPD_MATERIAL数据================================
    # copy NPD值
    df3 = copy_value3(df11,df3)

    # 生成Z_GL_MATERIAL数据================================
    # copy Class值
    # df4 = copy_value4(df11,df4)

    # 生成Rrmakd工作表
    remark = {
            '1':'调整Size',
            '2':'检查是否有WRONG的',
            '3':'调整UOM,清除UOM中Delet?的内容',
            '4':'双工厂的已拆分为2行，删除重复工厂&RDC,保留生产工厂行，注意产出方式应保留E的',
            '5':'移除Dele. Flag再做',
            '6':'UOM,number per pack',
            '7':'UOM，C6 Barcode整理后，删除C6列'
            }
    df44 = pd.Series(remark)

    # 保存生成的模板数据================================
    writer = pd.ExcelWriter('FGHM_Upload.xlsx')

    # 保存数据源
    df11.to_excel(writer,'FG2')
    # 保存数据
    df44.to_excel(writer,'Remark')
    df1.to_excel(writer,'Z1UMM24')
    df2.to_excel(writer,'UOM')
    df3.to_excel(writer,'NPD')
    # df4.to_excel(writer,'Z_GL_MATERIAL')

    writer.save()
    print('Please check FGHM_Upload.xlsx')
    
# 阶段III，非生产工厂 + RDC的财务视图
def GetSheet_fgIII():
    #提示做预处理 =================================
    text1 = '''预处理：
                1. NPD36,是否已approve?
                2. Z1UMM18，下载生产工厂财务视图,放入ACC
                如果OK，请输入12继续。。。'''
    print(text1)
    # 不OK就退出
    res = input('请输入：')
    if res != '12':
        print('程序退出。')
        sys.exit()

    # 生成模板 =================================
    GenaTemp()

    # 获取ACC数据源================================
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='ACC',index_col=0,dtype=str)

    # 获得ACC&CO模板================================
    # Z1UMM24模板
    df1 = pd.read_excel('FGHM_Upload.xlsx',sheet_name='ACC&CO',index_col=0,dtype=str)

    # 生成ACC&CO数据================================
    # copy值
    df1 = copy_value5(df11,df1)
    
    # 生成Rrmakd工作表
    remark = {
            '1':'填STD Price',
            '2':'调整 Valuation Class'
            }
    df44 = pd.Series(remark)

    # 保存生成的模板数据================================
    writer = pd.ExcelWriter('FGHM_Upload.xlsx')
    
    # 保存数据源
    df11.to_excel(writer,'FGHM')

    # 保存数据
    df44.to_excel(writer,'Remark')
    df1.to_excel(writer,'ACC&CO')

    writer.save()
    print('Please check FGHM_Upload.xlsx')

# 选择阶段
def FGHM_main():
    text = '''
            FG 内销创建：
            1.阶段I，NPD08预创建
            2.阶段II，NPD15，非生产工厂&RDC创建，不创财务视图。生产工厂参数校正
            3.阶段III，Jason邮件阶段，创建财务视图
            请输入相应序号'''
    
    print(text)
    res = input('请输入相应序号：')
    # 阶段I,程序调用,带财务视图，不带UOM
    if res == '1':
        GetSheet_fgI()

    # 阶段II,程序调用,带UOM，不带财务视图
    elif res == '2':
        GetSheet_fgII()

    # 阶段III,程序调用,仅财务视图    
    elif res == '3':
        GetSheet_fgIII()
    else:
        print('程序退出。')
        sys.exit()


# GenaTemp()
# GetSheet_fgI()
# GetSheet_fgII()
# GetSheet_fgIII()
# FGHM_main()