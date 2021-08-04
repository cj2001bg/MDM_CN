# 检查成品的HANA填写是否正确
# 生成HMDM顺序的数据表，供复制黏贴
# 描述不一样的要调整reference table


import pandas as pd
from Field_Tools import *

# HANAWeb,copy值
def copy_value(df11,df1):
    # 字段对应关系
    f_dict = {
            'Material Number':'fg code',
            'Short Description':'产品描述（英文）',
            'EAN':'Carton package barcode',
            'Category':'Category',
            'Primary pack type':'Primary pack type',
            'Net. Decl. Weight':'Primary Pack Net. Decl.Weight',
            'Net. Decl. Weight UoM':'Net. Decl.Weight UOM',
            'Num Piece per Pack':'Num Piece per Pack1',
            'Single Piece Wrapper':'Single piece wrapper',
            'Secondary Pack Type':'Secondary pack type',
            'Sec.Pack Type Inn.Qty':'分装方式1',
            'Base UoM - Hana':'Base UoM Inn.Qty',
            'Base UoM Inn.Qty':'分装方式1',
            'Additional pack remark':'Additional pack remark',
            'Brand':'Brand',
            'Sub Brand':'SubBrand',
            'Segment':'Segment',
            'Subsegment':'SubSegment',
            'Flavour Group':'Flavour group',
            'Flavour':'Flavour',
            'Manufacturing Type':'Manufacturing Type',
            'Fact Prod':'Fact Prod',
            'Fact Pack':'Fact Pack'
                }    
    for k,v in f_dict.items():
        df1[k] = df11[v]
    print('Copy value Done.')

# HANAWeb,固定值
def fix_value(df1):
    # 字段固定值
    f_dict = {
            'Better For You Claim':'No'
                }
    for k,v in f_dict.items():
        df1[k] = v
    
    print('Fixed value Done.')

# HANAWeb,拆分值
def split_value(df1):
    # 需拆分字段列表
    f_list = [
                'Net. Decl. Weight UoM',
                'Base UoM - Hana',
                'Manufacturing Type'
            ]
    # 逐列拆分，按‘-’取右边的
    for a in f_list:
        df00= df1[a].str.split('-',n=1,expand=True)
        df1[a] = df00[1]

    print('Split value Done.')

# HANAReport,copy值
def copy_value2(df11,df2):
    # 字段对应关系
    f_dict = {
            'Material':'fg code',
            'Description Short':'产品描述（英文）',
            'Category':'Category',
            'Brand':'Brand',
            'Sub Brand / Range':'SubBrand',
            'Segment':'Segment',
            'Sub Segment':'SubSegment',
            'Single Piece Wrapper':'Single piece wrapper',
            'Primary Pack Type':'Primary pack type',
            'Net Decl. Weight':'Primary Pack Net. Decl.Weight',
            'Net Declared Weight UOM':'Net. Decl.Weight UOM',
            'Secondary Pack Type':'Secondary pack type',
            'Secondary Pack Type Inner Quantity':'分装方式1',
            'Base UOM for Quantity Configuration':'Base UoM Inn.Qty',
            'BUOM Inner Quantity':'分装方式1',
            'Num Piece per Pack':'Num Piece per Pack1',
            'Additional Packaging Remark':'Additional pack remark',
            'Flavour Group':'Flavour group',
            'Flavour':'Flavour',
            'Manuf. Type':'Manufacturing Type',
            'Factory Production':'Fact Prod',
            'Factory Packing':'Fact Pack',
            'EAN Code':'Carton package barcode'
                }    
    for k,v in f_dict.items():
        df2[k] = df11[v]
    print('Copy value Done.')

# HANAReport,固定值
def fix_value2(df2):
    # 字段固定值
    f_dict = {
            'Better For You Claim':'No'
                }
    for k,v in f_dict.items():
        df2[k] = v
    
    print('Fixed value Done.')

# HANAReport,拆分值
def split_value2(df2):
    # 需拆分字段列表
    f_list = [
                'Net Declared Weight UOM',
                'Base UOM for Quantity Configuration',
                'Manuf. Type'
            ]
    # 逐列拆分，按‘-’取右边的
    for a in f_list:
        df00= df2[a].str.split('-',n=1,expand=True)
        df2[a] = df00[1]

    print('Split value Done.')

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
        subbrand = df1.at[i,'Sub Brand']
        segment = df1.at[i,'Segment']
        subsegment = df1.at[i,'Subsegment']
        manufacturing = df1.at[i,'Manufacturing Type']
        primary = df1.at[i,'Primary pack type']
        secondary = df1.at[i,'Secondary Pack Type']
        single = df1.at[i,'Single Piece Wrapper']
        flavourGroup = df1.at[i,'Flavour Group']
        flavour = df1.at[i,'Flavour']
        fact = df1.at[i,'Fact Prod']

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
    writer = pd.ExcelWriter('HANA_Check.xlsx')
    
    # 生成HANA Report格式 ，模板
    f1 = Hana_list()
    df1 = pd.DataFrame(columns=f1)
    df1.to_excel(writer,'HANAReport')

    #生成HANA 网页格式，模板
    f2 = Hana2_list()
    df2 = pd.DataFrame(columns=f2)
    df2.to_excel(writer,'HANAWeb')

    writer.save()
    print('GenaTemp Done.')

# 数据整理
def GetSheet_fghana():
    # 生成模板=================================
    GenaTemp()

    # 获取数据源================================
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='FG',index_col=0,dtype=str)

    # 获取模板================================
    df1 = pd.read_excel('HANA_Check.xlsx',sheet_name='HANAWeb',index_col=0,dtype=str)

    df2 = pd.read_excel('HANA_Check.xlsx',sheet_name='HANAReport',index_col=0,dtype=str)

    # 生成HANAWeb数据================================
    # copy值
    copy_value(df11,df1)
    # 固定值
    fix_value(df1)
    # 拆分值
    split_value(df1)

    # 生成HANAReport数据================================
    # copy值
    copy_value2(df11,df2)
    # 固定值
    fix_value2(df2)
    # 拆分值
    split_value2(df2)

    # 生成Rrmakd工作表======================================
    remark = {'1':'1',
            '2':'2'}
    df44 = pd.Series(remark)

    # 检查HANAWeb数据================================
    check_value(df1)

    # 保存生成的模板数据================================
    writer = pd.ExcelWriter('HANA_Check.xlsx')
    # 保存数据源
    df11.to_excel(writer,'FG')

    # 保存数据
    df44.to_excel(writer,'Remark')
    df1.to_excel(writer,'HANAWeb')
    df2.to_excel(writer,'HANAReport')

    writer.save()
    print('Pls check HANA_Check.xlsx')

# GenaTemp()
# GetSheet_fghana()