# 拉一下以下号码段的主数据信息，如主数据没有为空就好
import pandas as pd
import numpy as np

# 缩小范围，copy值
def step_one():
    # Material Type Code vs Material Type Description
    f_dic2 = {
                'FERC':'SH: Finished Goods',
                'FERF':'SZ: Finished Goods',
                'SFCH':'Semi Finished China',
                'HALC':'Semi fin. prod. PVM',
                'ZCFH':'Semi Finished Product',
                'SZMX':'SZ: Mixtures',
                'SHMX':'SH: Mixtures',
                'RMCH':'SH: Raw Materials',
                'ROHC':'SZ: Packaging Materials',
                'RMSZ':'SZ: RM (200000-209999)',
                'PMCH':'SH: Packaging Materials',
                'PAPC':'SZ: Raw Materials',
                'PMSZ':'SZ: PM (250000-299999)',
                'POCH':'POP Materials Shanghai'
                }
    df1 = pd.read_excel('DataSource.xlsx',sheet_name='Sheet2',index_col=0)
    df = pd.DataFrame(columns=['Number Range',
                        'Deletion Flag',
                        'Material Category',
                        'Material Type Code',
                        'Material Type Description',
                        'Material code',
                        'Material Description'])

    # 排除物料范围
    mtc_list = list(f_dic2.keys())
    df0 = df1[df1['MTyp'].isin(mtc_list)]

    # copy值
    f_dic1 = {'Deletion Flag':'Clt','Material Type Code':'MTyp','Material code':'Material','Material Description':'Material description'}
    for k,v in f_dic1.items():
        df[k] = df0[v]
    
    df = df.reset_index(drop=True)

    writer = pd.ExcelWriter('ouput.xlsx')
    df.to_excel(writer,'df')
    writer.save()
    print('Step One Done')

# 根据判断
def step_two():
    # 第一步整理
    step_one()
     # Material Type Code vs Material Type Description
    f_dic1 = {
                'FERC':'SH: Finished Goods',
                'FERF':'SZ: Finished Goods',
                'SFCH':'Semi Finished China',
                'HALC':'Semi fin. prod. PVM',
                'ZCFH':'Semi Finished Product',
                'SZMX':'SZ: Mixtures',
                'SHMX':'SH: Mixtures',
                'RMCH':'SH: Raw Materials',
                'ROHC':'SZ: Raw Materials',
                'RMSZ':'SZ: RM (200000-209999)',
                'PMCH':'SH: Packaging Materials',
                'PAPC':'SZ: Raw Materials',
                'PMSZ':'SZ: PM (250000-299999)',
                'POCH':'POP Materials Shanghai'
                }
    # Material Type Code vs Material Category
    f_dic2 = {'FERC':'FG',
            'FERF':'FG',
            'SFCH':'SFG',
            'HALC':'SFG',
            'ZCFH':'SFG',
            'SZMX':'SFG',
            'SHMX':'SFG',
            'RMCH':'RM',
            'ROHC':'RM',
            'RMSZ':'RM',
            'PMCH':'PM',
            'PAPC':'PM',
            'PMSZ':'PM',
            'POCH':'POP'}
    # Material Type Code vs Number Range
    f_dic3 = {'FERC':'410000 - 499999',
                'FERF':'630000 - 649999',
                'SFCH':'511000 - 519999',
                'HALC':'500000 - 510999',
                'ZCFH':'511000 - 519999',
                'SZMX':'560001- 569999',
                'SHMX':'550001- 559999',
                'RMCH':'240000 - 249999',
                'ROHC':'210000 - 239999',
                'RMSZ':'200000 - 209999',
                'PMCH':'300000 - 399999',
                'PAPC':'210000 - 239999',
                'PMSZ':'250000 - 299999',
                'POCH':'930404 - 939999'}
    # ==================================================================
    df1 = pd.read_excel('ouput.xlsx',sheet_name='df',index_col=0,dtype={'Material Type Description':str,'Material Category':str,'Number Range':str})
    df = df1.copy(deep=True)
    for i in df.index:
        mtc = df.at[i,'Material Type Code']
        df.at[i,'Material Type Description'] = f_dic1[mtc]
        df.at[i,'Material Category'] = f_dic2[mtc]
        df.at[i,'Number Range'] = f_dic3[mtc]
        
    writer = pd.ExcelWriter('ouput.xlsx')
    df.to_excel(writer,'df')
    writer.save()
    print('Check Output.xlsx')

step_two()
