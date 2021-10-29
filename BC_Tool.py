# 整理BC2.5,BC3数据
# BC2.5,BC3,8-12,OK

from RM_Tool import fix_value
import pandas as pd
import sys

# BC2,copy值
def copy_value(df11,df1):
    f_dict = {
            'Material':'产品代码',
            'Brand category 2.5':'Brand Category 2.5 Code'
            }
    for k,v in f_dict.items():
        df1[k] = df11[v]

    print('BC2.5 copy value done.')

# BC3,copy值
def copy_value2(df11,df2):
    f_dict = {
            'Material':'产品代码',
            'Declared Net Weight Text':'Brand Category 3 Code'
            }
    for k,v in f_dict.items():
        df2[k] = df11[v]    

    print('BC3 copy value done')

# BC3,拆分值
def split_value(df2):
    df2['Declared Net Weight Text'] = df2['Declared Net Weight Text'].str.extract('(.\d\d+[g])')
    df2['Declared Net Weight Text'] = df2['Declared Net Weight Text'].str.replace(' ','')

    print('BC3 split value done.')

# BC3TXT,copy值
def copy_value3(df11,df3):
    f_dict = {
            'Material Code':'产品代码',
            'Sales Text English':'Brand Category 3 Code',
            'Sales Text Chinese':'Brand Category 3'
            }
    for k,v in f_dict.items():
        df3[k] = df11[v]        

    print('BC3TXT copy value done.')

# BC3TXT,固定值
def fix_value1(df3):
    # 字段关系
    f_dict = {
            'Sales Org':'0078',
            'Distribution Channel':'01',
            'Material Description English':'N.A.',
            'Material Description Chinese':'N.A.'
                }
    for k,v in f_dict.items():
        df3[k] = v

    print('BC3TXT fixed value Done.')

# 生成模板
def GenaTemp():
    writer = pd.ExcelWriter('BC_Upload.xlsx')
    # 生成BC2.5上传模板
    # 获取BC2.5所有字段名
    f1 = ['Material','Class','Brand category 2.5']
    # 生成带标题的BC2.5的上传模板
    df1 = pd.DataFrame(columns=f1)
    # 后面有对应的writer.save()
    df1.to_excel(writer,'BC25')

    # 生成BC3上传模板
    # 获取BC3所有字段名
    f2 = ['Material','Class','Declared Net Weight Text']
    # 生成带标题的BC3的上传模板
    df2= pd.DataFrame(columns=f2)
    # 后面有对应的writer.save()
    df2.to_excel(writer,'BC3')

    # 生成BC3TXT上传模板
    # 获取BC3TXT所有字段名
    f3 = ['Material Code',
            'Sales Org',
            'Distribution Channel',
            'Sales Text English',
            'Sales Text Chinese',
            'Material Description English',
            'Material Description Chinese',
            ]
    df3= pd.DataFrame(columns=f3)
    df3.to_excel(writer,'BC3TXT')
    
    writer.save()
    print('GenaTemp Done')

# 数据整理
def GetSheet_bc():
    # 提示做预处理===========================
    text = '''预处理：
                1. BC2.5，BC3是否有数据
                2. 是否有remark
                如果OK，请输入12继续。。。'''
    print(text)
    # 不OK就退出
    res = input('请输入：')
    if res != '12':
        print('程序退出。')
        sys.exit()
    
    # 生成模板=================================
    GenaTemp()

    # 获取数据源================================
    df11 = pd.read_excel('DataSource.xlsx',sheet_name='BC',index_col=0,dtype=str)

    # 获取模板================================
    # BC25模板
    df1 = pd.read_excel('BC_Upload.xlsx',sheet_name='BC25',index_col=0,dtype=str)
    # BC3模板
    df2 = pd.read_excel('BC_Upload.xlsx',sheet_name='BC3',index_col=0,dtype=str)
    # BC3TXT模板
    df3 = pd.read_excel('BC_Upload.xlsx',sheet_name='BC3TXT',index_col=0,dtype=str)    

    # 生成BC2.5数据================================
    # copy值
    copy_value(df11,df1)
    # 固定值
    df1['Class'] = 'Z_BC25_MATERIAL'

    # 生成BC3数据================================
    # copy值
    copy_value2(df11,df2)
    # 拆分值
    split_value(df2)
    # 固定值
    df2['Class'] = 'Z_BC3_MATERIAL'

    # 生成BC3数据================================
    # copy值
    copy_value3(df11,df3)
    # 固定值
    fix_value1(df3)

    # 生成Rrmakd工作表======================================
    remark = {
            '1':'BC3重量超过3位数的，手工调整'
            }
    df44 = pd.Series(remark)

    # 保存生成的模板数据================================
    writer = pd.ExcelWriter('BC_Upload.xlsx')

    # 保存数据源
    df11.to_excel(writer,'BC')
    # 保存数据
    df44.to_excel(writer,'Remark')
    df1.to_excel(writer,'BC25')
    df2.to_excel(writer,'BC3')
    df3.to_excel(writer,'BC3TXT')

    writer.save()
    print('Pls check BC_Upload.xlsx')


# GenaTemp()
# GetSheet_bc()