# 小工具集合
import pandas as pd
import os,csv

# 连接多个表数据
def combine():
    df1 = pd.read_excel('DS.xlsx',sheet_name='0079',index_col=0,dtype={'Plant':str})
    df2 = pd.read_excel('DS.xlsx',sheet_name='0078',index_col=0,dtype={'Plant':str})
    df3 = pd.read_excel('DS.xlsx',sheet_name='0023',index_col=0,dtype={'Plant':str})
    df = pd.merge(df1,df2,on='Manufacturing Type code',how='outer')
    # df = pd.merge(df4,df3,on='Fact Prod Des',how='outer')

    writer = pd.ExcelWriter('Output.xlsx')
    df.to_excel(writer,'new')
    writer.save()
    print('Done')

# 合并多表，抽取几列
def combine2():
    # 读取数据源
    df1 = pd.read_excel('PMlist.xlsx',sheet_name='0023',index_col=0)
    df2 = pd.read_excel('PMlist.xlsx',sheet_name='0078',index_col=0)
    df3 = pd.read_excel('PMlist.xlsx',sheet_name='0079',index_col=0)
    df_list = [df1,df2,df3]

    # 生成结果表，指定列名
    f_list = ['Material number', 'Gross weight','Net weight','Weight Unit']
    df = pd.DataFrame(columns=f_list)
    
    # 合并数据,生成临时存放表df0
    for d in df_list:
        df0 = pd.DataFrame()
        for f in f_list:
            df0[f] = d[f]

        df = df.append(df0)
    # 重置索引号
    df = df.reset_index(drop=True)
    # 保存数据
    writer = pd.ExcelWriter('Output.xlsx')
    df.to_excel(writer,'new')
    writer.save()
    print('Done')

# 2个表，缩小范围
def reduce_scope():
    df1 = pd.read_excel('DataSource.xlsx',sheet_name='Sheet1',index_col='Material')
    df2 = pd.read_excel('DataSource.xlsx',sheet_name='Sheet2',index_col='Material',dtype={'EAN Code':str})
    # 按物料号遍历表2，按表1缩小范围
    for m in df2.index:
        # 如果物料号匹配，保留
        if m in df1.index:
            pass
        # 如果物料号不匹配，删除物料行
        else:
            df2 = df2.drop(index=[m])

    writer = pd.ExcelWriter('ouput.xlsx')
    df2.to_excel(writer,'df2')
    writer.save()
    print('Check Output.xlsx')

# 筛选值
def filtrate():
    df1 = pd.read_excel('DataSource.xlsx',sheet_name='Sheet1',index_col='Material')
    df2 = pd.read_excel('DataSource.xlsx',sheet_name='Sheet2',index_col='Material',dtype={'EAN Code':str})
    # 按物料号遍历表2
    for m in df2.index:
        # 如果物料号匹配，保留
        if m in df1.index:
            # 如果值匹配，删除物料行
            if df2.at[m,'EAN Code'] == '11223344':
                df2 = df2.drop(index=[m])
        # 如果物料号不匹配，删除物料行
        else:
            df2 = df2.drop(index=[m])

    writer = pd.ExcelWriter('ouput.xlsx')
    df2.to_excel(writer,'df2')
    writer.save()
    print('Check Output.xlsx')

# 合成字典格式
def gen_dict():
    df1 = pd.read_excel('DataSource.xlsx',sheet_name='Sheet1',index_col=0,dtype={'k':str,'v':str})
    # 按字典格式拼接2列
    df ="'" + df1['k'] + "':'" + df1['v'] + "',"
    # df ="'" + df1['k'] + "':" + "str,"

    writer = pd.ExcelWriter('output.xlsx')
    df.to_excel(writer,'df')
    writer.save()
    print('Check Output.xlsx')

# 合成列表格式
def gen_list():
    df1 = pd.read_excel('DataSource.xlsx',sheet_name='Sheet1',index_col=0,dtype=str)
    # 转换成列表格式
    df ="'" + df1['k'] + "',"

    writer = pd.ExcelWriter('output.xlsx')
    df.to_excel(writer,'df')
    writer.save()
    print('Check Output.xlsx')

# 拆分列
def split_column():
    df = pd.read_excel('DataSource.xlsx',sheet_name='Sheet1',index_col=0)
    df = df['k'].str.split('-',expand=True)

    writer = pd.ExcelWriter('ouput.xlsx')
    df.to_excel(writer,'df')
    writer.save()
    print('Check Output.xlsx')

# 检查2个表差异
def check_diff():
    f_num = input('请输入要匹对的列数：')

    # 2个表字段顺序要一致
    # 读取标准表
    df1 = pd.read_excel('DataSource.xlsx',sheet_name='Sheet2',index_col=0,dtype=str)
    # 读取SAP下载表
    df2 = pd.read_excel('DataSource.xlsx',sheet_name='Sheet3',index_col=0,dtype=str)
    # 为空单元格赋值，为了后面判断
    df1= df1.fillna(value=0)
    df2= df2.fillna(value=0)

    df0 = pd.DataFrame
    # 按关键字段来核对
    for a in df2.index:
        mat_cod = df2.at[a,'F1']
        # 用物料号来提取每个物料行到df0
        df0 = df2.loc[df2['F1'] == mat_cod]
        df0 = df0.reset_index(drop=True)

        # 找到df1中这个物料行放到临时df01
        df01 = df1.loc[df1['F1'] == mat_cod]
        df01 = df01.reset_index(drop=True)

        # 比对该物料行各字段值
        for b in df0.index:
            for f in df2.columns[:int(f_num)]:
                if df0.at[b,f] != df01.at[b,f]:
                    df2.at[a,f] ='WRONG' + str(df2.at[a,f])+ 'should be:' + str(df01.at[b,f])
    
    writer = pd.ExcelWriter('ouput.xlsx')
    df2.to_excel(writer,'df')
    writer.save()
    print('Check Output.xlsx')

# 生成带辅助列的表
def Generate_FZ(*columns):
    # 2个表列顺序要一致
    # 读取标准表
    df1 = pd.read_excel('DataSource.xlsx',sheet_name='Sheet2',index_col=0,dtype=str)
    # 读取SAP下载表
    df2 = pd.read_excel('DataSource.xlsx',sheet_name='Sheet3',index_col=0,dtype=str)
    # 为空单元格赋值，为了后面判断
    df1= df1.fillna(value='n.a')
    df2= df2.fillna(value='n.a')
    
    # 组成辅助列  
    df01 = df1.copy(deep=True)
    df02 = df2.copy(deep=True)
    df01['FZ'] = df01[columns[0]]
    df02['FZ'] = df02[columns[0]]
    
    for c in range(1,len(columns)):
        # print(columns[c])
        df01['FZ'] = df01['FZ'] + df01[columns[c]]
        df02['FZ'] = df02['FZ'] + df02[columns[c]]

    # 按FZ列排序
    df01.sort_values(by='FZ',inplace=True,ascending=False)
    df02.sort_values(by='FZ',inplace=True,ascending=False)

    writer = pd.ExcelWriter('output.xlsx')
    df01.to_excel(writer,'基准表')
    df02.to_excel(writer,'待定表')
    writer.save()
    print('Check Output.xlsx')

#核对2个表
def check_diff2():
    text = '''注意：Generate_FZ的参数要注意更改
            确认，请按1继续'''
    print(text)
    res = input('请输入相应序号：')
    if res == '1':
        # 正常检查
        Generate_FZ('Material Number','Plant','Storage Location')
        # for KA 检查
        # Generate_FZ('Material Code','Plant','Storage Location')
        # 其他检查
        # Generate_FZ('Material number')
        
    # 读取待定表
    df02 = pd.read_excel('output.xlsx',sheet_name='待定表',index_col='FZ',dtype=str)
    # 读取基准表
    df01 = pd.read_excel('output.xlsx',sheet_name='基准表',index_col='FZ',dtype=str)
    # 删除Index列
    df02.drop(columns='Index',inplace=True)
    df01.drop(columns='Index',inplace=True)
    # df00，存放核对结果
    df00 = df02.copy(deep=True)

    # 获取列名列表，FZ不算
    columns_l = list(df02.columns)
    # 遍历待定表
    for i in df02.index:
        for column in columns_l:
            # i = int(i)
            # 获得基准表相应数据
            v_df01 = df01.at[i,column]
            # 获得基准表相应数据
            v_df02 = df02.at[i,column]
            # 存放核对结果到df00表
            if v_df01 == v_df02:
                df00.at[i,column] = 'OK'
            else:
                df00.at[i,column] = v_df02 + '-should be-' + v_df01
    # 将待定表和核对结果表合并,按FZ排序
    df03 = df02.copy(deep=True)
    df03 = df03.append(df00)
    df03.sort_values(by='FZ',inplace=True,ascending=False)
    
    # 保存核对结果
    writer = pd.ExcelWriter('output.xlsx')
    df01.to_excel(writer,'基准表')
    df02.to_excel(writer,'待定表')
    df03.to_excel(writer,'核对结果')
    writer.save()
    print('Check output.xlsx')

# 按字段，删除重复项
def dele_dup():
    df1 = pd.read_excel('Changelog_2019.xlsx',sheet_name='Sheet1',index_col=0)
    df11 = df1.copy(deep=True)
    print('=============原始数据===========')
    print(df1.head())
    print(df1.columns)
    df11.drop_duplicates(['Object Value', 'Date'],inplace=True)
    df11 = df11.reset_index(drop=True)

    # 保存成文件
    writer = pd.ExcelWriter('ouput.xlsx')
    df11.to_excel(writer,'df11')
    writer.save()
    print('Check Output.xlsx')

# 从主表转换POTXT上传
def potxt():
    # 读取数据源
    df1 = pd.read_excel('POTXT.xlsx',sheet_name='POTXT',index_col=0)
    # df2 = pd.read_excel('Changelog_2018.xlsx',sheet_name='Sheet1',index_col=0)
    # df3 = pd.read_excel('Changelog_2017.xlsx',sheet_name='Sheet1',index_col=0)
    # df_list = [df1,df2,df3]

    # 生成结果表，指定列名
    f_list = ['序号',
                'Material Number',
                'Material Description (Chinese)',
                'Material Description (English)',
                'Text Language',
                'PO TXT']
    df = pd.DataFrame(columns=f_list)
    
    # copy value
    f1_list = ['序号',
                'Material Number',
                'Material Description (Chinese)',
                'Material Description (English)']
    for a in f1_list:
        df[a] = df1[a]

    # 固定值
    df['Text Language'] = 'E/1'

    # 转换值
    df['PO TXT'] = df1['Size/Dimensions'] + ';' + df1['Industry Std Desc']

    # 保存成文件
    writer = pd.ExcelWriter('ouput.xlsx')
    df.to_excel(writer,'df')
    writer.save()
    print('Check Output.xlsx')
# 临时测试程序
def temp01():
    # step1,缩小范围,step2,复制值
    df1 = pd.read_excel('DS.xlsx',sheet_name='Sheet1',index_col='Material',dtype=str)
    df2 = pd.read_excel('DS.xlsx',sheet_name='Sheet2',index_col='Material Number',dtype=str)
    # 按物料号遍历表1,按表2缩小范围
    f_dict = {'Procurement Group 1 - Desc.':'Procuremnet Group 1 Des.',
            'Procurement Group 2 - Desc.':'Procuremnet Group 2 Des.'}
    for m in df1.index:
        # 如果物料号匹配，保留copy值
        if m in df2.index:
            for k,v in f_dict.items():
                df1.at[m,k] = df2.at[m,v]

        # 如果物料号不匹配，删除物料行
        else:
            df1 = df1.drop(index=[m])

    writer = pd.ExcelWriter('ouput.xlsx')
    df1.to_excel(writer,'df1')
    writer.save()
    print('Check Output.xlsx')

# 读取文件夹下文件名
def doc_list():
    # 输入路径
    path = r'D:\PCN-JACN\Python\CN\Form'
    names = os.listdir(path)
    print(names)

    # 写入Output文档
    with open('Doclist.csv','w',newline='',encoding='utf-8') as f:
        writer = csv.writer(f)            
        writer.writerow(names)

    print('Done. Pls check Doclist.csv')

# potxt()
# filtrate()
# gen_dict()   
# gen_list()
# split_column()
# temp01()
# check_diff()
check_diff2()
# Generate_FZ('Material Number','Plant','Storage Location')
# combine2()
# dele_dup()
# doc_list()
