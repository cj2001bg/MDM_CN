# 创建物料工具集合

import sys,time
from SP_Tool import GetSheet_sp
from PM_Tool import GetSheet_pm
from FGEX_Tool import FG_main
from POP_Tool import GetSheet_pop
from SFG_Tool import GetSheet_sfg
from RM_Tool import GetSheet_rm
from FGHM_Tool import FGHM_main
from HANA_Tool import GetSheet_fghana
from BC_Tool import GetSheet_bc

while True:
    text1 = '''请翻牌子
                1. ID物料整理
                2. PM物料整理
                4. RM 整理
                5. SFG 整理
                6. FG(EX)物料整理
                7. FG(HT)物料整理
                8. FGHANA 整理检查
                9. POP 整理
                10.BC整理
                XX.敬请期待...
    '''
    print('===========我是分割线=============')
    print(text1)
    res = input('输入序号，或按任意键退出：')
    print('你的选择是：{0}'.format(res))

    if res == '1':
        GetSheet_sp()
        time.sleep(2)
    elif res == '2':
        GetSheet_pm()
        time.sleep(2)
    elif res == '4':
        GetSheet_rm()
        time.sleep(2) 
    elif res == '5':
        GetSheet_sfg()
        time.sleep(2)  
    elif res == '6':
        FG_main()
        time.sleep(2)
    elif res == '7':
        FGHM_main()
        time.sleep(2)  
    elif res == '8':
        GetSheet_fghana()
        time.sleep(2)  
    elif res == '9':
        GetSheet_pop()
        time.sleep(2)    
    elif res == '10':
        GetSheet_bc()
        time.sleep(2)  

    else:
        print('程序退出。')
        sys.exit()