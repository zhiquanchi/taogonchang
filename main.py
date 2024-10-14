# -*- coding:utf-8  -*-
# @Time     : 2022/8/4 22:38
# @Author   : BGLB
# @Software : PyCharm

if __name__ == '__main__':
    from package_import import * # 必须导入 由于编译成了 pyd 所以 pyinstalle 找不到pyd里面的包
    from core import main
    main()



