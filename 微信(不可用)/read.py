'''
读取excel获取电话名单
名单必须是我抓取的名单格式
'''
import os,xlrd,re

# def read():
#     # 获取电话文件
#     num = 1
#     phones = next(os.walk('./phones'))[2]
#     for x in phones:
#         book = xlrd.open_workbook('./phones/' + x)
#         sheet = book.sheets()[0]
#         # 不同的表格修改此地方的索引值即可
#         phonelist = sheet.col_values(5)[1:]
#         for i in phonelist:
#             if len(i) == 11:
#                 with open('phones.txt','a',encoding='utf-8') as f:
#                     f.write(str(num) + ',' + i)
#                     f.write('\n')
#                 num += 1
# read()

def read():
    # 获取电话文件
    num = 1
    phones = next(os.walk('./phones'))[2]
    for x in phones:
        book = xlrd.open_workbook('./phones/' + x)
        for sheet in book.sheets():
        # 不同的表格修改此地方的索引值即可
            phonelist = sheet.col_values(0)[0:]
            for i in phonelist:
                i = str(i)[:-2]
                with open('phones.txt','a',encoding='utf-8') as f:
                    f.write(str(num) + ',' + i)
                    f.write('\n')
                num += 1
read()
