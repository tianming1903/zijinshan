import xlrd
import xlwt
import os,time,shutil

time1 = time.time()
#检测待爬取的文件并获取所有的文件
folder = next(os.walk('./'))[1]
files = next(os.walk(folder[0]))[2]

# 创建一个暂且存放信息的地方
if not os.path.exists('./info'):
    os.mkdir('info')

title = ''
# 对信息进行处理
for fi in files:
    book = xlrd.open_workbook(folder[0] + '/' + fi)
    table = book.sheets()[0]
    # 找出行业所在的列
    title = table.row_values(0)
    index = title.index('行业')
    # 包信息保存在每一类的文件里面
    names = table.col_values(0)[1:]
    for i in range(1,len(names)+1):
        info = table.row_values(i)
        with open('info/' + info[index] + '.txt','a',encoding='utf-8') as f:
            f.write(','.join(info).strip(','))

# 把信息写入到excel
files = next(os.walk('./info'))[2]
for fi in files:
    text = []
    with open ('info/' + fi,'r',encoding='utf-8') as f:
        for i in f:
            if info not in text:
                text.append(i.split(','))
    now = time.strftime('%m%d%H%M%S', time.localtime(time.time()))
    book = xlwt.Workbook()
    sheet = book.add_sheet('my_book')
    for x,y in enumerate(title):
        sheet.write(0, x, y)
    x = 1
    for infos in text:
        for num in range(len(infos)):
            sheet.write(x,num,infos[num])
        x += 1
    if not os.path.exists('分类好的'):
        os.mkdir('分类好的')
    book.save('./分类好的/' + fi.split('.')[0] + now + '.xls')
    time.sleep(1)

# 删除info文件
filename = next(os.walk('./info'))[2]
for x in filename:
    os.remove('./info/' + x)
os.removedirs('./info')
time2 = time.time()

print('信息归类已经完成')
print('耗时：' + str(time2-time1))
