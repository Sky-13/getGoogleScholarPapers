import xlrd
import xlwt

class Article(object):
    id = ""
    title = ""
    article_link = ""
    journal = ""

    def __init__(self):
        title = "New Paper"

# 读取xls文件
p = xlrd.open_workbook_xls('./diffusion_PaperInfo.xls')
# 按索引值打开sheet
sh1 = p.sheet_by_index(0)
# 获取行数
nrow = sh1.nrows
# 获取列数
ncol = sh1.ncols
# 获取所有行的内容
row_content = []
for i in range(0, nrow):
    paper = Article()
    rows = sh1.row_values(i)
    paper.id = rows[0]
    paper.title = rows[1]
    paper.article_link = rows[2]
    paper.journal = rows[3]
    row_content.append(paper)

# 获取最后一列的内容 期刊
cols = sh1.col_values(ncol-1)
# 期刊关键字
journal_key = 'IEEE'
key_list = []
for i, j in enumerate(cols):
    if journal_key in j:
        key_list.append(i)

TotalNum = 0
myxls = xlwt.Workbook(encoding='utf-8')
sheet1 = myxls.add_sheet(u'PaperInfo', True)
column = ['序号', '文章题目', '文章链接', '期刊']
# 首行写入
for i in range(0, len(column)):
    sheet1.write(TotalNum, i, column[i])
TotalNum += 1

for i in key_list:
    paper_ = Article()
    paper_ = row_content[i]
    sheet1.write(TotalNum, 0, TotalNum)
    sheet1.write(TotalNum, 1, paper_.title)
    sheet1.write(TotalNum, 2, paper_.article_link)
    sheet1.write(TotalNum, 3, paper_.journal)
    TotalNum += 1
    myxls.save(journal_key + '_PaperInfo.xls')