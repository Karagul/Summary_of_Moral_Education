#
#
#
#
#
#  @ 韩馨空
# 程序  执行之前 一定要对  先按照 学号 排序  ，包括 汇总 项



import  xlrd
import xlwt

rBook = xlrd.open_workbook(r"D:\360MoveData\Users\Administrator\Desktop\5-德育分.xlsx")
rSheet = rBook.sheet_by_name("5月份")

List = []
# 得到 有效行数
lenth = rSheet.nrows
# 从的二行 开始 ，到 最终有效行结束
for row in range(1,lenth):
    temp=[]
    data = rSheet.cell_value(row,0)
    # 查找含有汇总 两字的 行  ，
    if "汇总" in data:
            # 得到 上一行的 学号 ，可以免去 汇总 二字
            matchData = rSheet.cell_value(row-1,0)
            # 得到  对应 行的  汇总 值
            value = rSheet.cell_value(row,3)
            # 添加学号，和 值到二级列表 List  中
            temp.append(matchData)
            temp.append(value)
            List.append(temp)

# 以utf-8 打开文档
wBook = xlwt.Workbook(encoding = 'utf-8')
# 添加 表
wSheet = wBook.add_sheet('总和')
# 用于写 标识 行
i = 1
# 遍历  所有总和数据  按照学号排序
for data in List:
    # 学号
    studentNumber = wSheet.write(i,0,label=data[0])
    # 总分
    value = wSheet.write(i,2,label=data[1])
    i += 1
# 保存为新文档
wBook.save(r"D:\360MoveData\Users\Administrator\Desktop\分.xlsx")