
# xlwings  支持性全满

# 打开文档
# 查找 ‘4月份’表 ‘汇总’所在行的 E 列 值

# 复制 ’值‘   到 ‘总和’ 表 ‘学号’对应行 F 列



# A B C D E F 
# 1  2  3 4 5 6


# 导入xlwings模块，打开Excel程序，
# 默认设置：程序可见，
# 只打开不新建工作薄，
# 屏幕更新关闭
import xlwings as xw
import  time

app = xw.App(visible=True, add_book=False)
app.display_alerts = True
app.screen_updating = True

# 文件位置：filepath，打开test文档，
# 然后保存，关闭，结束程序



class Excel:
    Sht = ''  #  4月份工作表
    Sht2 = ''   #  总和 工作表

    Sht_Sum = ''    #汇总 项
    Sum_Item = ''   #去掉  汇总  二字后  的项


    def __init__(self,Path):
        self.File_Path = Path
        self.__Sum = 0
        self.__R_Sum = 0


    def Open_Books(self):
        Wb = app.books.open(self.File_Path)   # 打开工作簿
        sht = Wb.sheets['4月份']   # 打开工作表  4月份
        sht2 = Wb.sheets['总和']    # 打开工作表  总和
        time.sleep(5)
        sht_Sum = sht.range('C2').expand().value  # 从 '4月份' C2单元格 开始遍历 整个表
        self.Sht = sht
        self.Sht2 = sht2
        self.Sht_Sum = sht_Sum

        return sht_Sum


    def Get_Data(self):
        sum_Item = []       #总数据项 学号  总分
        single_Item= []   #单项数据获取，用做临时列表

        sum = 0    # 总行数
        r_sum = 0  # 总人数
        for i in self.Sht_Sum:         #遍历整个表
            sum +=1   #行数加一
            if "汇总" in i[0]:        # 每个行 数据中是否有含有  汇总 二字的  项
                single_Item.append(i[0])        # 将 学号 赋值给 临时列表
                single_Item.append(i[2])        # 将 总分 赋值给 临时列表
                sum_Item.append(single_Item)        # 将数据放入 总数据项列表 形成二级列表
                single_Item = []                # 清空临时列表
                r_sum += 1                      # 人数 加一
        self.Sum_Item = sum_Item                 #赋值给全局变量
        self.Sum = sum
        self.R_Sum = r_sum
        print(sum_Item)

        self.Sum_Item_Split()

        return  sum_Item


    def Sum_Item_Split(self):

    #字符分割
        for i in self.Sum_Item:
            temp = i[0].split()         # 将有  汇总 二字的项进行 分割  分割为学号
            i[0] = temp[0]              #将学号 赋值到源 列表项    【【201820183455 汇总，66】 】  变为 【【20183455，66】】


    def Match(self):

        for k in self.Sum_Item:
            for i in range(2,33):  # 2 -32
                if k[0] == self.Sht2.range(i,1).value:  # 上一个单元格的值是否与学号相等    # 如果学号相等 则复制到对应分值项
                    self.Sht2.range(i,6).value = k[1]           # 赋值  学号对应 值 项 B==6
                    print("yes",'学号{0}  值{1}'.format(k[0],k[1]))          #打印匹配的学号  与值

    def Get_R_sum(self):
        return self.__R_Sum

if __name__=="__main__":
    File_Path = r'D:\360MoveData\Users\Administrator\Desktop\deyufen.xlsx'
    ex = Excel(File_Path)
    ex.Open_Books()
    ex.Get_Data()
    ex.Match()

# wb.save()
# wb.close()
# app.quit()