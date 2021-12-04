import xlrd


class ReadEx:

    def read_excel(self):
        # 打开excel表，填写路径
        book = xlrd.open_workbook("wuliu.xlsx")
        # 找到sheet页
        table = book.sheet_by_name("Sheet1")
        # 获取总行数总列数
        row_Num = table.nrows
        col_Num = table.ncols

        s =[]
        key =table.row_values(0)# 这是第一行数据，作为字典的key值

        if row_Num <= 1:
            print("没数据")
        else:
            j = 1
            for i in range(row_Num-1):
                d ={}
                values = table.row_values(j)
                for x in range(col_Num):
                    # 把key值对应的value赋值给key，每行循环
                    d[key[x]]=values[x]
                j += 1
                # 把字典加到列表中
                s.append(d)
            return s

if __name__ == '__main__':
    r = ReadEx()
    s=r.read_excel()
    for i in s:
        print(i)
    print(s)