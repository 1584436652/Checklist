from openpyxl import load_workbook


class CheckList:

    def __init__(self):
        self.wb = load_workbook('wuliu.xlsx')
        self.ws = self.wb.active

    def table(self):
        for row in self.ws.iter_rows():
            yield row

    def get_table(self):
        count = 0
        col = {}
        for r in self.table():

            if count == 0:
                if r not in col:
                    for c_name in r:
                        col[c_name.value] = None
                count -= 1
            else:
                print("-" * 20)
                # for k in col.keys():

                # for d in r:
                #     print(d)
                # print(*col.keys())


if __name__ == '__main__':
    ch = CheckList()
    ch.get_table()


