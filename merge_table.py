from detail_base import CheckList
from create_base import CreateTable


class Merge:
    def __init__(self):
        self.all_data = {"仓库": [], "物流": []}

    def depot_data(self):
        depot = CheckList('仓库.xlsx')
        de_l = []
        for de in depot.row_col():
            volume = de["材积"]
            volume_cut = volume.split('*')
            long, width, high = volume_cut
            # 材积重
            volume_weight = format((int(long)*int(width)*int(high)/6000-float(de["箱重"]))/2+float(de["箱重"]), '.3f')
            de["仓库长"] = int(long)
            de["仓库宽"] = int(width)
            de["仓库高"] = int(high)
            de["仓库材积重"] = float(volume_weight)
            de_l.append(de)
        self.all_data["仓库"] = de_l

    def logistics_data(self):
        logistics = CheckList('wuliu.xlsx')
        lo_d = []
        for lo in logistics.row_col():
            size = lo["尺寸"]
            size_cut = size.split('*')
            long, width, high = size_cut
            lo["物流长"] = int(long)
            lo["物流宽"] = int(width)
            lo["物流高"] = int(high)
            lo_d.append(lo)
            self.all_data["物流"] = lo_d

    def merge_data(self):
        """
        算差异
        :return:
        """
        l_data = {}
        d_data = {}
        a_l = []
        self.depot_data()
        self.logistics_data()
        for wl in self.all_data["物流"]:
            l_data[wl["转单号码"]] = [wl["物流长"], wl["物流宽"], wl["物流高"], wl["实重"], wl["材积"]]
        for wl in self.all_data["仓库"]:
            d_data[wl["箱号"]] = [wl["仓库长"], wl["仓库宽"], wl["仓库高"], wl["箱重"], wl["仓库材积重"]]
        for k, v in l_data.items():
            if k in d_data:
                difference = list(map(lambda x: format(x[0] - x[1], '.3f'), zip(v, d_data[k])))
                difference.insert(0, k)
                a_l.append(difference)
        return l_data, d_data, a_l

    @staticmethod
    def convert(item):
        item_d = []
        for k, v in item.items():
            v.insert(0, k)
            item_d.append(v)
        return item_d

    def saves(self):
        x, y, z = self.merge_data()
        cr = CreateTable()
        cr.ws.merge_cells('A1:F1')
        cr.ws['A1'] = "仓库"
        cr.cell_format('A1')
        cr.ws.merge_cells('H1:M1')
        cr.ws.merge_cells('O1:T1')
        for x_v in self.convert(x):
            cr.ws.append(x_v)
        print(x)
        # for y_v in self.convert(y):
        #     cr.ws.append(y_v)

        cr.wb.save('demo.xlsx')


if __name__ == '__main__':
    mer = Merge()
    mer.saves()



