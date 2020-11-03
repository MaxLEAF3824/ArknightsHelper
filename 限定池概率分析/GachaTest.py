import numpy as np
from random import random
from matplotlib import pyplot as plt
import xlwt
import pandas as pd


class GachaTest:
    def __init__(self, up_rate=0.5, up_num=2):
        self.total_pull_count = 0  # 总抽卡数计数器
        self.pity_indicator = 0  # 保底机制计数器
        self.star6_pulled = 0  # 6星角色计数器
        self.star5_pulled = 0  # 5星角色计数器
        self.star4_pulled = 0  # 4星角色计数器
        self.star3_pulled = 0  # 3星角色计数器
        self.up_num = up_num  # 当期池子UP角色个数
        self.up_rate = up_rate  # 当期池子UP概率
        self.star6_pulled_result = []  # up角色计数器,其中0号位用来存非UP角色个数
        self.initialize()

    def initialize(self):
        self.total_pull_count = 0
        self.pity_indicator = 0
        self.star6_pulled = 0
        self.star5_pulled = 0
        self.star4_pulled = 0
        self.star3_pulled = 0
        self.star6_pulled_result.clear()
        for i in range(self.up_num + 1):
            self.star6_pulled_result.append(0)

    # 单抽（抽卡机制）
    def one_pull(self):
        self.total_pull_count += 1  # 总抽卡次数+1
        # 定义本次抽卡概率，前50次0.02，后面依次递增，99次必中
        self.pity_indicator += 1
        pull_rate = 0.02
        if self.pity_indicator > 50:
            pull_rate = 0.02 * (self.pity_indicator - 49)
        # 开始抽卡
        result = random()  # 当次抽卡结果
        if result < pull_rate:  # 出6星了
            self.pity_indicator = 0  # 重置保底
            self.star6_pulled += 1
            if random() < self.up_rate:  # 出UP角色了
                index = int(self.up_num * random() + 1)
                self.star6_pulled_result[index] += 1
            else:  # 没出UP角色
                self.star6_pulled_result[0] += 1
        elif result < 0.1:  # 出5星了
            self.star5_pulled += 1
        elif result < 0.6:  # 出4星了
            self.star4_pulled += 1
        else:  # 出3星了
            self.star3_pulled += 1

    # 出一个UP角色的模拟
    def get1UpSimulation(self, people_num):
        gacha_data = np.zeros([people_num, 3]).astype(int)
        for i in range(people_num):
            self.initialize()
            while sum(self.star6_pulled_result[1:len(self.star6_pulled_result)]) == 0:
                self.one_pull()
            gacha_data[i, 0] = self.total_pull_count
            gacha_data[i, 1] = self.star6_pulled
            gacha_data[i, 2] = self.star5_pulled
        return gacha_data

    # 2个UP角色都有的抽卡模拟
    def get2UpSimulation(self, people_num):
        def check_no_zero(list1):
            for i in list1:
                if i == 0:
                    return False
            return True

        gacha_data = np.zeros([people_num, len(self.star6_pulled_result) + 5]).astype(int)
        for i in range(people_num):
            self.initialize()
            while check_no_zero(self.star6_pulled_result[1:len(self.star6_pulled_result)]) == 0:
                self.one_pull()
            gacha_data[i, 0] = self.total_pull_count
            gacha_data[i, 1] = self.star6_pulled
            gacha_data[i, 2] = self.star5_pulled
            for j in range(len(self.star6_pulled_result)):
                gacha_data[i, j + 5] = self.star6_pulled_result[j]
        return gacha_data


def Normalize(data):
    m = np.mean(data)
    mx = max(data)
    mn = min(data)
    return [(float(i) - m) / (mx - mn) for i in data]


def xlwt_save(data, path):
    f = xlwt.Workbook()  # 创建工作簿
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
    [h, l] = data.shape  # h为行数，l为列数
    for i in range(h):
        for j in range(l):
            sheet1.write(i, j, data[i, j])
    f.save(path)


if __name__ == "__main__":
    people_num = 100000
    a = GachaTest(up_rate=0.7)
    gacha_data1 = a.get1UpSimulation(people_num)
    gacha_data2 = a.get2UpSimulation(people_num)
    writer = pd.ExcelWriter('Save_Excel.xlsx')
    data1_df = pd.DataFrame(gacha_data1)
    data1_df.columns = ['总抽数', '6星个数', '5星个数']
    # data1_df.to_excel(writer, 'page_1', float_format='%.5f')
    data2_df = pd.DataFrame(gacha_data2)
    data2_df.columns = ['总抽数', '6星个数', '5星个数', '4星个数','3星个数','歪了的6星个数', 'UP6星角色1个数', 'UP6星角色2个数']
    # data2_df.to_excel(writer, 'page_2', float_format='%.5f')
    avg_get_one = np.mean(gacha_data1[:, 0])
    avg_get_two = np.mean(gacha_data2[:, 0])
    print("出1个UP平均抽数：", avg_get_one)
    print("双UP出全的平均抽数：", avg_get_two)
    bin_pull1 = np.bincount(gacha_data1[:, 0])[1:]
    data_p1_df = pd.DataFrame(bin_pull1)
    # data_p1_df.to_excel(writer, 'page_3', float_format='%.5f')
    bin_pull2 = np.bincount(gacha_data2[:, 0])[1:]
    data_p2_df = pd.DataFrame(bin_pull2)
    # data_p2_df.to_excel(writer, 'page_4', float_format='%.5f')
    # writer.save()

    # 绘图
    plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
    plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号
    #
    f1 = plt.figure()
    ax1 = f1.add_subplot(1, 1, 1)
    ax1.set_title('出一个限定的抽卡数分布情况')
    ax1.set_xlabel("抽卡次数")
    ax1.set_ylabel("分布情况")
    ax1.plot(bin_pull1)
    f1.savefig('1个限定.png')

    f2 = plt.figure()
    ax2 = f2.add_subplot(1, 1, 1)
    ax2.set_title('出两个限定的抽卡数分布情况')
    ax2.set_xlabel("抽卡次数")
    ax2.set_ylabel("分布情况")
    ax2.plot(bin_pull2)
    f2.savefig('2个限定.png')
    plt.show()
