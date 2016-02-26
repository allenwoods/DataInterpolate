# -*- coding: utf-8 -*-
"""
Created on Thu Feb 25 21:57:10 2016

@author: Enigma
"""

import pandas as pd
import numpy as np
from scipy.interpolate import interp1d
import matplotlib as mpl
import matplotlib.pyplot as plt
import os


class CTbonds:
    def __init__(self, filename):
        self.data = pd.read_excel(filename, sheetname="整合").T  # 读取Excel文件并转置
        self.int_data = pd.DataFrame()  # 插值后数据
        self.keys, self.axislen, self.x = self.gen_x()
        self.newx = np.NaN
        self.cwd = os.getcwd()

    def gen_x(self):
        keys = self.data.keys().tolist()
        # 由于X轴[0年,1个月...]非数值, 此处生成等差序列代替原X轴
        axislen = len(self.data[keys[0]])
        x = np.linspace(0, 1, axislen)
        return keys, axislen, x

    def interpolate(self, length, method='cubic'):
        """
        使用每一评级数据生成插值函数, 默认使用三次样条插值,
        更多插值方法参见scipy文档:http://docs.scipy.org/doc/scipy/reference/tutorial/interpolate.html
        """
        Y = [interp1d(self.x, self.data[k].tolist(), kind=method) for k in self.keys]
        # 变量length表示插值长度, 如length=3, 原数据有10个点, 则从第二个节点开始, 每个点前增加length个空节点
        self.newx = np.linspace(0, 1, self.axislen + length * (self.axislen - 1))
        self.int_data = self.__append_cols__(self.data, length)
        k_count = 0
        for k in self.keys:
            self.int_data[k] = Y[k_count](self.newx)
            k_count += 1

    @staticmethod
    def __append_cols__(df, length, transpose=True):
        """
        拓展原数据序列长度
        """
        if transpose:
            df = df.T
        cols = df.columns.tolist()
        new_cols = [cols[0]]
        for item in cols[1:]:
            new_cols.extend(["" for i in range(length)] + [item])  # 从第二列开始, 每个时间点前增加length个空节点
        return df.reindex(columns=new_cols).T

    def plot(self, output_dir="", interpolate=True, length=3, method='cubic', export=True):
        """
        plot中可指定:
        插值: interpolate=True
        插值步长:length=3
        插值方法: method='cubic'
        具体参见类内方法interpolate
        """
        mpl.style.use('fivethirtyeight')  # 设定配色方案, 使用mpl.style.available查看可选方案
        from matplotlib.font_manager import FontProperties
        font = FontProperties(fname=r"c:\windows\fonts\msyh.ttc", size=14)  # 指定中文字体
        mpl.rcParams['axes.unicode_minus'] = False  # 解决保存图像是负号'-'显示为方块的问题
        plt.figure(figsize=(10, 6), dpi=320, facecolor=None)
        plt.title(u'城投债', fontproperties=font)  # 指定字体
        plt.xlabel(u'期限', fontproperties=font)
        plt.ylabel(u'收益率', fontproperties=font)
        plt.xticks(self.x, self.data.index, fontproperties=font, rotation=45)
        plt.grid(b=False, axis='x')
        if interpolate:
            self.interpolate(length, method)
            plt.xticks(self.newx, self.int_data.index, fontproperties=font, rotation=45)
            plt.plot(self.x, self.data, 'o', self.newx, self.int_data, '-', linewidth=2)
            plt.legend(self.keys, loc='best')
        else:
            plt.plot(self.x, self.data, linewidth=2)
            plt.legend(self.keys, fontproperties=font)
        plt.tight_layout()
        if len(output_dir) == 0:
            output_dir = self.cwd
        plt.savefig(os.path.join(output_dir, 'fig.png'))
        # plt.show()


if __name__ == '__main__':
    cwd = os.getcwd()
    excel_file = os.path.join(cwd, "CTbonds.xls")
    bonds = CTbonds(excel_file)
    bonds.plot()
    int_file = os.path.join(cwd, "interpolate.csv")
    bonds.int_data.to_csv(int_file)  # 储存插值后数据
