import ctypes  # An included library with Python install.
import sys
import os
import logging
import tkinter as tk
import tkinter.filedialog
import numpy as np
import matplotlib.pyplot as plt
import math

# from __future__ import print_function


global mypath
mypath = 'c:\\'


def qquit():
    # window.destroy()
    quit()


def open_imp_file(path):                                           # imp File
    parameters = []
    d_freq = []
    d_ohm = []
    d_deg = []

    f = open(path, 'r')
    s = f.read().split('\n')
    i_documents_id = path
    f.close()
    for i in range(25):
        # print(str(i) + ',' + s[i])
        if i < 5:
            a = s[i].split(',')
            parameters.append(float(a[1]))
        if 6 < i < 9:
            a = s[i].split(',')
            parameters.append(a[1].strip())
        if i == 9:
            a = s[i].split(',')
            parameters.append(float(a[1].replace("KHZ", "")))
        if i == 10:
            a = s[i].split(',')
            parameters.append(a[1].strip())
        if i == 11:
            a = s[i].split(',')
            parameters.append(a[1].strip())
        if i == 12:
            a = s[i].split(',')
            parameters.append(float(a[1].replace("pF", "")))
        if i == 13:
            a = s[i].split(',')
            parameters.append(float(a[1].replace("KHZ", "")))
            parameters.append(float(a[3].replace("OHM", "")))
        if i == 14:
            a = s[i].split(',')
            parameters.append(float(a[3].replace("DEG", "")))
        if 14 < i < 18:
            a = s[i].split(',')
            parameters.append(float(a[1].replace("KHZ", "")))
            parameters.append(float(a[3].replace("OHM", "")))
        if i == 19:
            a = s[i].split(',')
            parameters.append(float(a[1].replace("K31:", "")))
            parameters.append(float(a[2].replace("K33", "")))
            parameters.append(float(a[4]))
        if i == 20:
            a = s[i].split(',')
            parameters.append(float(a[1].replace("C1K:", "")))

    j = int((parameters[1] - parameters[0]) / parameters[7])
    for i in range(j + 1):
        # print(s[i + 25])
        a = s[i + 25].split(',')
        d_freq.append(float(a[0]))
        d_ohm.append(float(a[1]))
        d_deg.append(float(a[2]))
    # print(i_documents_id)
    return parameters, d_freq, d_ohm, d_deg


def message_box(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


def OpenFile():
    global mypath
    # print(mypath)
    name = tkinter.filedialog.askopenfilename(initialdir=mypath, filetypes=(
        ("imp", "*.imp"), ("All Files", "*.*")), title="Choose a file.")
    # print(name)
    mypath = os.path.dirname(name)                      # 改變路徑
    # Using try in case user types in unknown file or closes without choosing
    # a file.
    try:
        return name
    except BaseException:
        print("No file exists")


def handle_close(evt):
    print('Closed Figure!')


def imp_y_axis_coordinates(l, h):
    ds1 = [1, 10, 100, 1000, 10000, 100000, 1000000, 10000000, 100000000]
    ds2 = ['1', '10', '100', '1K', '10K', '100K', '1M', '10M', '100M']
    da1 = []
    da2 = []
    a1 = int(math.log10(l))
    a2 = int(math.log10(h)) + 1
    for i in range(a1, a2):
        print(ds1[i], ds2[i])
        da1.append(ds1[i])
        da2.append(ds2[i])
    return da1, da2


def hit_me():
    try:
        d1, d2, d3, d4 = open_imp_file(OpenFile())
        # Create some mock data
        t = d2
        data1 = d3
        data2 = d4

        fig, ax1 = plt.subplots()

        color = 'tab:red'                                               # 設定紅色
        ax1.set_xlabel('Frequency (KHz)')                               # 設定標籤文字
        ax1.set_xlim(d1[0], d1[1])
        new_ticks = np.linspace(
            d1[0], d1[1], (int((d1[1] - d1[0]) / d1[2] + 1)))                # 座標軸標籤數
        ax1.set_xticks(new_ticks)

        ax1.set_ylabel('Z(Ω)', color=color)                            # 標籤文字顏色
        ax1.tick_params(axis='y', labelcolor=color)                     # 座標標籤上色
        ax1.semilogy(t, data1, color=color)                             # 加線,及上色
        ax1.set_ylim(d1[3], d1[4])
        ds1, ds2 = imp_y_axis_coordinates(d1[3], d1[4])                 # y軸標 變文字
        plt.yticks(ds1, ds2)
        # ax1.set_xticklabels(['10', '100', '1K', '10k', '100k'])
        plt.annotate(                                                   # 標記
            str(int(d1[17])) + 'Ω@' + str(d1[16]) + 'KHz', xy=(
                d1[16], d1[17]), xycoords='data', xytext=(
                -60, +20), textcoords='offset points', arrowprops=dict(
                arrowstyle="->", connectionstyle="arc3,rad=.2"))
        plt.grid(which='both')                                          # 劃格線參數 which='minor' , which='both'

        ax2 = ax1.twinx()                                               # 插入第二軸於第一軸

        color = 'tab:blue'                                              # 設定藍色
        ax2.set_ylabel('θ(°)', color=color)                            # 設定標籤文字
        ax2.tick_params(
            axis='y',
            labelcolor=color)                                           # 座標標籤上色
        # 加線,及上色
        ax2.plot(t, data2, color=color)
        ax2.set_ylim(-90, 90)
        # 座標軸標籤數
        new_ticks = np.linspace(-90, 90, 13)
        ax2.set_yticks(new_ticks)
        fig.tight_layout()                                              # 調整左邊標籤位置
        # 設定上的邊界
        plt.subplots_adjust(top=0.91)
        plt.title(
            "Impedance/Phase Angle vs. Frequency:\n",
            loc='left')  # 圖標題,靠左
        plt.title('Right Title', loc='right')
        # plt.annotate(
        #     '123dB@40KHz',
        #     xy=(40, 0), arrowprops=dict(arrowstyle='-'), xytext=(38, 10))

        # plt.annotate('local max', xy=(0, 0), fontsize=15)
        # plt.grid(which="both")
        plt.show()
        return
    except Exception as e:
        logging.exception(e)                                        # 輸出錯誤,配合 import logging:
        message_box('警告', '沒有選擇檔案', 0)                      # 跳出對話方塊
        return


if __name__ == '__main__':

    # print(sys.argv)
    # input()
    window = tk.Tk()
    window.protocol('WM_DELETE_WINDOW', qquit)      # tk視窗關閉事件,呼叫qquit
    window.title('my window')
    window.geometry('300x100+10+10')
    b = tk.Button(window,
                  text='開啟檔案',                  # 显示在按钮上的文字
                  width=15, height=2,
                  command=hit_me)                   # 点击按钮式执行的命令
    b.pack(anchor='nw', side='left')                # 按钮位置
    c = tk.Button(window,                           # 關閉按鍵
                  text='Close',                     # 显示在按钮上的文字
                  width=15, height=2,
                  command=qquit)                    # 点击按钮式执行的命令
    c.pack(anchor='nw', side='left')                # 按钮位置

    # print(len(sys.argv))
    try:
        if len(sys.argv) > 1:
            a = 'path = ' + os.path.splitext(sys.argv[1])[-1]
            message_box('Your title', sys.argv[1] + ' , ' + a, 0)  # 跳出對話方塊
            # OpenFile()
            # merge(sys.argv[1])
        else:
            message_box('Your title', 'No file', 0)  # 跳出對話方塊
            window.mainloop()
            
    except Exception as e:
        logging.exception(e)                # 輸出錯誤,配合 import logging
        # merge(OpenFile())

    # message_box('Your title', 'open file', 0)  # 跳出對話方塊
