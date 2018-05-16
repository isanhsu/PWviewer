import ctypes  # An included library with Python install.
import sys
import os
import logging
import tkinter as tk
import tkinter.filedialog
import numpy as np
import matplotlib.pyplot as plt
# from __future__ import print_function


global mypath
mypath = 'c:\\'


def open_imp_file(path):
    parameters = []
    d_freq = []
    d_ohm = []
    d_deg = []

    f = open(path, 'r')
    s = f.read().split('\n')
    i_documents_id = path
    f.close()
    for i in range(25):
        print(str(i) + ',' + s[i])
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
        print(s[i + 25])
        a = s[i + 25].split(',')
        d_freq.append(float(a[0]))
        d_ohm.append(float(a[1]))
        d_deg.append(float(a[2]))
    print(i_documents_id)
    return parameters, d_freq, d_ohm, d_deg


def message_box(title, text, style):
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


def OpenFile():
    global mypath
    print(mypath)
    name = tkinter.filedialog.askopenfilename(initialdir=mypath, filetypes=(
        ("imp", "*.imp"), ("All Files", "*.*")), title="Choose a file.")
    print(name)
    mypath = os.path.dirname(name)                  # 改變路徑
    # Using try in case user types in unknown file or closes without choosing
    # a file.
    try:
        return name
    except BaseException:
        print("No file exists")


def handle_close(evt):
    print('Closed Figure!')


def hit_me():
    try:
        d1, d2, d3, d4 = open_imp_file(OpenFile())
        # Create some mock data
        t = d2
        data1 = d3
        data2 = d4

        fig, ax1 = plt.subplots()

        color = 'tab:red'  # 設定紅色
        ax1.set_xlabel('Frequency (KHz)')  # 設定標籤文字
        ax1.set_xlim(d1[0], d1[1])
        new_ticks = np.linspace(
            d1[0], d1[1], ((d1[1] - d1[0]) / d1[2] + 1))  # 座標軸標籤數
        ax1.set_xticks(new_ticks)

        ax1.set_ylabel('Z(Ω)', color=color)  # 標籤文字顏色
        ax1.tick_params(axis='y', labelcolor=color)  # 座標標籤上色
        ax1.semilogy(t, data1, color=color)  # 加線,及上色
        ax1.set_ylim(d1[3], d1[4])
        # ax1.set_xticks(range(5))
        # ax1.set_xticklabels(['10', '100', '1K', '10k', '100k'])

        plt.grid(which='both')                                      # 劃格線參數
        # plt.grid(which='minor')
        # plt.grid()

        ax2 = ax1.twinx()  # 插入第二軸於第一軸

        color = 'tab:blue'  # 設定藍色
        ax2.set_ylabel('θ(°)', color=color)  # 設定標籤文字
        ax2.tick_params(axis='y', labelcolor=color)  # 座標標籤上色
        ax2.plot(t, data2, color=color)  # 加線,及上色
        ax2.set_ylim(-90, 90)
        new_ticks = np.linspace(-90, 90, 13)  # 座標軸標籤數
        ax2.set_yticks(new_ticks)
        fig.tight_layout()  # 調整左邊標籤位置
        plt.subplots_adjust(top=0.9)  # 設定上的邊界
        plt.title("Impedance/Phase Angle vs. Frequency:")

        # plt.annotate('local max', xy=(0, 0), fontsize=15)
        # plt.grid(which="both")
        plt.show()
        return
    except BaseException:
        message_box('警告', '沒有選擇檔案', 0)  # 跳出對話方塊
        return


if __name__ == '__main__':

    # print(sys.argv)
    # input()
    window = tk.Tk()
    window.title('my window')
    window.geometry('300x100')
    b = tk.Button(window,
                  text='開啟檔案',  # 显示在按钮上的文字
                  width=15, height=2,
                  command=hit_me)  # 点击按钮式执行的命令
    b.pack(anchor='nw', side='left')  # 按钮位置
    print(len(sys.argv))
    try:
        if len(sys.argv) > 1:
            a = 'path = ' + os.path.splitext(sys.argv[1])[-1]
            # message_box('Your title', sys.argv[1] + ' , ' + a, 0)  # 跳出對話方塊
            # OpenFile()
            # merge(sys.argv[1])

    except Exception as e:
        logging.exception(e)                # 輸出錯誤,配合 import logging
        # merge(OpenFile())

    # message_box('Your title', 'open file', 0)  # 跳出對話方塊
    window.mainloop()
