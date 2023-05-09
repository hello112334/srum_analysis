"""
SRUM分析工具
"""
import os
import subprocess
import configparser

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime
import pandas as pd
import numpy as np

import matplotlib.pyplot as plt
# import matplotlib.backends.backend_tkagg as tkagg
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
from matplotlib.ticker import MultipleLocator

from tkcalendar import DateEntry
import mplcursors

# win32
import win32api
import win32security
import psutil

class CustomToolbar(NavigationToolbar2Tk):
    """
    自定義工具列
    """

    def __init__(self, canvas, parent):
        """initialize the parent class"""
        NavigationToolbar2Tk.__init__(self, canvas, parent)

    def pack(self, *args, **kwargs):
        """pack pass"""

    def grid(self, *args, **kwargs):
        """grid pass"""

    def _update_cursor(self, event):
        """_update_cursor"""
        if event.inaxes and event.inaxes.get_navigate():
            x, y = event.xdata, event.ydata
            self.set_message(f"X: {x:.2f}, Y: {y:.2f}")
        else:
            self.set_message("")


class CustomDateEntry(DateEntry):
    """自定義日期選擇器"""

    def __init__(self, master=None, **kw):
        DateEntry.__init__(self, master=master,
                           date_pattern='yyyy年mm月dd日', **kw, width=15)

class Application(tk.Frame):
    """
    GUI介面先取得最新SRUM檔案再顯示按鈕
    """

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.start_date = None
        self.end_date = None
        self.create_widgets()

    def create_widgets(self):
        """
            創建視窗元件
        """
        # 取得最新的 SRUM 檔案按鈕
        self.get_srum_button = ttk.Button(
            self, text="取得最新的 SRUM 檔案", command=self.get_srum_file, state=tk.DISABLED)
        self.get_srum_button.grid(row=0, column=0)

        # 查詢電量狀態按鈕
        self.query_energy_button = ttk.Button(
            self, text="查詢電量狀態", command=self.query_energy_usage, state=tk.DISABLED)
        self.query_energy_button.grid(row=0, column=1)

        # 查詢 CPU 使用率按鈕
        self.query_cpu_button = ttk.Button(
            self, text="查詢CPU使用率", command=self.query_cpu_usage, state=tk.DISABLED)
        self.query_cpu_button.grid(row=0, column=2)

        # 查詢網路流量按鈕
        self.query_network_button = ttk.Button(
            self, text="查詢應用程式網路流量", command=self.query_network_usage, state=tk.DISABLED)
        self.query_network_button.grid(row=0, column=3)

        # 查詢網路流量按鈕
        self.query_cpu_table_button = ttk.Button(
            self, text="查詢應用程式CPU時間", command=self.query_cpu_table, state=tk.DISABLED)
        self.query_cpu_table_button.grid(row=0, column=4)

        # 偵測異常紀錄按鈕
        # self.detect_anomaly_button = ttk.Button(
        #     self, text="偵測異常紀錄", command=self.detect_anomaly)
        # self.detect_anomaly_button.grid(row=0, column=4)

        # 開始日期 Label 和 Calendar
        self.start_date_label = ttk.Label(self, text="開始日期：")
        self.start_date_label.grid(row=1, column=0)
        self.start_cal = CustomDateEntry(self)
        # self.start_cal = DateEntry(self)
        self.start_cal.grid(row=1, column=1)

        # 結束日期 Label 和 Calendar
        self.end_date_label = ttk.Label(self, text="結束日期：")
        self.end_date_label.grid(row=2, column=0)
        self.end_cal = CustomDateEntry(self)
        # self.end_cal = DateEntry(self)
        self.end_cal.grid(row=2, column=1)

        # 確定按鈕
        self.confirm_button = ttk.Button(
            self, text="確定", command=self.confirm_dates)
        self.confirm_button.grid(row=2, column=2, columnspan=2)

    def confirm_dates(self):
        """note"""
        # self.srum_btn_flg = False
        self.start_date = self.start_cal.get_date()
        self.end_date = self.end_cal.get_date()

        # 顯示按鈕
        if not self.start_date is None and not self.end_date is None:
            print("開始日期:", self.start_date)
            print("結束日期:", self.end_date)
            self.get_srum_button.config(state=tk.NORMAL)

    def get_srum_file(self):
        """note"""
        print("取得最新的 SRUM 檔案")
        # 在此添加取得最新的 SRUM 檔案的程式碼

        # 執行完畢後顯示按鈕
        if self.start_date is None or self.end_date is None:
            messagebox.showerror("錯誤", "日期尚未選擇")
        elif not os.path.exists(file_path):
            messagebox.showerror("提醒", "今日的檔案尚未生成")
            subprocess.Popen('python srum_dump2.py')
        else:
            # 執行 srum_dump2.exe
            self.show_buttons()

    def show_buttons(self):
        """
            顯示按鈕
        """
        # 執行顯示按鈕
        self.query_energy_button.config(state=tk.NORMAL)
        self.query_cpu_button.config(state=tk.NORMAL)
        self.query_network_button.config(state=tk.NORMAL)
        self.query_cpu_table_button.config(state=tk.NORMAL)
        # self.detect_anomaly_button.config(state=tk.NORMAL)

    def query_energy_usage(self):
        """
            查詢電量狀態
        """
        print("查詢電量狀態")
        # 在此添加查詢電量狀態的程式碼

        try:
            plt.close()

            # 讀取SRUM_DUMP_OUTPUT.xlsx檔案
            df = pd.read_excel(file_path, sheet_name='Energy Usage')
            print(df['Event Time Stamp'].describe())

            # 取出需要的欄位
            df = df[['Event Time Stamp', 'DesignedCapacity',
                    'FullChargedCapacity', 'Battery Level']]

            # 將 '日期' 列轉換為 datetime
            df['Event Time Stamp'] = pd.to_datetime(df['Event Time Stamp'])

            # 將開始和結束日期轉換為 datetime
            start_date = pd.to_datetime(self.start_date)
            end_date = pd.to_datetime(self.end_date)

            # 使用布林索引來篩選 DataFrame
            df = df[(df['Event Time Stamp'] >= start_date)
                    & (df['Event Time Stamp'] <= end_date)]

            # 如果沒有資料，則顯示錯誤訊息
            if df.empty:
                raise Exception

            print(start_date, end_date, df.describe())

            # 繪圖
            _, ax = plt.subplots(num="電量狀態")
            battery_line1, = ax.plot(
                df["Event Time Stamp"], df["DesignedCapacity"], label="DesignedCapacity", c="red", alpha=0.5)
            ax.plot(
                df["Event Time Stamp"], df["FullChargedCapacity"], label="FullChargedCapacity", c="blue", alpha=0.5)
            ax.plot(
                df["Event Time Stamp"], df["Battery Level"], label="Battery Level", c="green", alpha=0.5)

            ax.set_xlabel('Event Time Stamp')
            ax.set_ylabel('Capacity / DesignedCapacity')

            # 設定圖例
            ax.legend(
                ['DesignedCapacity', 'FullChargedCapacity', 'Battery Level'])

            # 使用 mplcursors 套件顯示數據標籤
            cursor1 = mplcursors.cursor(battery_line1, hover=True)
            cursor1.connect("add", lambda sel: sel.annotation.set_text(
                f"DesignedCapacity: {sel.target[1]:.2f}"))

            # 計算最佳刻度
            data_min = min(df['DesignedCapacity'].min(), df['FullChargedCapacity'].min(), df['Battery Level'].min())
            data_max = min(df['DesignedCapacity'].max(), df['FullChargedCapacity'].max(), df['Battery Level'].max())
            print(data_min, data_max)
            # min_x, max_x, interval_x = calculate_ticks(df['Event Time Stamp'].min().toordinal(), df['Event Time Stamp'].max().toordinal())
            min_y, max_y, interval_y = calculate_ticks(data_min, data_max)

            # 設置x軸和y軸的最小值和最大值
            # plt.xlim(min_x, max_x)
            plt.ylim(min_y, max_y)

            # 設置主要刻度
            # plt.gca().xaxis.set_major_locator(plt.MultipleLocator(interval_x))
            plt.gca().yaxis.set_major_locator(plt.MultipleLocator(interval_y))

            # 設置輔助刻度
            # plt.gca().xaxis.set_minor_locator(plt.MultipleLocator(interval_x / 5))
            plt.gca().yaxis.set_minor_locator(plt.MultipleLocator(interval_y / 5))

            # 添加邊距
            plt.xticks(rotation=45)
            plt.tight_layout()

            # 自定義座標顯示格式
            def format_coord(x, y):
                x_date = mdates.num2date(x)
                return f"Time: {x_date.strftime('%Y-%m-%d %H:%M:%S')}, Battery: {y:.2f}"
            ax.format_coord = format_coord

            # 添加格線
            ax.grid(which='major', color='gray', linestyle='-', linewidth=0.8) # 主要
            ax.grid(which='minor', color='gray', linestyle='--', linewidth=0.4, alpha=0.5) # 輔助

            # 顯示圖表
            plt.show()

        except Exception as err:
            messagebox.showerror("錯誤", f"發生錯誤: {err}")

    def query_cpu_usage(self):
        """
           查詢CPU使用率
        """
        print("查詢CPU使用率")
        # 在此添加查詢CPU使用率的程式碼

        try:
            # 讀取Excel檔案中名為 "Application Resource Usage" 的資料表
            df = pd.read_excel(
                file_path, sheet_name="Application Resource Usage")
            df['Srum Entry Creation'] = pd.to_datetime(
                df['Srum Entry Creation']) # 將開始和結束日期轉換為 datetime

            # App Timeline Provider
            df_timeline = pd.read_excel(
                file_path, sheet_name="App Timeline Provider")
            df_timeline['Srum Entry Creation'] = pd.to_datetime(
                df_timeline['Srum Entry Creation']) # 將開始和結束日期轉換為 datetime

            # 將開始和結束日期轉換為 datetime
            start_date = pd.to_datetime(self.start_date)
            end_date = pd.to_datetime(self.end_date)

            # 使用布林索引來篩選 DataFrame
            df = df[(df['Srum Entry Creation'] >= start_date)
                    & (df['Srum Entry Creation'] <= end_date)]

            # 如果沒有資料，則顯示錯誤訊息
            if df.empty:
                raise Exception

            # 1.計算應用程序的CPU時間消耗(秒)
            # unit: 0.0000001 秒
            df["CPU time in Forground"] *= 0.0000001
            df["CPU time in background"] *= 0.0000001

            # # 初始化實際執行時間和百分比列表
            # cpu_actual_runtime = []
            # cpu_usage = []

            # # 迭代 df1 的每一行
            # for app_name, row in df.iterrows():
            #     # 在 df2 中查找應用名稱對應的結束時間
            #     end_time = df_timeline.loc[app_name, 'EndTime'] if app_name in df_timeline.index else None

            #     # 計算實際執行時間
            #     if end_time is not None:
            #         runtime = (end_time - row['Srum Entry Creation']).seconds
            #     else:
            #         runtime = row['CPU time in Forground']

            #     cpu_actual_runtime.append(runtime)

            #     # 計算百分比
            #     perc = round(runtime / row['CPU執行時間'] * 100, 1)
            #     cpu_usage.append(perc)

            # # 將計算出的實際執行時間和百分比添加到 df1 中
            # df['cpu_actual_runtime'] = cpu_actual_runtime
            # df['cpu_usage'] = cpu_usage

            # 繪製圖表
            fig, ax = plt.subplots(num="CPU使用率")
            foreground_line, = ax.plot(
                df["Srum Entry Creation"], df["CPU time in Forground"], label="CPU time in Forground", c="red", alpha=0.5)
            background_line, = ax.plot(
                df["Srum Entry Creation"], df["CPU time in background"], label="CPU time in background", c="blue", alpha=0.5)
            # app_line, = ax.plot(
            #     df["Srum Entry Creation"], df["Application"], label="Application", c="green", alpha=0.5)
            ax.set_xlabel("Srum Entry Creation")
            ax.set_ylabel("CPU time (normalized)")
            ax.legend()

            # 使用 mplcursors 套件顯示數據標籤
            cursor1 = mplcursors.cursor(foreground_line, hover=True)

            @cursor1.connect("add")
            def on_add(sel):
                i = int(sel.target.index)
                cpu_front = df.iloc[i, 4]
                app_name = df.iloc[i, 2]  # 提取名字 Application

                # 顯示名字
                if app_name:
                    app_name = app_name.split('\\')[-1]  # 提取路徑最後一個斜線後的名字
                    sel.annotation.set_text(f"{app_name}\n{cpu_front:.2f}")
                else:
                    sel.annotation.set_text('-')

            # cursor1.connect("add", lambda sel: sel.annotation.set_text(
            #     f"Foreground: {sel.target[1]:.2f}"))
            # cursor2 = mplcursors.cursor(background_line, hover=True)
            # cursor2.connect("add", lambda sel: sel.annotation.set_text(
            #     f"Background: {sel.target[1]:.2f}"))

            # 計算最佳刻度
            # min_x, max_x, interval_x = calculate_ticks(df['Event Time Stamp'].min().toordinal(), df['Event Time Stamp'].max().toordinal())
            min_y, max_y, interval_y = calculate_ticks(df['CPU time in Forground'].min(), df['CPU time in Forground'].max())

            # 設置x軸和y軸的最小值和最大值
            # plt.xlim(min_x, max_x)
            plt.ylim(min_y, max_y)

            # 設置主要刻度
            # plt.gca().xaxis.set_major_locator(plt.MultipleLocator(interval_x))
            plt.gca().yaxis.set_major_locator(plt.MultipleLocator(interval_y))

            # 設置輔助刻度
            # plt.gca().xaxis.set_minor_locator(plt.MultipleLocator(interval_x / 5))
            plt.gca().yaxis.set_minor_locator(plt.MultipleLocator(interval_y / 5))

            # 添加格線
            ax.grid(which='major', color='gray', linestyle='-', linewidth=0.8) # 主要
            ax.grid(which='minor', color='gray', linestyle='--', linewidth=0.4, alpha=0.5) # 輔助

            # 添加邊距
            plt.xticks(rotation=45)
            plt.tight_layout()

            def format_coord(x, y):
                x_date = mdates.num2date(x)
                return f"Time: {x_date.strftime('%Y-%m-%d %H:%M:%S')}, CPU: {y:.2f}"
            ax.format_coord = format_coord

            # 顯示圖表
            plt.show()

        except:
            messagebox.showerror("錯誤", "這個範圍區間沒有資料")

    def query_network_usage(self):
        """
            查詢網路流量
        """
        print("查詢網路流量")
        # 在此添加查詢網路流量的程式碼
        try:
            # 讀取 Excel 檔案
            df = pd.read_excel(file_path, sheet_name='Network Data Usage')

            # 將 User SID 映射為對應的名稱
            df['User SID'] = df['User SID'].apply(map_user_sid)

            # 選擇需要的欄位
            df = df[['SRUM ENTRY CREATION', 'Application', 'User SID',
                    'Interface', 'Bytes Sent', 'Bytes Received']]

            # 將 '日期' 列轉換為 datetime
            df['SRUM ENTRY CREATION'] = pd.to_datetime(
                df['SRUM ENTRY CREATION'])

            # 將開始和結束日期轉換為 datetime
            start_date = pd.to_datetime(self.start_date)
            end_date = pd.to_datetime(self.end_date)

            # 使用布林索引來篩選 DataFrame
            df = df[(df['SRUM ENTRY CREATION'] >= start_date)
                    & (df['SRUM ENTRY CREATION'] <= end_date)]

            # 如果沒有資料，則顯示錯誤訊息
            if df.empty:
                raise Exception

            # 建立 GUI
            root_network_usege = tk.Tk()
            root_network_usege.title('網路流量')

            # 建立表格
            table = ttk.Treeview(root_network_usege)

            # 設定捲軸
            vsb = ttk.Scrollbar(root_network_usege,
                                orient="vertical", command=table.yview)
            vsb.pack(side='right', fill='y')
            table.configure(yscrollcommand=vsb.set)

            # 建立表格欄位
            table['columns'] = list(df.columns)

            # 設定欄位顯示名稱
            table.column('#0', width=0, stretch=tk.NO)
            for i, col in enumerate(df.columns):
                # 自動調整欄位寬度
                max_width = max([len(str(item)) for item in df[col]]) * 10
                max_width = min(max_width, 300)

                # App欄
                if i == 1:
                    table.column(col, width=max_width,
                                 minwidth=50, anchor=tk.W)
                else:
                    table.column(col, width=max_width,
                                 minwidth=50, anchor=tk.CENTER)
                table.heading(
                    col, text=col, command=lambda _col=col: sort_column(table, _col, False))

            # 將資料填入表格中
            for index, row in df.iterrows():
                table.insert(parent='', index='end', iid=index,
                             text='', values=list(row))

            # 調整表格大小
            table.pack(fill='both', expand=True)

            # 執行 GUI
            root_network_usege.mainloop()

        except:
            messagebox.showerror("錯誤", "這個範圍區間沒有資料")

    def query_cpu_table(self):
        """
            查詢CPU使用率(表格)
        """
        print("查詢CPU使用率(表格)")
        # 在此添加查詢CPU使用率(表格)的程式碼
        try:
            # 讀取 Excel 檔案
            df = pd.read_excel(
                file_path, sheet_name='Application Resource Usage')

            # 將 User SID 映射為對應的名稱
            df['User SID'] = df['User SID'].apply(map_user_sid)

            # 選擇需要的欄位
            df = df[['Srum Entry Creation', 'Application', 'User SID',
                    'CPU time in Forground', 'CPU time in background']]

            # 將 '日期' 列轉換為 datetime
            df['Srum Entry Creation'] = pd.to_datetime(
                df['Srum Entry Creation'])

            # 將開始和結束日期轉換為 datetime
            start_date = pd.to_datetime(self.start_date)
            end_date = pd.to_datetime(self.end_date)

            # 使用布林索引來篩選 DataFrame
            df = df[(df['Srum Entry Creation'] >= start_date)
                    & (df['Srum Entry Creation'] <= end_date)]

            # 如果沒有資料，則顯示錯誤訊息
            if df.empty:
                raise Exception

            # 建立 GUI
            root_network_usege = tk.Tk()
            root_network_usege.title('CPU使用率(表格)')

            # 建立表格
            table = ttk.Treeview(root_network_usege)

            # 設定捲軸
            vsb = ttk.Scrollbar(root_network_usege,
                                orient="vertical", command=table.yview)
            vsb.pack(side='right', fill='y')
            table.configure(yscrollcommand=vsb.set)

            # 建立表格欄位
            table['columns'] = list(df.columns)

            # 設定欄位顯示名稱
            table.column('#0', width=0, stretch=tk.NO)
            for i, col in enumerate(df.columns):
                # 自動調整欄位寬度
                max_width = max([len(str(item)) for item in df[col]]) * 10
                max_width = min(max_width, 300)

                # App欄
                if i == 1:
                    table.column(col, width=max_width,
                                 minwidth=50, anchor=tk.W)
                else:
                    table.column(col, width=max_width,
                                 minwidth=50, anchor=tk.CENTER)
                table.heading(
                    col, text=col, command=lambda _col=col: sort_column(table, _col, False))

            # 將資料填入表格中
            for index, row in df.iterrows():
                table.insert(parent='', index='end', iid=index,
                             text='', values=list(row))

            # 調整表格大小
            table.pack(fill='both', expand=True)

            # 執行 GUI
            root_network_usege.mainloop()

        except:
            messagebox.showerror("錯誤", "這個範圍區間沒有資料")

    def detect_anomaly(self):
        """
            偵測異常紀錄
        """
        print("偵測異常紀錄")
        # 在此添加偵測異常紀錄的程式碼

        try:
            # 讀取 Excel 檔案
            df = pd.read_excel(
                file_path, sheet_name="Application Resource Usage")

            # 篩選符合條件的資料
            filtered_df = df[df["CPU time in Forground"]/cpu_speed > 1000]
            filtered_df = filtered_df[[
                "Srum Entry Creation", "Application", "User SID", "CPU time in Forground"]]
            filtered_df = filtered_df.reset_index(drop=True)
            filtered_df["User SID"] = filtered_df["User SID"].apply(
                map_user_sid)

            # 創建新視窗，顯示篩選後的資料
            if len(filtered_df) == 0:
                # result_label.config(text="No data found")
                print("Error: Could not read file")
            else:
                # root = tk.Tk()
                new_window = tk.Tk()
                new_window.title("Filtered Data")

                # 創建 Treeview 元件
                tree = ttk.Treeview(new_window, selectmode='browse')
                tree.pack(fill='both', expand=True)

                # 設定欄位
                tree["columns"] = ["Srum Entry Creation",
                                   "Application", "User SID", "CPU time in Forground"]
                tree.column("#0", width=0, stretch=tk.NO)
                tree.column("Srum Entry Creation", width=150,
                            minwidth=50, anchor=tk.CENTER)
                tree.column("Application", width=150,
                            minwidth=50, anchor=tk.CENTER)
                tree.column("User SID", width=150,
                            minwidth=50, anchor=tk.CENTER)
                tree.column("CPU time in Forground", width=150,
                            minwidth=50, anchor=tk.CENTER)

                # 設定欄位標題
                tree.heading("Srum Entry Creation",
                             text="Srum Entry Creation", anchor=tk.CENTER)
                tree.heading("Application", text="Application",
                             anchor=tk.CENTER)
                tree.heading("User SID", text="User SID", anchor=tk.CENTER)
                tree.heading("CPU time in Forground/2300000000",
                             text="CPU time in Forground(S)", anchor=tk.CENTER)

                # 加入資料
                for i, row in filtered_df.iterrows():
                    tree.insert("", i, values=tuple(row))

            root.mainloop()

        except:
            messagebox.showerror("錯誤", "這個範圍區間沒有資料")

def calculate_ticks(min_value, max_value):
    """
    計算刻度
    """
    if min_value == max_value:
        if min_value == 0:
            min_value = -1
            max_value = 1
        else:
            min_value *= 0.9
            max_value *= 1.1

    range_value = max_value - min_value
    scales = [1, 2, 3, 5]
    scale = 1

    ck_time = 0
    while True:
        for s in scales:
            interval = scale * s
            num_ticks = range_value / interval
            if 2 <= num_ticks <= 5:
                min_tick = np.floor(min_value / interval) * interval
                max_tick = np.ceil(max_value / interval) * interval

                # 檢查最大刻度與數據最大值的比例
                ratio = (max_tick - max_value) / max_value
                if ratio <= 0.1:
                    max_tick += interval

                # 檢查最小刻度與數據最小值的比例
                ratio = (min_value - min_tick) / min_value
                if ratio <= 0.1:
                    min_tick -= interval

                return min_tick, max_tick, interval
        scale *= 10

        if ck_time > 100:
            break
        else:
            ck_time += 1

def sort_column(tree, col, reverse):
    """
    排序表格欄位
    """
    data_list = [(tree.set(child, col), child)
                 for child in tree.get_children('')]

    # 檢查是否為數字
    is_number = all([num_str.replace(".", "", 1).isdigit()
                    for num_str, _ in data_list])

    if is_number:
        data_list.sort(key=lambda t: float(t[0]), reverse=reverse)
    else:
        data_list.sort(key=lambda t: t[0], reverse=reverse)

    for index, (_, child) in enumerate(data_list):
        tree.move(child, '', index)

    tree.heading(col, command=lambda: sort_column(tree, col, not reverse))


def map_user_sid(sid):
    """
    將 User SID 映射為對應的名稱
    """
    # 檢查是否為空值
    if sid == 'S-1-5-18 ( Local System)':
        return 'Local System'
    # elif sid == 'S-1-5-21-2506843646-2876841158-3598199272-1001( unknown)':
    #     return 'user'
    elif sid == 'S-1-5-19 ( NT Authority)':
        return 'NT Authority'
    elif sid == 'S-1-5-20 ( NT Authority)':
        return 'NT Authority'
    elif get_user_name_from_sid(sid):
        # 使用函數獲取 SID 對應的用戶名稱
        return get_user_name_from_sid(sid)
    else:
        return sid


def get_user_name_from_sid(sid):
    """
    將 SID 映射為對應的用戶名稱
    """
    # 檢查是否為空值
    if sid is None:
        return None

    try:
        # 去除 SID 字串中的 "( unknown)"
        sid = sid.replace(" ", "")
        sid = sid.replace("(unknown)", "")
        # 獲取本地電腦名稱
        computer_name = win32api.GetComputerName()
        # 將 SID 字串轉換為 SID 對象
        sid_object = win32security.ConvertStringSidToSid(sid)
        # 獲取用戶名
        name, _, _ = win32security.LookupAccountSid(computer_name, sid_object)

        return name

    except:
        return None


# 執行程式
if __name__ == "__main__":

    # 獲取當前日期
    today = datetime.now()

    # 將日期格式化為 "yyyymmdd"
    formatted_today = today.strftime("%Y%m%d")

    # 讀取 config.ini 檔案
    config = configparser.ConfigParser()
    config.read('config.ini')

    # 讀取 [Settings] 區段的各個配置設定
    output_directory = config.get('Settings', 'output_directory')

    # SRUM_DUMP_OUTPUT.xlex檔案路徑
    file_path = f"{output_directory}/SRUM_DUMP_OUTPUT_{formatted_today}.xlsx"

    # output.xlex儲存資料夾
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)
    
    # 獲取 CPU 速度
    cpu_info = psutil.cpu_freq()
    cpu_speed = cpu_info.current*1000000 # CPU速度（週期/秒）ex:2300000000
    print(f"CPU速度：{cpu_speed} Hz")

    # 創建 GUI
    root = tk.Tk()
    root.geometry("600x100")
    root.title("系統資源監控")
    app = Application(master=root)

    # 執行程式
    app.mainloop()
