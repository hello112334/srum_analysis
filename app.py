"""
SRUM分析工具
"""
import os
import subprocess
# import configparser
import yaml
import mplcursors

from datetime import datetime
import pandas as pd
import numpy as np
# tkinter
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from tkcalendar import DateEntry
# matplotlib
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
# from matplotlib.ticker import MultipleLocator

# win32
import win32api
import win32security
import psutil

# google drive
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.service_account import ServiceAccountCredentials

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
    GUI介面
    先取得最新SRUM檔案再顯示按鈕
    """

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.start_date = None
        self.end_date = None
        self.srum_date = formatted_today
        self.create_widgets()

    def create_widgets(self):
        """
            創建視窗元件
        """
        btn_width = 22

        # row 1 =====================================================================================================
        # 取得最新的 SRUM 檔案按鈕
        self.get_srum_button = ttk.Button(
            self, text="取得最新的 SRUM 檔案", command=self.get_srum_file, width=btn_width)
        self.get_srum_button.grid(row=0, column=0)

        self.select_srum_button = ttk.Button(
            self, text="選擇 SRUM 檔案", command=self.select_srum_file, width=btn_width, state=tk.DISABLED)
        self.select_srum_button.grid(row=0, column=1)

        # row 2 =====================================================================================================
        # 查詢電量狀態按鈕
        self.query_energy_button = ttk.Button(
            self, text="查詢電量狀態", command=self.query_energy_usage, width=btn_width, state=tk.DISABLED)
        self.query_energy_button.grid(row=1, column=0)

        # 查詢 CPU 使用率按鈕
        self.query_cpu_button = ttk.Button(
            self, text="查詢CPU使用率", command=self.query_cpu_usage, width=btn_width, state=tk.DISABLED)
        self.query_cpu_button.grid(row=1, column=1)

        # 查詢網路流量按鈕
        self.query_network_button = ttk.Button(
            self, text="查詢應用程式網路流量", command=lambda: self.query_network_usage(0), width=btn_width, state=tk.DISABLED)
        self.query_network_button.grid(row=1, column=2)

        # 查詢網路流量按鈕
        self.query_cpu_table_button = ttk.Button(
            self, text="查詢應用程式CPU時間", command=lambda: self.query_cpu_table(0), width=btn_width, state=tk.DISABLED)
        self.query_cpu_table_button.grid(row=1, column=3)

        # 偵測異常紀錄按鈕
        self.detect_anomaly_button = ttk.Button(
            self, text="偵測異常紀錄", command=lambda: [self.query_network_usage(1), self.query_cpu_table(1)], width=btn_width, state=tk.DISABLED)
        self.detect_anomaly_button.grid(row=1, column=4)

        # row 3 =====================================================================================================
        # 開始日期 Label 和 Calendar
        self.start_date_label = ttk.Label(self, text="開始日期：")
        self.start_date_label.grid(row=2, column=0)
        self.start_cal = CustomDateEntry(self)
        self.start_cal.grid(row=2, column=1)

        # 開始/結束確定按鈕
        self.date_confirm_button = ttk.Button(
            self, text="確定", command=self.confirm_dates)
        self.date_confirm_button.grid(row=2, column=2, columnspan=1)

        # row 4 =====================================================================================================
        # 結束日期 Label 和 Calendar
        self.end_date_label = ttk.Label(self, text="結束日期：")
        self.end_date_label.grid(row=3, column=0)
        self.end_cal = CustomDateEntry(self)
        self.end_cal.grid(row=3, column=1)

        # 初期化
        self.start_date = self.start_cal.get_date()
        self.end_date = self.end_cal.get_date()

        # row 5 =====================================================================================================

    def select_srum_file(self):
        """note"""
        global file_path
        selected_file_path = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                            filetypes=(("Xlsx files", "*.xlsx"), ("All files", "*.*")))
        if selected_file_path:
            print(f"Seletted: {selected_file_path}")
            file_path = selected_file_path

    def confirm_dates(self):
        """note"""
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
        if not os.path.exists(file_path):
            messagebox.showerror("提醒", "今日的檔案尚未生成")
            subprocess.Popen(f'python {dir_path}/srum_dump2.py')  # 生成output檔案
        else:
            self.show_buttons()

            # 上傳檔案到 Google Drive
            upload_file(f"SRUM_DUMP_OUTPUT_{formatted_today}.xlsx", file_path)

    def show_buttons(self):
        """
            顯示按鈕
        """
        # 執行顯示按鈕
        self.select_srum_button.config(state=tk.NORMAL)

        self.query_energy_button.config(state=tk.NORMAL)
        self.query_cpu_button.config(state=tk.NORMAL)
        self.query_network_button.config(state=tk.NORMAL)
        self.query_cpu_table_button.config(state=tk.NORMAL)
        self.detect_anomaly_button.config(state=tk.NORMAL)

    def query_energy_usage(self):
        """
            查詢電量狀態
        """
        print("查詢電量狀態")
        # 在此添加查詢電量狀態的程式碼

        try:
            # 執行前處理
            pre_execute(file_path)

            # 讀取SRUM_DUMP_OUTPUT.xlsx檔案
            df = pd.read_excel(file_path, sheet_name='Energy Usage')

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
                print('此範圍區間沒有資料')
                raise Exception('此範圍區間沒有資料')

            # 繪圖
            _, ax = plt.subplots(num="電量狀態")
            ax.plot(
                df["Event Time Stamp"], df["DesignedCapacity"], label="DesignedCapacity", c="red", alpha=0.5)
            ax.plot(
                df["Event Time Stamp"], df["FullChargedCapacity"], label="FullChargedCapacity", c="blue", alpha=0.5)
            battery_line1, = ax.plot(
                df["Event Time Stamp"], df["Battery Level"], label="Battery Level", c="green", alpha=0.5)

            ax.set_xlabel('Event Time Stamp')
            ax.set_ylabel('Capacity / DesignedCapacity')

            # 設定圖例
            ax.legend(
                ['DesignedCapacity', 'FullChargedCapacity', 'Battery Level'])

            # 使用 mplcursors 套件顯示數據標籤
            cursor1 = mplcursors.cursor(battery_line1, hover=True)
            cursor1.connect("add", lambda sel: sel.annotation.set_text(
                f"Battery Level: {sel.target[1]:.2f}"))

            # 計算最佳刻度
            data_min = min(df['DesignedCapacity'].min(
            ), df['FullChargedCapacity'].min(), df['Battery Level'].min())
            data_max = max(df['DesignedCapacity'].max(
            ), df['FullChargedCapacity'].max(), df['Battery Level'].max())
            min_y, max_y, interval_y = calculate_ticks(data_min, data_max)

            # 設置x軸和y軸的最小值和最大值
            plt.ylim(min_y, max_y)

            # 設置主要刻度
            plt.gca().yaxis.set_major_locator(plt.MultipleLocator(interval_y))

            # 設置輔助刻度
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
            ax.grid(which='major', color='gray',
                    linestyle='-', linewidth=0.8)  # 主要
            ax.grid(which='minor', color='gray', linestyle='--',
                    linewidth=0.4, alpha=0.5)  # 輔助

            # 顯示圖表
            plt.show()

        except Exception as e:
            messagebox.showerror("錯誤", f"{e}")

    def query_cpu_usage(self):
        """
           查詢CPU使用率
        """
        print("查詢CPU使用率")
        # 在此添加查詢CPU使用率的程式碼

        try:
            # 執行前處理
            pre_execute(file_path)

            # 讀取Excel檔案中名為 "Application Resource Usage" 的資料表
            df = pd.read_excel(
                file_path, sheet_name="Application Resource Usage")
            df['Srum Entry Creation'] = pd.to_datetime(
                df['Srum Entry Creation'])  # 將開始和結束日期轉換為 datetime

            # App Timeline Provider
            df_timeline = pd.read_excel(
                file_path, sheet_name="App Timeline Provider")
            df_timeline['Srum Entry Creation'] = pd.to_datetime(
                df_timeline['Srum Entry Creation'])  # 將開始和結束日期轉換為 datetime

            # 將開始和結束日期轉換為 datetime
            start_date = pd.to_datetime(self.start_date)
            end_date = pd.to_datetime(self.end_date)

            # 使用布林索引來篩選 DataFrame
            df = df[(df['Srum Entry Creation'] >= start_date)
                    & (df['Srum Entry Creation'] <= end_date)]

            # 如果沒有資料，則顯示錯誤訊息
            if df.empty:
                raise Exception('此範圍區間沒有資料')

            # 1.計算應用程序的CPU時間消耗(秒)
            # unit: 0.0000001 秒
            # df["CPU time in Forground"] *= 0.0000001
            # df["CPU time in background"] *= 0.0000001

            # 繪製圖表
            _, ax = plt.subplots(num="CPU使用率")
            foreground_line, = ax.plot(
                df["Srum Entry Creation"], df["CPU time in Forground"], label="CPU time in Forground", c="red", alpha=0.5)
            _, = ax.plot(
                df["Srum Entry Creation"], df["CPU time in background"], label="CPU time in background", c="blue", alpha=0.5)

            ax.set_xlabel("Srum Entry Creation")
            ax.set_ylabel("CPU time (normalized)")
            ax.legend()

            # 使用 mplcursors 套件顯示數據標籤
            cursor1 = mplcursors.cursor(foreground_line, hover=True)

            @cursor1.connect("add")
            def on_add(sel):
                i = int(sel.index)
                cpu_front = df.iloc[i, 4]
                app_name = df.iloc[i, 2]  # 提取名字 Application

                # 顯示名字
                if app_name:
                    app_name = app_name.split('\\')[-1]  # 提取路徑最後一個斜線後的名字
                    sel.annotation.set_text(f"{app_name}\n{cpu_front:.2f}")
                else:
                    sel.annotation.set_text('-')

            # 計算最佳刻度
            min_y, max_y, interval_y = calculate_ticks(
                df['CPU time in Forground'].min(), df['CPU time in Forground'].max())

            # 設置x軸和y軸的最小值和最大值
            plt.ylim(min_y, max_y)

            # 設置主要刻度
            plt.gca().yaxis.set_major_locator(plt.MultipleLocator(interval_y))

            # 設置輔助刻度
            plt.gca().yaxis.set_minor_locator(plt.MultipleLocator(interval_y / 5))

            # 添加格線
            ax.grid(which='major', color='gray',
                    linestyle='-', linewidth=0.8)  # 主要
            ax.grid(which='minor', color='gray', linestyle='--',
                    linewidth=0.4, alpha=0.5)  # 輔助

            # 添加邊距
            plt.xticks(rotation=45)
            plt.tight_layout()

            def format_coord(x, y):
                x_date = mdates.num2date(x)
                return f"Time: {x_date.strftime('%Y-%m-%d %H:%M:%S')}, CPU: {y:.2f}"
            ax.format_coord = format_coord

            # 顯示圖表
            plt.show()

        except Exception as e:
            messagebox.showerror("錯誤", f"{e}")

    def query_network_usage(self, type):
        """
            查詢網路流量
        """
        print(f"查詢網路流量 {type}")
        # 在此添加查詢網路流量的程式碼

        try:
            # 執行前處理
            pre_execute(file_path)

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
            
            # 偵測異常紀錄時篩選網路使用率
            bytes_Sent_less = 1*10**5
            bytes_Sent_high = 5*10**7
            bytes_Received_less = 1*10**5
            bytes_Received_high = 2*10**9

            if type == 0:
                print('偵測異常紀錄時篩選網路使用率')
            elif type == 1:
                df = df[(df['Bytes Sent'] >= bytes_Sent_high)
                        | (df['Bytes Received'] >= bytes_Received_high)]

            # 如果沒有資料，則顯示錯誤訊息
            if df.empty:
                raise Exception('此範圍區間沒有資料')

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
            tag_counts = {'low': 0, 'normal': 0, 'high': 0}
            for index, row in df.iterrows():
                if row['Bytes Sent'] >= bytes_Sent_high or row['Bytes Received'] >= bytes_Received_high:
                    table.insert(parent='', index='end', iid=index,
                                 text='', values=list(row), tags=('high'))
                    tag_counts['high'] += 1
                elif row['Bytes Sent'] >= bytes_Sent_less or row['Bytes Received'] >= bytes_Received_less:
                    table.insert(parent='', index='end', iid=index,
                                 text='', values=list(row), tags=('normal'))
                    tag_counts['normal'] += 1
                else:
                    table.insert(parent='', index='end', iid=index,
                                 text='', values=list(row), tags=('low'))
                    tag_counts['low'] += 1

            # 調整表格大小
            table.pack(fill='both', expand=True)

            # 設定不同的背景顏色
            table.tag_configure('low', background='#90EE90') # 綠
            table.tag_configure('normal', background='#FFFACD') # 黃
            table.tag_configure('high', background='#FF6347') # 紅

            # 顯示圓餅圖
            if type == 0:
                labels = tag_counts.keys()
                sizes = tag_counts.values()
                colors = ['#90EE90', '#FFFACD', '#FF6347']  # 綠、黃、紅

                _, ax1 = plt.subplots()
                ax1.pie(sizes, labels=labels, colors=colors, autopct='%1.2f%%',
                        startangle=90)
                ax1.axis('equal')  # 使圓餅圖比例相等
                plt.show()

        except Exception as e:
            messagebox.showerror("錯誤", f"{e}")

    def query_cpu_table(self, type):
        """
            查詢CPU使用率(表格)
        """
        print(f"查詢CPU使用率(表格) {type}")
        # 在此添加查詢CPU使用率(表格)的程式碼

        try:
            # 執行前處理
            pre_execute(file_path)

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
            
            # 偵測異常紀錄時篩選CPU使用率
            cpu_Forground_less = 1*10**9
            cpu_Forground_high = 4*10**12
            cpu_background_less = 1*10**6
            cpu_background_high = 4*10**11

            if type == 0:
                print("偵測異常紀錄時篩選CPU使用率")
            elif type == 1:
                df = df[(df['CPU time in Forground'] >= cpu_Forground_high)
                        | (df['CPU time in background'] >= cpu_background_high)]

            # 如果沒有資料，則顯示錯誤訊息
            if df.empty:
                raise Exception('此範圍區間沒有資料')

            # 建立 GUI
            root_cpu_usege = tk.Tk()
            root_cpu_usege.title('CPU使用率(表格)')

            # 建立表格
            table = ttk.Treeview(root_cpu_usege)

            # 設定捲軸
            vsb = ttk.Scrollbar(root_cpu_usege,
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
            tag_counts = {'low': 0, 'normal': 0, 'high': 0}
            for index, row in df.iterrows():
                if row['CPU time in Forground'] >= cpu_Forground_high or row['CPU time in background'] >= cpu_background_high:
                    table.insert(parent='', index='end', iid=index,
                                 text='', values=list(row), tags=('high'))
                    tag_counts['high'] += 1
                elif row['CPU time in Forground'] >= cpu_Forground_less or row['CPU time in background'] >= cpu_background_less:
                    table.insert(parent='', index='end', iid=index,
                                 text='', values=list(row), tags=('normal'))
                    tag_counts['normal'] += 1
                else:
                    table.insert(parent='', index='end', iid=index,
                                 text='', values=list(row), tags=('low'))
                    tag_counts['low'] += 1

            # 調整表格大小
            table.pack(fill='both', expand=True)

            # 設定不同的背景顏色
            table.tag_configure('low', background='#90EE90') # 綠
            table.tag_configure('normal', background='#FFFACD') # 黃
            table.tag_configure('high', background='#FF6347') # 紅

            # 顯示圓餅圖
            if type == 0:
                labels = tag_counts.keys()
                sizes = tag_counts.values()
                colors = ['#90EE90', '#FFFACD', '#FF6347']  # 綠、黃、紅

                _, ax1 = plt.subplots()
                ax1.pie(sizes, labels=labels, colors=colors, autopct='%1.2f%%',
                        startangle=90)
                ax1.axis('equal')  # 使圓餅圖比例相等
                plt.show()

        except Exception as e:
            messagebox.showerror("錯誤", f"{e}")


def pre_execute(get_file_path):
    """
    執行前的準備
    """
    # 關閉所有圖表
    plt.close()

    # 檢查檔案是否存在
    if not os.path.isfile(get_file_path):
        messagebox.showerror("錯誤", "檔案不存在")


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
            num_ticks = range_value / interval if interval != 0 else 0
            if 2 <= num_ticks <= 5:
                min_tick = np.floor(min_value / interval) * interval if interval != 0 else min_value
                max_tick = np.ceil(max_value / interval) * interval if interval != 0 else max_value

                # 檢查最大刻度與數據最大值的比例
                ratio = (max_tick - max_value) / max_value if max_value != 0 else 0
                if ratio <= 0.1:
                    max_tick += interval

                # 檢查最小刻度與數據最小值的比例
                ratio = (min_value - min_tick) / min_value if min_value != 0 else 0
                if ratio <= 0.1 and min_tick > 0:
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
    # 檢查各個 SID 項目
    if sid == 'S-1-5-18 ( Local System)':
        return 'Local System'
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

def upload_file(filename, filepath):
    """ 上傳檔案到 Google Drive """

    # 設定 Google Drive API 金鑰檔案路徑
    gauth = GoogleAuth()

    # Try to load saved client credentials
    if os.path.isfile("mycreds.txt"):
        gauth.LoadCredentialsFile("mycreds.txt")

    if gauth.credentials is None:
        gauth.LocalWebserverAuth()
        gauth.SaveCredentialsFile("mycreds.txt")
    elif gauth.access_token_expired:
        gauth.Refresh()
        gauth.SaveCredentialsFile("mycreds.txt")
    else:
        gauth.Authorize()
    
    # 建立 Google Drive 連線
    drive = GoogleDrive(gauth)
    FOLDER_ID = "1LBIoTDKmNjihOsLx52tr3Px-8rRGB1az" # 資料夾 ID
    
    # 取得檔案清單
    file_list = drive.ListFile({'q': "'{}' in parents".format(FOLDER_ID)}).GetList()
    
    # 檢查檔案是否已存在
    for file in file_list:
        if file['title'] == filename:
            print(f'檔案已存在 Google Drive: {filename}')
            break
    else:
        file = drive.CreateFile({'title': filename, "parents": [{"id": FOLDER_ID}]})
        file.SetContentFile(filepath)
        file.Upload()
        print(f"檔案已上傳至 Google Drive: {filename}")

# 執行程式
if __name__ == "__main__":
    
    try:
        print(f"{'-'*20} START {'-'*20}")

        # 獲取當前目錄路徑
        dir_path = os.path.dirname(__file__)
        print(f"當前目錄路徑: {dir_path}")

        # 獲取當前日期
        today = datetime.now()

        # 將日期格式化為 "yyyymmdd"
        formatted_today = today.strftime("%Y%m%d")

        # 讀取 config.ini 檔案
        config_path = os.path.join(dir_path, 'config.yaml')

        # 檢查檔案是否存在
        if not os.path.exists(config_path):
            # print("Error: config.ini 檔案不存在")
            # 處理檔案不存在的情況，例如拋出異常或設定默認值
            raise Exception("config.ini 檔案不存在")

        with open(config_path, 'r') as f:
            config = yaml.safe_load(f)

        # 讀取 [Settings] 區段的各個配置設定
        output_directory = config['settings']['output_directory']
        output_directory = os.path.join(dir_path, output_directory)

        # SRUM_DUMP_OUTPUT.xlex檔案路徑
        file_path = f"{output_directory}/SRUM_DUMP_OUTPUT_{formatted_today}.xlsx"

        # output.xlex儲存資料夾
        if not os.path.exists(output_directory):
            os.makedirs(output_directory)

        # 獲取 CPU 速度
        cpu_info = psutil.cpu_freq()
        cpu_speed = cpu_info.current*1000000  # CPU速度（週期/秒）ex:2300000000
        print(f"CPU速度：{cpu_speed} Hz")

        # 創建 GUI
        root = tk.Tk()
        root.geometry("720x100")
        root.title("系統資源監控")
        app = Application(master=root)

        # 執行程式
        app.mainloop()

    except Exception as err:
        print(f"[ERROR] {err}")

    finally:
        print(f"{'-'*20} END {'-'*20}")
