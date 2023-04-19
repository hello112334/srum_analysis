"""
SRUM分析工具
"""
import tkinter as tk
from tkinter import ttk
# from tkinter import filedialog
# from tkinter import Frame
from tkinter import messagebox
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
# import matplotlib.backends.backend_tkagg as tkagg
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
from tkcalendar import DateEntry
import mplcursors
# import os
import subprocess

# 獲取當前日期
today = datetime.now()

# 將日期格式化為 "yyyymmdd"
formatted_today = today.strftime("%Y%m%d")

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
            self, text="取得最新的 SRUM 檔案", command=self.get_srum_file)
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
            self, text="查詢網路流量", command=self.query_network_usage, state=tk.DISABLED)
        self.query_network_button.grid(row=0, column=3)

        # 偵測異常紀錄按鈕
        # self.detect_anomaly_button = ttk.Button(
        #     self, text="偵測異常紀錄", command=self.detect_anomaly)
        # self.detect_anomaly_button.grid(row=0, column=4)


        # 開始日期 Label 和 Calendar
        self.start_date_label = ttk.Label(self, text="開始日期：")
        self.start_date_label.grid(row=1, column=0)
        self.start_cal = DateEntry(self)
        self.start_cal.grid(row=1, column=1)

        # 結束日期 Label 和 Calendar
        self.end_date_label = ttk.Label(self, text="結束日期：")
        self.end_date_label.grid(row=2, column=0)
        self.end_cal = DateEntry(self)
        self.end_cal.grid(row=2, column=1)

        # 確定按鈕
        self.confirm_button = ttk.Button(
            self, text="確定", command=self.confirm_dates)
        self.confirm_button.grid(row=3, column=0, columnspan=2)

        # 圖表顯示區域
        # self.fig, self.ax = plt.subplots()
        # self.canvas = FigureCanvasTkAgg(self.fig, self)
        # self.canvas.get_tk_widget().grid(row=4, column=0, columnspan=5)

        # self.toolbar = CustomToolbar(self.canvas, self)
        # self.toolbar.grid(row=5, column=0, columnspan=5)

    def confirm_dates(self):
        """note"""
        self.start_date = self.start_cal.get_date()
        self.end_date = self.end_cal.get_date()
        print("開始日期:", self.start_date)
        print("結束日期:", self.end_date)


    def get_srum_file(self):
        """note"""
        print("取得最新的 SRUM 檔案")
        # 在此添加取得最新的 SRUM 檔案的程式碼

        # 執行 srum_dump2.exe
        subprocess.Popen('python srum_dump2.py')

        # 執行顯示按鈕
        self.query_energy_button.config(state=tk.NORMAL)
        self.query_cpu_button.config(state=tk.NORMAL)
        self.query_network_button.config(state=tk.NORMAL)
        # self.detect_anomaly_button.config(state=tk.NORMAL)

    def query_energy_usage(self):
        """
            查詢電量狀態
        """
        print("查詢電量狀態")
        # 在此添加查詢電量狀態的程式碼
        
        try:
                
            # 讀取SRUM_DUMP_OUTPUT.xlsx檔案
            df = pd.read_excel(f'SRUM_DUMP_OUTPUT_{formatted_today}.xlsx', sheet_name='Energy Usage')

            # 取出需要的欄位
            df = df[['Event Time Stamp', 'DesignedCapacity',
                    'FullChargedCapacity', 'Battery Level']]

            # 將 '日期' 列轉換為 datetime
            df['Event Time Stamp'] = pd.to_datetime(df['Event Time Stamp'])

            # 將開始和結束日期轉換為 datetime
            start_date = pd.to_datetime(self.start_date)
            end_date = pd.to_datetime(self.end_date)

            # 使用布林索引來篩選 DataFrame
            df = df[(df['Event Time Stamp'] >= start_date) & (df['Event Time Stamp'] <= end_date)]

            # 設定X軸、Y軸數據
            x_data = df['Event Time Stamp']
            y_data = df[['DesignedCapacity',
                        'FullChargedCapacity', 'Battery Level']]

            # 繪圖
            plt.plot(x_data, y_data)
            plt.legend(['DesignedCapacity', 'FullChargedCapacity', 'Battery Level'])
            plt.xlabel('Event Time Stamp')
            plt.xticks(rotation=45)
            plt.ylabel('Capacity / Battery Level')

            # 添加邊距
            plt.tight_layout()

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
            # 設定檔案路徑
            file_path = f"SRUM_DUMP_OUTPUT_{formatted_today}.xlsx"

            # 讀取Excel檔案中名為 "Application Resource Usage" 的資料表
            df = pd.read_excel(
                file_path, sheet_name="Application Resource Usage")

            # 將 '日期' 列轉換為 datetime
            df['Srum Entry Creation'] = pd.to_datetime(df['Srum Entry Creation'])

            # 將開始和結束日期轉換為 datetime
            start_date = pd.to_datetime(self.start_date)
            end_date = pd.to_datetime(self.end_date)

            # 使用布林索引來篩選 DataFrame
            df = df[(df['Srum Entry Creation'] >= start_date) & (df['Srum Entry Creation'] <= end_date)]

            # 將 "CPU time in Forground" 和 "CPU time in background" 欄位的值除以 230000000
            df["CPU time in Forground"] /= 230000000
            df["CPU time in background"] /= 230000000

            # 繪製圖表
            fig, ax = plt.subplots()
            foreground_line, = ax.plot(
                df["Srum Entry Creation"], df["CPU time in Forground"], label="CPU time in Forground")
            background_line, = ax.plot(
                df["Srum Entry Creation"], df["CPU time in background"], label="CPU time in background")
            ax.set_xlabel("Srum Entry Creation")
            ax.set_ylabel("CPU time (normalized)")
            ax.legend()

            # 使用 mplcursors 套件顯示數據標籤
            cursor1 = mplcursors.cursor(foreground_line, hover=True)
            cursor1.connect("add", lambda sel: sel.annotation.set_text(
                f"Foreground: {sel.target[1]:.2f}"))
            cursor2 = mplcursors.cursor(background_line, hover=True)
            cursor2.connect("add", lambda sel: sel.annotation.set_text(
                f"Background: {sel.target[1]:.2f}"))

            # 調整 X 軸標籤角度
            plt.xticks(rotation=45, ha='right')

            # 添加邊距
            plt.tight_layout()

            # 顯示圖表
            plt.show()

        except Exception as err:
            messagebox.showerror("錯誤", f"發生錯誤: {err}")

    def query_network_usage(self):
        """
            查詢網路流量
        """
        print("查詢網路流量")
        # 在此添加查詢網路流量的程式碼
        try:
            # 讀取 Excel 檔案
            df = pd.read_excel(f'SRUM_DUMP_OUTPUT_{formatted_today}.xlsx',
                            sheet_name='Network Data Usage')

            # 將 User SID 映射為對應的名稱
            df['User SID'] = df['User SID'].apply(map_user_sid)

            # 選擇需要的欄位
            df = df[['SRUM ENTRY CREATION', 'Application', 'User SID',
                    'Interface', 'Bytes Sent', 'Bytes Received']]


            # 將 '日期' 列轉換為 datetime
            df['SRUM ENTRY CREATION'] = pd.to_datetime(df['SRUM ENTRY CREATION'])

            # 將開始和結束日期轉換為 datetime
            start_date = pd.to_datetime(self.start_date)
            end_date = pd.to_datetime(self.end_date)

            # 使用布林索引來篩選 DataFrame
            df = df[(df['SRUM ENTRY CREATION'] >= start_date) & (df['SRUM ENTRY CREATION'] <= end_date)]

            # 建立 GUI
            root = tk.Tk()
            root.title('Network Data Usage')

            # 建立表格
            table = ttk.Treeview(root)

            # 設定捲軸
            vsb = ttk.Scrollbar(root, orient="vertical", command=table.yview)
            vsb.pack(side='right', fill='y')
            table.configure(yscrollcommand=vsb.set)

            # 建立表格欄位
            table['columns'] = list(df.columns)

            # 設定欄位顯示名稱
            table.column('#0', width=0, stretch=tk.NO)
            for i, col in enumerate(df.columns):
                table.column(col, width=100, minwidth=50, anchor=tk.CENTER)
                table.heading(col, text=col, command=lambda _col=col: sort_column(table, _col, False))

            # 將資料填入表格中
            for index, row in df.iterrows():
                table.insert(parent='', index='end', iid=index,
                            text='', values=list(row))

            # 調整表格大小
            table.pack(fill='both', expand=True)

            # 執行 GUI
            root.mainloop()

        except Exception as err:
            messagebox.showerror("錯誤", f"發生錯誤: {err}")

    def detect_anomaly(self):
        """
            偵測異常紀錄
        """
        print("偵測異常紀錄")
        # 在此添加偵測異常紀錄的程式碼

        try:
            # 讀取 Excel 檔案
            df = pd.read_excel(f'SRUM_DUMP_OUTPUT_{formatted_today}.xlsx', sheet_name="Application Resource Usage")

            # 篩選符合條件的資料
            filtered_df = df[df["CPU time in Forground"]/2300000000 > 1]
            filtered_df = filtered_df[[
                "Srum Entry Creation", "Application", "User SID", "CPU time in Forground"]]
            filtered_df = filtered_df.reset_index(drop=True)
            filtered_df["User SID"] = filtered_df["User SID"].apply(map_user_sid)

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
                tree.column("User SID", width=150, minwidth=50, anchor=tk.CENTER)
                tree.column("CPU time in Forground", width=150,
                            minwidth=50, anchor=tk.CENTER)

                # 設定欄位標題
                tree.heading("Srum Entry Creation",
                            text="Srum Entry Creation", anchor=tk.CENTER)
                tree.heading("Application", text="Application", anchor=tk.CENTER)
                tree.heading("User SID", text="User SID", anchor=tk.CENTER)
                tree.heading("CPU time in Forground/2300000000",
                            text="CPU time in Forground(S)", anchor=tk.CENTER)

                # 加入資料
                for i, row in filtered_df.iterrows():
                    tree.insert("", i, values=tuple(row))

            root.mainloop()

        except Exception as err:
            messagebox.showerror("錯誤", f"發生錯誤: {err}")

    # def plot_data(self, data, xlabel, ylabel, title):
    #     """note"""
    #     self.ax.clear()
    #     self.ax.plot(data)
    #     self.ax.set_xlabel(xlabel)
    #     self.ax.set_ylabel(ylabel)
    #     self.ax.set_title(title)
    #     self.canvas.draw()

    #     mplcursors.cursor(self.ax, hover=True)

# def sid_to_user(sid):
#     """
#         處理 User SID 的顯示
#     """
#     if sid == 'S-1-5-21-2506843646-2876841158-3598199272-1001( unknown)':
#         return "user"
#     elif sid == "S-1-5-18 ( Local System)":
#         return "Local System"
#     elif sid == "S-1-5-19 ( NT Authority)":
#         return "NT Authority"
#     elif sid == "S-1-5-19 ( NT Authority)":
#         return "NT Authority"
#     else:
#         return sid

def sort_column(tree, col, reverse):
    """
    排序表格欄位
    """
    data_list = [(tree.set(child, col), child) for child in tree.get_children('')]
    
    # 檢查是否為數字
    is_number = all([num_str.replace(".", "", 1).isdigit() for num_str, _ in data_list])

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
    if sid == 'S-1-5-18 ( Local System)':
        return 'Local System'
    elif sid == 'S-1-5-21-2506843646-2876841158-3598199272-1001( unknown)':
        return 'user'
    elif sid == 'S-1-5-19 ( NT Authority)':
        return 'NT Authority'
    elif sid == 'S-1-5-20 ( NT Authority)':
        return 'NT Authority'
    else:
        return sid

# 執行程式
if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("600x100")
    root.title("系統資源監控")
    app = Application(master=root)
    app.mainloop()
