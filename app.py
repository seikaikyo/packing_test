import time
import pyautogui
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import threading
import datetime
import os

class BarcodeStressTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("包裝輸入程式壓力測試與自動輸入工具")
        self.root.geometry("500x600")

        self.is_running = False
        self.total_inputs = 0
        self.successful_inputs = 0
        self.failed_inputs = 0
        self.start_time = None
        self.end_time = None
        self.progress_file = "progress.txt"

        # 延遲時間設置
        tk.Label(root, text="延遲(秒):").pack(pady=5)
        self.delay_entry = tk.Entry(root)
        self.delay_entry.insert(0, "5")
        self.delay_entry.pack(pady=5)

        # 輸入內容設置
        tk.Label(root, text="工單號碼:").pack(pady=5)
        self.first_field_entry = tk.Entry(root)
        self.first_field_entry.insert(0, "預設工單號碼")
        self.first_field_entry.pack(pady=5)

        tk.Label(root, text="包裝工站:").pack(pady=5)
        self.second_field_entry = tk.Entry(root)
        self.second_field_entry.insert(0, "預設包裝工站")
        self.second_field_entry.pack(pady=5)

        tk.Label(root, text="包裝員工:").pack(pady=5)
        self.third_field_entry = tk.Entry(root)
        self.third_field_entry.insert(0, "預設包裝員工")
        self.third_field_entry.pack(pady=5)

        # 倒數計時顯示
        self.countdown_label = tk.Label(root, text="", font=("Helvetica", 18))
        self.countdown_label.pack(pady=10)

        # Excel 檔案選擇按鈕
        tk.Button(root, text="選擇 Excel 檔案", command=self.select_file).pack(pady=10)
        self.file_label = tk.Label(root, text="未選擇檔案")
        self.file_label.pack(pady=5)

        # 開始測試按鈕
        self.start_button = tk.Button(root, text="開始測試", command=self.start_test)
        self.start_button.pack(pady=10)

        # 中斷測試按鈕
        self.stop_button = tk.Button(root, text="中斷測試", command=self.stop_test)
        self.stop_button.pack(pady=10)

        self.status_label = tk.Label(root, text="狀態: 閒置")
        self.status_label.pack(pady=20)

    def select_file(self):
        file_path = filedialog.askopenfilename(title="選擇 Excel 檔案", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.file_label.config(text=file_path)
            self.data = pd.read_excel(file_path)
        else:
            self.file_label.config(text="未選擇檔案")

    def start_test(self):
        if not self.is_running:
            if hasattr(self, 'data'):
                self.interval = float(self.delay_entry.get())
                self.is_running = True
                self.start_time = datetime.datetime.now()
                self.status_label.config(text="狀態: 執行中")
                self.thread = threading.Thread(target=self.run_test)
                self.thread.start()
            else:
                messagebox.showwarning("警告", "請先選擇 Excel 檔案")

    def run_test(self):
        # 載入進度
        last_index = self.load_progress()

        # 顯示倒數計時
        for i in range(int(self.interval), 0, -1):
            self.countdown_label.config(text=f"倒數 {i} 秒")
            self.root.update()
            time.sleep(1)

        self.countdown_label.config(text="")

        try:
            # 輸入第一個欄位
            pyautogui.typewrite(self.first_field_entry.get())
            pyautogui.press('enter')
            pyautogui.press('tab')  # 添加 TAB 跳到下一個輸入框
            time.sleep(0.5)

            # 輸入第二個欄位
            pyautogui.typewrite(self.second_field_entry.get())
            pyautogui.press('enter')
            time.sleep(0.5)

            # 輸入第三個欄位
            pyautogui.typewrite(self.third_field_entry.get())
            pyautogui.press('enter')
            time.sleep(0.5)
        except Exception as e:
            messagebox.showwarning("警告", f"輸入前三個欄位時發生錯誤: {e}")
            self.is_running = False
            self.generate_report()
            return

        # 從 Excel 檔案中忽略第一行，從上次未完成的地方繼續輸入
        for index, row in self.data.iloc[last_index:].iterrows():
            if not self.is_running:
                break

            try:
                # 每 60 筆檢查一次第二和第三欄位是否有值
                if index > 0 and index % 60 == 0:
                    if not self.check_field_values():
                        self.fill_default_values()

                # 輸入從 Excel 讀取的產品序號到第四個欄位
                pyautogui.typewrite(str(row['serial_number']))  # 使用產品序號的欄位名稱“serial_number”
                pyautogui.press('enter')

                self.successful_inputs += 1
                self.save_progress(index + 1)  # 保存當前進度
            except Exception as e:
                self.failed_inputs += 1

            self.total_inputs += 1

        self.is_running = False
        self.end_time = datetime.datetime.now()
        self.status_label.config(text="狀態: 完成")
        self.generate_report()

    def stop_test(self):
        if self.is_running:
            self.is_running = False
            self.status_label.config(text="狀態: 已中斷")
            self.generate_report()

    def generate_report(self):
        if self.start_time and self.end_time:
            total_time = (self.end_time - self.start_time).total_seconds()
            avg_time_per_input = total_time / self.total_inputs if self.total_inputs > 0 else 0

            report = f"""
            壓力測試報告
            -------------------------
            測試開始時間: {self.start_time}
            測試結束時間: {self.end_time}
            總耗時: {total_time:.2f} 秒
            總輸入次數: {self.total_inputs}
            成功輸入次數: {self.successful_inputs}
            失敗輸入次數: {self.failed_inputs}
            每次輸入平均耗時: {avg_time_per_input:.2f} 秒
            """

            # 顯示報告
            messagebox.showinfo("測試報告", report)
        else:
            messagebox.showwarning("警告", "尚未完成測試，無法生成報告")

    def save_progress(self, index):
        with open(self.progress_file, "w") as f:
            f.write(str(index))

    def load_progress(self):
        if os.path.exists(self.progress_file):
            with open(self.progress_file, "r") as f:
                return int(f.read().strip())
        return 1  # 默認從第二行開始（忽略標題）

    def check_field_values(self):
        # 模擬檢查第二和第三欄位是否有值
        return bool(self.second_field_entry.get().strip()) and bool(self.third_field_entry.get().strip())

    def fill_default_values(self):
        # 如果第二或第三欄位為空，填入預設值
        if not self.second_field_entry.get().strip():
            pyautogui.typewrite(self.second_field_entry.get())
            pyautogui.press('enter')
            time.sleep(0.5)

        if not self.third_field_entry.get().strip():
            pyautogui.typewrite(self.third_field_entry.get())
            pyautogui.press('enter')
            time.sleep(0.5)

if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodeStressTestApp(root)
    root.mainloop()
