import time
import pyautogui
import tkinter as tk
from tkinter import messagebox
import threading
import datetime

class BarcodeStressTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("濾筒 包裝輸入程式壓力測試與自動輸入工具")
        self.root.geometry("500x850")

        self.is_running = False
        self.total_inputs = 0
        self.successful_inputs = 0
        self.failed_inputs = 0
        self.start_time = None
        self.end_time = None

        # 延遲時間設置
        tk.Label(root, text="延遲(秒):").pack(pady=5)
        self.delay_entry = tk.Entry(root)
        self.delay_entry.insert(0, "5")
        self.delay_entry.pack(pady=5)

        # 輸入內容設置
        tk.Label(root, text="工單號碼:").pack(pady=5)
        self.first_field_entry = tk.Entry(root)
        self.first_field_entry.insert(0, "1457")
        self.first_field_entry.pack(pady=5)

        tk.Label(root, text="包裝工站:").pack(pady=5)
        self.second_field_entry = tk.Entry(root)
        self.second_field_entry.insert(0, "AUTO")
        self.second_field_entry.pack(pady=5)

        tk.Label(root, text="包裝員工:").pack(pady=5)
        self.third_field_entry = tk.Entry(root)
        self.third_field_entry.insert(0, "YS00472")
        self.third_field_entry.pack(pady=5)

        # 產品序號輸入間隔
        tk.Label(root, text="產品序號輸入間隔(秒):").pack(pady=5)
        self.serial_interval_entry = tk.Entry(root)
        self.serial_interval_entry.insert(0, "0.5")
        self.serial_interval_entry.pack(pady=5)

        # 成棧數量
        tk.Label(root, text="成棧數量:").pack(pady=5)
        self.items_per_pallet_entry = tk.Entry(root)
        self.items_per_pallet_entry.insert(0, "60")
        self.items_per_pallet_entry.pack(pady=5)
        self.items_per_pallet_entry.bind("<KeyRelease>", self.update_total_serials)

        # 棧板數量
        tk.Label(root, text="棧板數量:").pack(pady=5)
        self.pallet_count_entry = tk.Entry(root)
        self.pallet_count_entry.insert(0, "10")
        self.pallet_count_entry.pack(pady=5)
        self.pallet_count_entry.bind("<KeyRelease>", self.update_total_serials)

        # 產品序號編碼邏輯
        tk.Label(root, text="產品序號編碼邏輯:").pack(pady=5)
        self.serial_logic_entry = tk.Entry(root)
        self.serial_logic_entry.insert(0, "Y001-XXXX")  # 固定公司代號和廠商代號
        self.serial_logic_entry.pack(pady=5)

        # 顯示產品序號總數
        self.total_serials_label = tk.Label(root, text="總序號數: 0", font=("Helvetica", 14))
        self.total_serials_label.pack(pady=10)

        # 倒數計時顯示
        self.countdown_label = tk.Label(root, text="", font=("Helvetica", 18))
        self.countdown_label.pack(pady=10)

        # 開始測試按鈕
        self.start_button = tk.Button(root, text="開始測試", command=self.start_test)
        self.start_button.pack(pady=10)

        # 中斷測試按鈕
        self.stop_button = tk.Button(root, text="中斷測試", command=self.stop_test)
        self.stop_button.pack(pady=10)

        self.status_label = tk.Label(root, text="狀態: 閒置")
        self.status_label.pack(pady=20)

        # 初次計算總序號數
        self.update_total_serials()

    def update_total_serials(self, event=None):
        try:
            items_per_pallet = int(self.items_per_pallet_entry.get())
            pallet_count = int(self.pallet_count_entry.get())
            self.total_serials = items_per_pallet * pallet_count
            self.total_serials_label.config(text=f"總序號數: {self.total_serials}")
        except ValueError:
            self.total_serials_label.config(text="總序號數: 0")

    def start_test(self):
        if not self.is_running:
            self.is_running = True
            self.total_inputs = 0
            self.successful_inputs = 0
            self.failed_inputs = 0
            self.start_time = datetime.datetime.now()

            # 確認計算總數
            self.update_total_serials()

            self.status_label.config(text="狀態: 執行中")
            self.thread = threading.Thread(target=self.run_test)
            self.thread.start()

    def generate_serial_number(self, index):
        base_logic = self.serial_logic_entry.get().split("-")[0]
        sequence_number = f"{index + 1:04d}"
        return f"{base_logic}-{sequence_number}"

    def run_test(self):
        # 顯示倒數計時
        interval = float(self.delay_entry.get())
        for i in range(int(interval), 0, -1):
            self.countdown_label.config(text=f"倒數 {i} 秒")
            self.root.update()
            time.sleep(1)

        self.countdown_label.config(text="")

        try:
            # GUI 輸入前三個欄位
            pyautogui.typewrite(self.first_field_entry.get())
            pyautogui.press('enter')
            pyautogui.typewrite(self.second_field_entry.get())
            pyautogui.press('enter')
            pyautogui.typewrite(self.third_field_entry.get())
            pyautogui.press('enter')
        except Exception as e:
            messagebox.showwarning("警告", f"GUI 輸入前三個欄位時發生錯誤: {e}")
            self.is_running = False
            self.generate_report()
            return

        # 依次輸入序號
        serial_interval = float(self.serial_interval_entry.get())
        for i in range(self.total_serials):
            if not self.is_running:
                break

            try:
                serial_number = self.generate_serial_number(i)  # 生成產品序號
                pyautogui.typewrite(serial_number)
                pyautogui.press('enter')
                time.sleep(serial_interval)

                self.successful_inputs += 1
            except Exception as e:
                self.failed_inputs += 1
                continue

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

if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodeStressTestApp(root)
    root.mainloop()
