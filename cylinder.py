import time
import pyautogui
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import threading
import datetime
import os
import paho.mqtt.client as mqtt

class MQTTBarcodeStressTestApp:
    def __init__(self, root):
        self.log_file = "debug_log_mqtt.txt"
        self.log("初始化 MQTTBarcodeStressTestApp")

        self.root = root
        self.root.title("濾筒 MQTT 包裝輸入程式壓力測試與自動輸入工具")
        self.root.geometry("500x600")

        self.is_running = False
        self.total_inputs = 0
        self.successful_inputs = 0
        self.failed_inputs = 0
        self.start_time = None
        self.end_time = None
        self.progress_file = "progress_mqtt.txt"
        self.mqtt_client = mqtt.Client()

        # MQTT 設定
        self.mqtt_host = "10.6.51.204"
        self.mqtt_port = 1883
        self.mqtt_topic = "RFID/Terminal/PC02"

        # 延遲時間設置
        tk.Label(root, text="延遲(秒):").pack(pady=5)
        self.delay_entry = tk.Entry(root)
        self.delay_entry.insert(0, "5")
        self.delay_entry.pack(pady=5)

        # 輸入內容設置
        tk.Label(root, text="工單號碼:").pack(pady=5)
        self.first_field_entry = tk.Entry(root)
        self.first_field_entry.insert(0, "1407")
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
        self.serial_interval_entry.insert(0, "1")
        self.serial_interval_entry.pack(pady=5)

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

        # 連接 MQTT
        self.connect_mqtt()

    def log(self, message):
        with open(self.log_file, "a") as log:
            log.write(f"{datetime.datetime.now()}: {message}\n")

    def connect_mqtt(self):
        try:
            self.mqtt_client.connect(self.mqtt_host, self.mqtt_port, 60)
            self.mqtt_client.loop_start()
            self.log("成功連接到 MQTT 伺服器")
        except Exception as e:
            self.log(f"連接 MQTT 伺服器失敗: {e}")
            messagebox.showwarning("警告", f"無法連接到 MQTT 伺服器: {e}")

    def select_file(self):
        self.log("選擇 Excel 檔案")
        file_path = filedialog.askopenfilename(title="選擇 Excel 檔案", filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.file_label.config(text=file_path)
            self.data = pd.read_excel(file_path)
            self.log(f"成功載入 Excel 檔案: {file_path}")
        else:
            self.file_label.config(text="未選擇檔案")
            self.log("未選擇任何 Excel 檔案")

    def start_test(self):
        self.log("開始測試")
        if not self.is_running:
            if hasattr(self, 'data'):
                self.interval = float(self.delay_entry.get())
                self.serial_interval = float(self.serial_interval_entry.get())
                self.is_running = True
                self.start_time = datetime.datetime.now()
                self.status_label.config(text="狀態: 執行中")
                self.thread = threading.Thread(target=self.run_test)
                self.thread.start()
                self.log("測試執行緒已啟動")
            else:
                messagebox.showwarning("警告", "請先選擇 Excel 檔案")
                self.log("測試無法開始，因為沒有選擇 Excel 檔案")

    def run_test(self):
        self.log("進入 run_test 函數")

        # 載入進度
        last_index = self.load_progress()
        self.log(f"進度載入完成，從第 {last_index} 行開始")

        # 顯示倒數計時
        for i in range(int(self.interval), 0, -1):
            self.countdown_label.config(text=f"倒數 {i} 秒")
            self.root.update()
            time.sleep(1)

        self.countdown_label.config(text="")

        try:
            # 透過 MQTT 送出第一個欄位
            self.send_mqtt_message(self.first_field_entry.get())
            time.sleep(1)

            # 透過 MQTT 送出第二個欄位
            self.send_mqtt_message(self.second_field_entry.get())
            time.sleep(1)

            # 透過 MQTT 送出第三個欄位
            self.send_mqtt_message(self.third_field_entry.get())
            time.sleep(1)
            self.log("前三個欄位送出完成")
        except Exception as e:
            self.log(f"送出前三個欄位時發生錯誤: {e}")
            messagebox.showwarning("警告", f"送出前三個欄位時發生錯誤: {e}")
            self.is_running = False
            self.generate_report()
            return

        # 從 Excel 檔案中忽略第一行，從上次未完成的地方繼續送出資料
        for index, row in self.data.iloc[last_index:].iterrows():
            if not self.is_running:
                break

            try:
                # 每 60 筆檢查一次第二和第三欄位是否有值
                if index > 0 and index % 60 == 0:
                    self.log(f"檢查第 {index} 筆後的欄位值")
                    self.check_and_fill_fields()

                # 透過 MQTT 送出從 Excel 讀取的產品序號到第四個欄位
                self.send_mqtt_message(str(row['serial_number']))  # 使用產品序號的欄位名稱“serial_number”
                time.sleep(self.serial_interval)

                self.successful_inputs += 1
                self.save_progress(index + 1)  # 保存當前進度
            except Exception as e:
                self.log(f"在第 {index} 筆送出時發生錯誤: {e}")
                self.failed_inputs += 1

            self.total_inputs += 1
            self.log(f"成功送出第 {index + 1} 筆")

        self.is_running = False
        self.end_time = datetime.datetime.now()
        self.status_label.config(text="狀態: 完成")
        self.log("測試完成")
        self.generate_report()

    def send_mqtt_message(self, message):
        try:
            self.mqtt_client.publish(self.mqtt_topic, message)
            self.log(f"已送出 MQTT 訊息: {message}")
        except Exception as e:
            self.log(f"MQTT 訊息送出失敗: {e}")

    def stop_test(self):
        if self.is_running:
            self.is_running = False
            self.status_label.config(text="狀態: 已中斷")
            self.log("測試中斷")
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
            總送出次數: {self.total_inputs}
            成功送出次數: {self.successful_inputs}
            失敗送出次數: {self.failed_inputs}
            每次送出平均耗時: {avg_time_per_input:.2f} 秒
            """

            self.log(f"生成報告: {report}")

            # 顯示報告
            messagebox.showinfo("測試報告", report)
        else:
            messagebox.showwarning("警告", "尚未完成測試，無法生成報告")
            self.log("測試尚未完成，無法生成報告")

    def save_progress(self, index):
        with open(self.progress_file, "w") as f:
            f.write(str(index))
        self.log(f"保存進度: {index}")

    def load_progress(self):
        if os.path.exists(self.progress_file):
            with open(self.progress_file, "r") as f:
                progress = int(f.read().strip())
                self.log(f"載入進度: {progress}")
                return progress
        self.log("無進度檔案，從頭開始")
        return 1  # 默認從第二行開始（忽略標題）

    def check_field_values(self):
        # 檢查第二和第三欄位是否有值
        if not self.second_field_entry.get().strip():
            return False
        if not self.third_field_entry.get().strip():
            return False
        return True

    def fill_default_values(self):
        # 如果第二或第三欄位為空，直接在 GUI 中填入預設值
        if not self.second_field_entry.get().strip():
            self.second_field_entry.delete(0, tk.END)
            self.second_field_entry.insert(0, "AUTO")
            self.log("填入預設包裝工站: AUTO")

        if not self.third_field_entry.get().strip():
            self.third_field_entry.delete(0, tk.END)
            self.third_field_entry.insert(0, "YS00472")
            self.log("填入預設包裝員工: YS00472")

    def check_and_fill_fields(self):
        # 檢查並填入預設值，直到欄位正確為止
        while not self.check_field_values():
            self.fill_default_values()
            time.sleep(0.5)  # 等待更新
            self.log("重新檢查欄位值")

if __name__ == "__main__":
    root = tk.Tk()
    app = MQTTBarcodeStressTestApp(root)
    root.mainloop()

