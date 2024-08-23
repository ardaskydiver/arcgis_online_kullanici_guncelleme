import sys
import os
import pandas as pd
from arcgis.gis import GIS
import requests
import threading
import tkinter as tk
from tkinter import filedialog, ttk

class WorkerThread(threading.Thread):
    def __init__(self, excel_path, update_callback, progress_callback, stop_event):
        super().__init__()
        self.excel_path = excel_path
        self.update_callback = update_callback
        self.progress_callback = progress_callback
        self.stop_event = stop_event

    def run(self):
        df = pd.read_excel(self.excel_path, dtype={'kullanici_adi': str, 'sifre': str, 'yeni_sifre': str})
        total_rows = len(df)
        for index, row in df.iterrows():
            if self.stop_event.is_set():
                break

            kullanici_adi = row['kullanici_adi']
            sifre = row['sifre']
            yeni_sifre = row['yeni_sifre']

            try:
                gis = GIS("https://www.arcgis.com", kullanici_adi, sifre)
                self.update_callback(f"Başarılı giriş: {gis.properties.user.username}")
            except Exception:
                self.update_callback(f"Kullanıcı adı veya şifre hatalı: {kullanici_adi}")
                continue

            items = gis.content.search(query=f"owner:{kullanici_adi}", max_items=10000)
            for item in items:
                if item.protected:
                    item.protect(False)
                    self.update_callback(f"'{item.title}' için koruma kaldırıldı.")
                item.delete()
                self.update_callback(f"'{item.title}' başarıyla silindi.")

            gruplar = gis.groups.search(query="", max_groups=500)
            for grup in gruplar:
                if grup.owner != gis.users.me.username:
                    try:
                        grup.leave()
                        self.update_callback(f"'{grup.title}' grubundan başarıyla çıkış yapıldı.")
                    except Exception:
                        pass

            url = f"https://www.arcgis.com/sharing/rest/community/users/{kullanici_adi}/update"
            params = {
                'f': 'json',
                'token': gis._con.token,
                'password': yeni_sifre,
                'currentPassword': sifre
            }
            response = requests.post(url, data=params)
            response_json = response.json()
            if response.status_code == 200 and response_json.get('success'):
                self.update_callback(f"Şifre başarıyla değiştirildi.")
            else:
                self.update_callback(f"Şifre değiştirilemedi: {response_json}")

            # İlerleme durumu güncelle
            progress = int((index + 1) / total_rows * 100)
            self.progress_callback(progress)

        self.update_callback("Tüm kullanıcılar için işlemler tamamlandı.")
        self.progress_callback(100)  # İşlemler tamamlandığında %100 göster

class App:
    def __init__(self, root):
        self.root = root
        self.root.title('ArcGIS Kullanıcı İşlemleri')

        self.excel_path = None
        self.thread = None
        self.stop_event = threading.Event()

        self.create_widgets()

    def create_widgets(self):
        self.label = tk.Label(self.root, text="Excel dosyasını seçin:")
        self.label.pack(pady=10)

        # Butonları yerleştirmek için bir Frame oluştur
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=5)

        self.select_button = tk.Button(button_frame, text='Excel Dosyasını Seç', command=self.select_file)
        self.select_button.grid(row=0, column=0, padx=5)

        self.start_button = tk.Button(button_frame, text='Başlat', command=self.start_process, state=tk.DISABLED)
        self.start_button.grid(row=0, column=1, padx=5)

        self.stop_button = tk.Button(button_frame, text='Durdur', command=self.stop_process, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=2, padx=5)

        self.progress_bar = ttk.Progressbar(self.root, length=300, mode='determinate')
        self.progress_bar.pack(pady=10)

        self.log = tk.Text(self.root, height=10, width=50)
        self.log.pack(pady=10)

    def select_file(self):
        self.excel_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.excel_path:
            self.label.config(text=f"Seçilen dosya: {os.path.basename(self.excel_path)}")
            self.start_button.config(state=tk.NORMAL)

    def start_process(self):
        if self.excel_path:
            self.stop_event.clear()
            self.thread = WorkerThread(
                excel_path=self.excel_path,
                update_callback=self.update_log,
                progress_callback=self.update_progress,
                stop_event=self.stop_event
            )
            self.thread.start()
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.progress_bar['value'] = 0

    def stop_process(self):
        if self.thread:
            self.stop_event.set()
            self.thread.join()
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.update_log("İşlem durduruldu.")
            self.progress_bar['value'] = 0

    def update_log(self, message):
        self.log.insert(tk.END, message + "\n")
        self.log.yview(tk.END)

    def update_progress(self, value):
        self.progress_bar['value'] = value
        self.root.update_idletasks()

if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
