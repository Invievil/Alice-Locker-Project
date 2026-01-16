import customtkinter as ctk
import requests
import threading
import json
import os
import time
import pandas as pd
from datetime import datetime
from flask import Flask, request

# --- ФАЙЛЫ ДАННЫХ ---
DB_FILE = "locker_assignments.json"
CONFIG_FILE = "config.json"
EXCEL_REPORT = "security_report.xlsx"

# --- ЗАГРУЗКА НАСТРОЕК ---
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f: return json.load(f)
    return {
        "base_url": "http://192.168.100.127:9777/api/v1",
        "token": "YOUR_TOKEN_HERE",
        "zone_id": 1,
        "server_port": 5000
    }

config = load_config()

# --- ЛОГИРОВАНИЕ В EXCEL ---
def log_to_excel(event_type, message, sheet_name):
    new_row = {
        "Timestamp": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        "Type": event_type,
        "Details": message
    }
    df_new = pd.DataFrame([new_row])
    try:
        if os.path.exists(EXCEL_REPORT):
            with pd.ExcelWriter(EXCEL_REPORT, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                try:
                    df_existing = pd.read_excel(EXCEL_REPORT, sheet_name=sheet_name)
                    df_final = pd.concat([df_existing, df_new], ignore_index=True)
                    df_final.to_excel(writer, sheet_name=sheet_name, index=False)
                except: df_new.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            with pd.ExcelWriter(EXCEL_REPORT, engine='openpyxl') as writer:
                df_new.to_excel(writer, sheet_name=sheet_name, index=False)
    except: pass

# --- ФИЗИЧЕСКОЕ ОТКРЫТИЕ ---
opened_locks = set()
assignments = {}

def physical_open(num):
    url = f"{config['base_url']}/zones/open-lock"
    headers = {"Authorization": f"Bearer {config['token']}", "Content-Type": "application/json"}
    payload = {"zoneId": int(config['zone_id']), "lockNumber": int(num)}
    try:
        r = requests.post(url, json=payload, headers=headers, timeout=3)
        if r.status_code in [200, 201]:
            opened_locks.add(num)
            threading.Timer(5.0, lambda: opened_locks.discard(num)).start()
            return True
        log_to_excel("ERROR", f"Замок {num} код {r.status_code}", "System_Log")
    except Exception as e:
        log_to_excel("CRITICAL", f"Ошибка сети: {e}", "System_Log")
    return False

# --- СЕРВЕР ПРИЕМА СОБЫТИЙ ---
server = Flask(__name__)
waiting_card = None

@server.route('/sigur-event', methods=['POST', 'GET'])
def sigur_webhook():
    global waiting_card
    raw_data = request.data.decode('utf-8').strip()
    card_id = request.form.get('card_id') or raw_data
    if card_id:
        waiting_card = card_id.replace("card_id=", "").replace('"', '').strip()
        log_to_excel("CARD_READ", f"Карта {waiting_card} считана", "Usage_Log")
        return "OK", 200
    return "No Data", 200

# --- ОКНО НАСТРОЕК ---
class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Настройки")
        self.geometry("400x500")
        self.attributes("-topmost", True)
        self.u = self.add_f("API URL", config['base_url'])
        self.t = self.add_f("Token", config['token'])
        self.z = self.add_f("Zone ID", str(config['zone_id']))
        self.p = self.add_f("Server Port", str(config['server_port']))
        ctk.CTkButton(self, text="СОХРАНИТЬ", fg_color="#27AE60", command=self.save).pack(pady=20)

    def add_f(self, t, v):
        ctk.CTkLabel(self, text=t).pack(pady=(10,0))
        e = ctk.CTkEntry(self, width=320); e.insert(0, v); e.pack(); return e

    def save(self):
        config.update({"base_url": self.u.get(), "token": self.t.get(), "zone_id": int(self.z.get()), "server_port": int(self.p.get())})
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(config, f, indent=4)
        self.destroy()

# --- АДМИНИСТРАТОР (СВЕТЛАЯ ТЕМА) ---
class AdminWindow(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("ALICE ADMIN - CENTRAL CONTROL")
        self.geometry("620x900")
        self.configure(fg_color="#F0F0F0") 

        top_frame = ctk.CTkFrame(self, fg_color="#E0E0E0", corner_radius=0)
        top_frame.pack(fill="x", side="top")
        
        ctk.CTkButton(top_frame, text="⚙️", width=40, fg_color="#95a5a6", command=lambda: SettingsWindow(self)).pack(side="right", padx=10, pady=10)
        ctk.CTkButton(top_frame, text="📊 EXCEL ОТЧЕТ", fg_color="#27ae60", font=("Arial", 12, "bold"), 
                      command=lambda: os.startfile(EXCEL_REPORT) if os.path.exists(EXCEL_REPORT) else None).pack(side="right", padx=5)

        fix_frame = ctk.CTkFrame(self, fg_color="#FFFFFF", border_width=1, border_color="#BDC3C7")
        fix_frame.pack(fill="x", padx=20, pady=20)
        
        ctk.CTkLabel(fix_frame, text="ПРИВЯЗКА КАРТЫ К ЯЧЕЙКЕ", font=("Arial", 16, "bold"), text_color="#2C3E50").pack(pady=10)
        
        in_cont = ctk.CTkFrame(fix_frame, fg_color="transparent")
        in_cont.pack(fill="x", padx=20, pady=5)
        
        self.e_card = ctk.CTkEntry(in_cont, placeholder_text="ID КАРТЫ", height=50, 
                                   font=("Consolas", 20, "bold"), text_color="#2C3E50", fg_color="#F9F9F9")
        self.e_card.pack(side="left", expand=True, fill="x", padx=5)
        
        self.e_num = ctk.CTkEntry(in_cont, placeholder_text="№", width=80, height=50, 
                                  font=("Consolas", 20, "bold"), text_color="#2C3E50", fg_color="#F9F9F9")
        self.e_num.pack(side="left", padx=5)
        
        ctk.CTkButton(fix_frame, text="ЗАКРЕПИТЬ В БАЗЕ", height=45, font=("Arial", 14, "bold"), 
                      fg_color="#2980B9", command=self.manual_fix).pack(fill="x", padx=25, pady=15)

        ctk.CTkButton(self, text="🔓 ОТКРЫТЬ ВСЕ ЯЧЕЙКИ (АВАРИЙНО)", fg_color="#C0392B", height=40,
                      font=("Arial", 13, "bold"), command=self.open_all).pack(fill="x", padx=20, pady=5)

        self.scroll = ctk.CTkScrollableFrame(self, label_text="МОНИТОРИНГ", fg_color="#FFFFFF", label_text_color="#2C3E50")
        self.scroll.pack(expand=True, fill="both", padx=20, pady=10)
        self.update_list()

    def update_list(self):
        for w in self.scroll.winfo_children(): w.destroy()
        for i in range(1, 17):
            owner = assignments.get(i)
            row_bg = "#FDEDEC" if owner else "#FFFFFF"
            f = ctk.CTkFrame(self.scroll, fg_color=row_bg, border_width=1, border_color="#D5DBDB")
            f.pack(fill="x", pady=2, ipady=5)
            ctk.CTkLabel(f, text=f"Ячейка №{i:02d}", font=("Arial", 14, "bold"), text_color="#34495E", width=100).pack(side="left", padx=15)
            status_txt = f"ЗАНЯТА (ID: {owner})" if owner else "СВОБОДНО"
            ctk.CTkLabel(f, text=status_txt, font=("Consolas", 14), text_color="#E74C3C" if owner else "#27AE60", width=200, anchor="w").pack(side="left")
            
            b_cont = ctk.CTkFrame(f, fg_color="transparent")
            b_cont.pack(side="right", padx=10)
            ctk.CTkButton(b_cont, text="⚡", width=40, height=35, fg_color="#F1C40F", text_color="black", command=lambda n=i: physical_open(n)).pack(side="left", padx=2)
            if owner:
                ctk.CTkButton(b_cont, text="СБРОС", width=70, height=35, fg_color="#95A5A6", font=("Arial", 11, "bold"), command=lambda n=i: self.unfix(n)).pack(side="left", padx=2)

    def manual_fix(self):
        try:
            c, n = self.e_card.get().strip(), int(self.e_num.get().strip())
            if c and 1 <= n <= 16:
                assignments[n] = c
                with open(DB_FILE, 'w', encoding='utf-8') as f: json.dump({str(k):v for k,v in assignments.items()}, f)
                log_to_excel("ADMIN", f"Привязка {n} к {c}", "Usage_Log")
                self.update_list(); self.e_card.delete(0, 'end'); self.e_num.delete(0, 'end')
        except: pass

    def unfix(self, num):
        o = assignments.pop(num, None)
        with open(DB_FILE, 'w', encoding='utf-8') as f: json.dump({str(k):v for k,v in assignments.items()}, f)
        log_to_excel("ADMIN", f"Сброс {num} (был {o})", "Usage_Log")
        self.update_list()

    def open_all(self):
        def t(): 
            for i in range(1, 17): physical_open(i); time.sleep(0.3)
        threading.Thread(target=t, daemon=True).start()

# --- КЛИЕНТ (ТЕМНАЯ ТЕМА) ---
class ClientWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ALICE LOCKER SYSTEM")
        self.geometry("1100x850")
        
        # СТАТИКА: Начальный заголовок
        self.status = ctk.CTkLabel(self, text="ПРИЛОЖИТЕ КАРТУ", 
                                   font=("Arial", 30, "bold"), 
                                   text_color="#FFFFFF") # Можно поменять на черный #000000
        self.status.pack(pady=30)
        
        self.grid = ctk.CTkFrame(self)
        self.grid.pack(expand=True, fill="both", padx=20, pady=20)
        self.btns = {}
        for i in range(1, 17):
            btn = ctk.CTkButton(self.grid, text="", height=140, font=("Arial", 16, "bold"), command=lambda n=i: self.click(n))
            btn.grid(row=(i-1)//4, column=(i-1)%4, padx=5, pady=5, sticky="nsew")
            self.btns[i] = btn
            self.grid.grid_columnconfigure((0,1,2,3), weight=1)
        self.update_loop()

    def click(self, num):
        global waiting_card
        if not waiting_card: return
        owner = assignments.get(num)
        user_has = [k for k, v in assignments.items() if v == waiting_card]
        if owner is None and not user_has:
            if physical_open(num):
                assignments[num] = waiting_card
                with open(DB_FILE, 'w', encoding='utf-8') as f: json.dump({str(k):v for k,v in assignments.items()}, f)
                log_to_excel("USER", f"Взял шкаф {num}", "Usage_Log")
                waiting_card = None
                app.admin.update_list()
        elif owner == waiting_card:
            if physical_open(num):
                log_to_excel("USER", f"Открыл шкаф {num}", "Usage_Log")
                waiting_card = None
        else:
            self.status.configure(text="❌ Это ни ваш шкаф", text_color="red")

    def update_loop(self):
        # ДИНАМИКА: Тут меняем цвет "ПРИЛОЖИТЕ КАРТУ"
        active_color = "#2ECC71"  # Зеленый при карте
        waiting_color = "#000000" # Черный (твой выбор)
        
        current_color = active_color if waiting_card else waiting_color
        
        self.status.configure(
            text=f"КАРТА {waiting_card} АКТИВНА" if waiting_card else "ПРИЛОЖИТЕ КАРТУ", 
            text_color=current_color
        )
        
        for i, btn in self.btns.items():
            own = assignments.get(i); is_op = i in opened_locks
            btn.configure(fg_color="#27AE60" if (own and is_op) or not own else "#C0392B",
                          text=f"№{i}\n{'ОТКРЫТО' if is_op else 'ЗАНЯТО' if own else 'СВОБОДНО'}\n{f'ID: {own}' if own else ''}")
        self.after(300, self.update_loop)

# --- ЗАПУСК ---
if __name__ == "__main__":
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, 'r', encoding='utf-8') as f:
                d = json.load(f); assignments = {int(k):v for k,v in d.items()}
        except: pass
    log_to_excel("SYSTEM", "Старт", "System_Log")
    threading.Thread(target=lambda: server.run(host='0.0.0.0', port=config['server_port']), daemon=True).start()
    app = ClientWindow()
    app.admin = AdminWindow(app)
    app.mainloop()