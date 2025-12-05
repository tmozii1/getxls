import os
import json
import tkinter as tk
from tkinter import ttk, messagebox

import main

SETTING_FILE = "setting.env"


# ---------------------------------------------------------       
# setting.env 생성
# ---------------------------------------------------------
def create_default_setting():
    default = (
        "COLS=3\n"
        "ROWS=2\n"
        "OFFSETX=-60\n"
        "OFFSETY=150\n"
        'SERVER_URL="http://127.0.0.1:8000"\n'
        "ITEMS=[\"Gold\", \"CrudeOil\", \"Euro\", \"Silver\", \"SNP\", \"Nasdaq\"]\n"
    )
    with open(SETTING_FILE, "w", encoding="utf-8") as f:
        f.write(default)


# ---------------------------------------------------------
# setting.env 불러오기
# ---------------------------------------------------------
def load_settings():
    if not os.path.exists(SETTING_FILE):
        create_default_setting()

    settings = {}
    with open(SETTING_FILE, "r", encoding="utf-8") as f:
        for line in f:
            if "=" in line:
                key, val = line.strip().split("=", 1)
                settings[key] = val

    settings["COLS"] = int(settings["COLS"])
    settings["ROWS"] = int(settings["ROWS"])
    settings["OFFSETX"] = int(settings["OFFSETX"])
    settings["OFFSETY"] = int(settings["OFFSETY"])
    
    # SERVER_URL 큰따옴표 제거
    raw = settings["SERVER_URL"]
    if raw.startswith('"') and raw.endswith('"'):
        settings["SERVER_URL"] = raw[1:-1]   # 큰따옴표 제거
    else:
        settings["SERVER_URL"] = raw
    
    settings["ITEMS"] = json.loads(settings["ITEMS"])

    return settings


# ---------------------------------------------------------
# setting.env 저장하기
# ---------------------------------------------------------
def save_settings(COLS, ROWS, OFFSETX, OFFSETY, SERVER_URL, ITEMS):
    items_json = json.dumps(ITEMS)

    data = (
        f"COLS={COLS}\n"
        f"ROWS={ROWS}\n"
        f"OFFSETX={OFFSETX}\n"
        f"OFFSETY={OFFSETY}\n"
        f'SERVER_URL={SERVER_URL}\n'
        f"ITEMS={items_json}\n"
    )

    with open(SETTING_FILE, "w", encoding="utf-8") as f:
        f.write(data)


# ---------------------------------------------------------
# 설정 UI 클래스
# ---------------------------------------------------------
class SettingsWindow:
    def __init__(self, start_callback):
        self.start_callback = start_callback
        self.cfg = load_settings()

        # ---------------------------------------------------------
        # DPI 대응 (Windows 화면비율 자동감지)
        # ---------------------------------------------------------
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass

        # DPI 스케일 얻기
        def get_dpi_scale():
            user32 = ctypes.windll.user32
            dc = user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(dc, 88)  # LOGPIXELSX
            return dpi / 96.0  # 96dpi = 100%

        scale = get_dpi_scale()

        # Tkinter scaling 적용
        self.root = tk.Tk()
        self.root.tk.call('tk', 'scaling', scale)

        self.root.title("설정 (BuJa Chart Saver)")
        self.root.geometry(f"{int(420 * scale)}x{int(420 * scale)}")
        self.root.resizable(False, False)

        pad = int(15 * scale)
        entry_width = int(15 * scale)

        # COLS
        tk.Label(self.root, text="가로(COLS)").pack(anchor="w", padx=15, pady=(10, 0))
        self.cols_var = tk.StringVar(value=str(self.cfg["COLS"]))
        self.cols_box = ttk.Combobox(self.root, textvariable=self.cols_var, values=["1","2","3","4","5"], width=10)
        self.cols_box.pack(anchor="w", padx=15)

        # ROWS
        tk.Label(self.root, text="세로(ROWS)").pack(anchor="w", padx=15, pady=(10, 0))
        self.rows_var = tk.StringVar(value=str(self.cfg["ROWS"]))
        self.rows_box = ttk.Combobox(self.root, textvariable=self.rows_var, values=["1","2","3","4","5"], width=10)
        self.rows_box.pack(anchor="w", padx=15)

        # OFFSET X
        tk.Label(self.root, text="OFFSET_X").pack(anchor="w", padx=15, pady=(10, 0))
        self.offsetx_var = tk.StringVar(value=str(self.cfg["OFFSETX"]))
        tk.Entry(self.root, textvariable=self.offsetx_var, width=15).pack(anchor="w", padx=15)

        # OFFSET Y
        tk.Label(self.root, text="OFFSET_Y").pack(anchor="w", padx=15, pady=(10, 0))
        self.offsety_var = tk.StringVar(value=str(self.cfg["OFFSETY"]))
        tk.Entry(self.root, textvariable=self.offsety_var, width=15).pack(anchor="w", padx=15)

        # SERVER URL
        tk.Label(self.root, text="서버URL (공백 - 전송안함)").pack(anchor="w", padx=15, pady=(10, 0))
        self.server_var = tk.StringVar(value=self.cfg["SERVER_URL"])
        tk.Entry(self.root, textvariable=self.server_var, width=40).pack(anchor="w", padx=15)

        # ITEMS
        tk.Label(self.root, text="ITEMS (쉼표로 구분)").pack(anchor="w", padx=15, pady=(10, 0))
        self.items_var = tk.StringVar(value=",".join(self.cfg["ITEMS"]))
        tk.Entry(self.root, textvariable=self.items_var, width=40).pack(anchor="w", padx=15)

        # 안내 라벨
        tk.Label(self.root, text="파일명을 ,로 구분해서 가로X세로 만큼 입력해주세요.",
                 fg="red", font=("맑은 고딕", 9)).pack(anchor="w", padx=15, pady=5)

        # Start 버튼
        self.start_btn = tk.Button(self.root, text="저장", width=12, command=self.on_start)
        self.start_btn.pack(pady=20)

        self.root.mainloop()


    # ---------------------------------------------------------
    # Start 버튼 클릭
    # ---------------------------------------------------------
    def on_start(self):
        try:
            COLS = int(self.cols_var.get())
            ROWS = int(self.rows_var.get())
            OFFSETX = int(self.offsetx_var.get())
            OFFSETY = int(self.offsety_var.get())
        except:
            messagebox.showerror("입력 오류", "숫자 입력을 확인하세요.")
            return
        
        SERVER_URL = self.server_var.get().strip()
        ITEMS = [x.strip() for x in self.items_var.get().split(",") if x.strip()]

        if len(ITEMS) != COLS * ROWS:
            messagebox.showerror("ITEMS 오류", f"ITEMS 개수({len(ITEMS)})가 COLS*ROWS({COLS*ROWS})와 다릅니다.")
            return

        # 저장
        save_settings(COLS, ROWS, OFFSETX, OFFSETY, SERVER_URL, ITEMS)

        # 설정이 저장
        messagebox.showinfo("저장 완료", "설정이 저장되었습니다.")
        self.root.destroy()
        #self.start_callback()



# ---------------------------------------------------------
# 사용 예시
# ---------------------------------------------------------
if __name__ == "__main__":
    def fake_start():
        print("자동작업 시작!")

    SettingsWindow(main.main)
