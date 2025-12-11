import os
import json
import re
import pyautogui
import tkinter as tk
from tkinter import ttk, messagebox

import main

os.makedirs("dist", exist_ok=True)
SETTING_DIR = os.path.join(os.path.dirname(__file__), "dist")
SETTING_FILE = os.path.join(SETTING_DIR, "setting.env")

# ---------------------------------------------------------       
# setting.env 생성
# ---------------------------------------------------------
def create_default_setting():
    default = (
        "COLS=5\n"
        "ROWS=5\n"
        "OFFSETX=-60\n"
        "OFFSETY=150\n"
        "X=0\n"
        "Y=0\n"
        "WIDTH=1920\n"
        "HEIGHT=960\n"
        "PATH=\"C:\\\\bujachart\"\n"
        'SERVER_URL="https://buja.tim.pe.kr"\n'
        "ITEMS=[\"Gold\", \"CrudeOil\", \"EuroFX\", \"MiniSP500\", \"AustralianDollar\", \"Silver\", \"NaturalGas\", \"BritishPound\", \"MiniNASDAQ100\", \"NewZealandDollars\", \"Copper\", \"VIX\", \"Corn\", \"MiniDow\", \"CanadianDollar\", \"2YrUSTNote\", \"10YrUSTNote\", \"Soybeans\", \"HangSengIndex\", \"JapaneseYen\", \"5YrUSTNote\", \"30YrUSTBond\", \"BitCoin\", \"Ether\", \"Wheat\"]\n"
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
    settings["X"] = int(settings["X"])
    settings["Y"] = int(settings["Y"])
    settings["WIDTH"] = int(settings["WIDTH"])
    settings["HEIGHT"] = int(settings["HEIGHT"])

    # PATH 기본값 설정
    raw = settings.get("PATH", "").strip()
    if raw.startswith('"') and raw.endswith('"'):
        settings["PATH"] = raw[1:-1]   # 큰따옴표 제거
    else:
        settings["PATH"] = raw

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
def save_settings(COLS, ROWS, OFFSETX, OFFSETY, X, Y, WIDTH, HEIGHT, PATH, SERVER_URL, ITEMS):
    items_json = json.dumps(ITEMS)

    data = (
        f"COLS={COLS}\n"
        f"ROWS={ROWS}\n"
        f"OFFSETX={OFFSETX}\n"
        f"OFFSETY={OFFSETY}\n"
        f"X={X}\n"
        f"Y={Y}\n"
        f"WIDTH={WIDTH}\n"
        f"HEIGHT={HEIGHT}\n"
        f'PATH="{PATH}"\n'
        f'SERVER_URL="{SERVER_URL}"\n'
        f"ITEMS={items_json}\n"
    )

    with open(SETTING_FILE, "w", encoding="utf-8") as f:
        f.write(data)


# ---------------------------------------------------------
# 설정 UI 클래스
# ---------------------------------------------------------
class SettingsWindow:
    def __init__(self):
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
        self.root.geometry(f"{int(800 * scale)}x{int(600 * scale)}")
        self.root.resizable(False, False)

        # ===== 메인 프레임 (2열 레이아웃) =====
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # 왼쪽 열 (1열)
        left_frame = tk.Frame(main_frame)
        left_frame.pack(side="left", fill="both", padx=(0, 10))

        # COLS
        tk.Label(left_frame, text="가로(COLS)").pack(anchor="w", pady=(0, 2))
        self.cols_var = tk.StringVar(value=str(self.cfg["COLS"]))
        self.cols_box = ttk.Combobox(left_frame, textvariable=self.cols_var, values=["1","2","3","4","5"], width=15)
        self.cols_box.pack(anchor="w", pady=(0, 10))

        # ROWS
        tk.Label(left_frame, text="세로(ROWS)").pack(anchor="w", pady=(0, 2))
        self.rows_var = tk.StringVar(value=str(self.cfg["ROWS"]))
        self.rows_box = ttk.Combobox(left_frame, textvariable=self.rows_var, values=["1","2","3","4","5"], width=15)
        self.rows_box.pack(anchor="w", pady=(0, 10))

        # OFFSET X
        tk.Label(left_frame, text="OFFSET_X").pack(anchor="w", pady=(0, 2))
        self.offsetx_var = tk.StringVar(value=str(self.cfg["OFFSETX"]))
        tk.Entry(left_frame, textvariable=self.offsetx_var, width=15).pack(anchor="w", pady=(0, 10))

        # OFFSET Y
        tk.Label(left_frame, text="OFFSET_Y").pack(anchor="w", pady=(0, 2))
        self.offsety_var = tk.StringVar(value=str(self.cfg["OFFSETY"]))
        tk.Entry(left_frame, textvariable=self.offsety_var, width=15).pack(anchor="w", pady=(0, 10))

        # X
        tk.Label(left_frame, text="X (창 위치)").pack(anchor="w", pady=(0, 2))
        self.x_var = tk.StringVar(value=str(self.cfg["X"]))
        tk.Entry(left_frame, textvariable=self.x_var, width=15).pack(anchor="w", pady=(0, 10))

        # Y
        tk.Label(left_frame, text="Y (창 위치)").pack(anchor="w", pady=(0, 2))
        self.y_var = tk.StringVar(value=str(self.cfg["Y"]))
        tk.Entry(left_frame, textvariable=self.y_var, width=15).pack(anchor="w", pady=(0, 10))

        # WIDTH
        tk.Label(left_frame, text="WIDTH (창 크기)").pack(anchor="w", pady=(0, 2))
        self.width_var = tk.StringVar(value=str(self.cfg["WIDTH"]))
        tk.Entry(left_frame, textvariable=self.width_var, width=15).pack(anchor="w", pady=(0, 10))

        # HEIGHT
        tk.Label(left_frame, text="HEIGHT (창 크기)").pack(anchor="w", pady=(0, 2))
        self.height_var = tk.StringVar(value=str(self.cfg["HEIGHT"]))
        tk.Entry(left_frame, textvariable=self.height_var, width=15).pack(anchor="w", pady=(0, 10))

        # PATH
        tk.Label(left_frame, text="저장경로 (PATH)").pack(anchor="w", pady=(0, 2))
        self.path_var = tk.StringVar(value=self.cfg["PATH"])
        tk.Entry(left_frame, textvariable=self.path_var, width=50).pack(anchor="w", pady=(0, 10))

        # SERVER URL
        tk.Label(left_frame, text="서버URL").pack(anchor="w", pady=(0, 2))
        self.server_var = tk.StringVar(value=self.cfg["SERVER_URL"])
        tk.Entry(left_frame, textvariable=self.server_var, width=50).pack(anchor="w", pady=(0, 10))

        # 오른쪽 열 (2열) — ITEMS 텍스트박스
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side="right", fill="both", expand=True)

        tk.Label(right_frame, text="ITEMS (쉼표 또는 줄바꿈으로 구분)").pack(anchor="w", pady=(0, 5))
        
        items_count = len(self.cfg["ITEMS"])
        self.items_text = tk.Text(right_frame, width=40, height=17)
        self.items_text.pack(fill="both", expand=True)
        self.items_text.insert("1.0", "\n".join(self.cfg["ITEMS"]))

        # ===== 버튼 영역 =====
        button_frame = tk.Frame(self.root)
        button_frame.pack(fill="x", padx=15, pady=10)

        # 안내 라벨
        tk.Label(button_frame, text="파일명을 쉼표 또는 줄바꿈으로 구분해서 가로X세로 만큼 입력해주세요.",
                 fg="red", font=("맑은 고딕", 9)).pack(anchor="w", pady=(0, 10))

        # Start 버튼
        self.start_btn = tk.Button(button_frame, text="저장", width=12, command=self.on_start)
        self.start_btn.pack(anchor="center")

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
            X = int(self.x_var.get())
            Y = int(self.y_var.get())
            WIDTH = int(self.width_var.get())
            HEIGHT = int(self.height_var.get())
        except:
            messagebox.showerror("입력 오류", "숫자 입력을 확인하세요.")
            return
        
        PATH = self.path_var.get().strip()
        SERVER_URL = self.server_var.get().strip()

        items_raw = self.items_text.get("1.0", "end").strip()
        if items_raw:
            parts = re.split(r'[,\n]+', items_raw)
            ITEMS = [x.strip() for x in parts if x.strip()]
        else:
            ITEMS = []

        if len(ITEMS) != COLS * ROWS:
            messagebox.showerror("ITEMS 오류", f"ITEMS 개수({len(ITEMS)})가 COLS*ROWS({COLS*ROWS})와 다릅니다.")
            return

        # 저장
        save_settings(COLS, ROWS, OFFSETX, OFFSETY, X, Y, WIDTH, HEIGHT, PATH, SERVER_URL, ITEMS)

        # 설정이 저장
        messagebox.showinfo("저장 완료", "설정이 저장되었습니다.")
        self.root.destroy()
        #self.start_callback()


def main():
    SettingsWindow()

# ---------------------------------------------------------
# 사용 예시
# ---------------------------------------------------------
if __name__ == "__main__":
    def fake_start():
        print("자동작업 시작!")

    main()
