import os
import sys
import time
import json
import ctypes
import pyautogui
import win32gui
import win32con
import requests
import tkinter as tk


SETTING_FILE = "setting.env"
WINDOW_TITLE_KEYWORD = "BuJa Chart"

def get_base_dir():
    if getattr(sys, 'frozen', False):   # EXE로 실행되는 경우
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

# ---------------------------------------------------------
# DPI 스케일 가져오기
# ---------------------------------------------------------
def get_dpi_scale():
    user32 = ctypes.windll.user32
    user32.SetProcessDPIAware()
    LOGPIXELSX = 88
    hdc = user32.GetDC(0)
    dpi_x = ctypes.windll.gdi32.GetDeviceCaps(hdc, LOGPIXELSX)
    return dpi_x / 96.0


# ---------------------------------------------------------
# BuJa Chart 윈도우 찾기
# ---------------------------------------------------------
def find_target_window():
    def enum_handler(hwnd, result):
        title = win32gui.GetWindowText(hwnd)
        if WINDOW_TITLE_KEYWORD.lower() in title.lower():
            result.append(hwnd)

    result = []
    win32gui.EnumWindows(enum_handler, result)

    if not result:
        raise Exception("❌ 'BuJa Chart'가 포함된 창을 찾을 수 없습니다.")

    return result[0]


def get_window_rect(hwnd):
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    return left, top, right - left, bottom - top


# ---------------------------------------------------------
# 설정파일 생성
# ---------------------------------------------------------
def create_default_setting():
    default = (
        "COLS=3\n"
        "ROWS=2\n"
        "OFFSETX=-60\n"
        "OFFSETY=150\n"
        "SERVER_URL=\n"
        "ITEMS=[\"Gold\", \"CrudeOil\", \"Euro\", \"Silver\", \"SNP\", \"Nasdaq\"]\n"
    )
    with open(SETTING_FILE, "w", encoding="utf-8") as f:
        f.write(default)


# ---------------------------------------------------------
# 설정 파일 로드
# ---------------------------------------------------------
def load_settings():
    settings = {}
    with open(SETTING_FILE, "r", encoding="utf-8") as f:
        for line in f:
            if "=" in line:
                key, val = line.strip().split("=", 1)
                settings[key.strip()] = val.strip()

    settings["COLS"] = int(settings["COLS"])
    settings["ROWS"] = int(settings["ROWS"])
    settings["OFFSETX"] = int(settings["OFFSETX"])
    settings["OFFSETY"] = int(settings["OFFSETY"])
    str = settings["SERVER_URL"]
    settings["SERVER_URL"].translate(str.maketrans('', '', '"\'')).strip()
    # JSON 형식 파싱
    settings["ITEMS"] = json.loads(settings["ITEMS"])

    return settings


# ---------------------------------------------------------
# 작은 GUI 오버레이
# ---------------------------------------------------------
class OverlayGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("자동저장")
        self.root.attributes("-topmost", True)
        self.root.overrideredirect(True)
        self.root.geometry(f"300x120+20+880")

        self.label = tk.Label(self.root, text="준비중...", justify="left", anchor="w", font=("맑은 고딕", 10))
        self.label.pack(pady=10, padx=10, anchor="w")

        self.btn = tk.Button(self.root, text="중지", width=10, command=self.stop)
        self.btn.pack(pady=5)

        self.is_stop = False

    def update_log(self, text):
        self.label.config(text=text)
        self.root.update()

    def stop(self):
        self.is_stop = True
        self.root.destroy()

    def finish(self):
        self.label.config(text="완료되었습니다.")
        self.btn.config(text="완료", command=self.root.destroy)
        self.root.update()


# ---------------------------------------------------------
# 마우스 제어
# ---------------------------------------------------------
def do_right_click(x, y):
    pyautogui.moveTo(x, y, duration=0.15)
    pyautogui.click(button="right")
    time.sleep(0.25)


def open_simulation():
    pyautogui.press("up")
    pyautogui.press("up")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.7)


def save_excel(name):
    for _ in range(6):
        pyautogui.press("tab")

    pyautogui.press("enter")
    time.sleep(0.6)

    pyautogui.typewrite(name + ".xls")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.6)

    pyautogui.press("esc")
    time.sleep(0.4)


# ---------------------------------------------------------
# 메인 루틴
# ---------------------------------------------------------
def main():
    # 설정 파일 없으면 생성 후 종료
    if not os.path.exists(SETTING_FILE):
        create_default_setting()
        tk.messagebox.showinfo("설정 생성", "설정파일이 생성되었습니다. (setting.env)")
        return
    
    gui = OverlayGUI()
    # 설정 로드
    cfg = load_settings()

    COLS = cfg["COLS"]
    ROWS = cfg["ROWS"]
    OFFSETX = cfg["OFFSETX"]
    OFFSETY = cfg["OFFSETY"]
    SERVER_URL = cfg["SERVER_URL"]
    ITEMS = cfg["ITEMS"]

    if len(ITEMS) != COLS * ROWS:
        tk.messagebox.showerror("오류", "items 개수와 COLS*ROWS가 일치하지 않습니다.")
        return
    
    clear_xls_files(DATA_DIR)
    gui.update_log(f"{DATA_DIR} 폴더 엑셀파일 삭제완료...")

    # BuJa Chart 창 찾기
    try:
        hwnd = find_target_window()
    except Exception as e:
        gui.update_log("BuJa Chart 창을 찾을 수 없습니다.")
        gui.btn.config(text="종료", command=gui.root.destroy)
        gui.root.mainloop()
        return

    # 창을 최상단으로
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)

    left, top, width, height = get_window_rect(hwnd)
    dpi_scale = get_dpi_scale()

    
    gui.update_log("작업 준비중...")

    time.sleep(2)

    idx = 0
    for y in range(ROWS):
        for x in range(COLS):

            if gui.is_stop:
                return

            item_name = ITEMS[idx]

            click_x = left + ((width / COLS) * (x + 1) + OFFSETX) / dpi_scale
            click_y = top + ((height / ROWS) * y + OFFSETY) / dpi_scale

            gui.update_log(f"차트 {idx+1} 저장 (x={int(click_x)}, y={int(click_y)})")

            do_right_click(click_x, click_y)
            open_simulation()
            save_excel(item_name)
            
            upload_to_server(item_name, SERVER_URL)

            idx += 1

    gui.finish()
    gui.root.mainloop()


def normalize_server_url(url: str):
    url = url.replace('"', '').replace('"', '')
    url = url.strip().rstrip("/")          # 마지막 / 제거
    if not url.endswith("/upload"):
        url = url + "/upload"              # upload 추가
    return url

def upload_to_server(item_name, server_url):
    
    # 서버가 비워져 있으면 전송 안함 
    if len(server_url) == 0:
        return

    local_file_path = os.path.join(DATA_DIR, item_name + ".xls")    
    try: 
        files = {"file": open(local_file_path, "rb")}
        url = normalize_server_url(server_url)
        r = requests.post(url, files=files, timeout=10)

        if r.status_code == 200:
            print(f"[UPLOAD] 성공 → {local_file_path}")
        else:
            print(f"[UPLOAD] 실패 → {r.text}")
    except Exception as e:
        print(f"[UPLOAD] 오류: {e}")


def clear_xls_files(folder):
    print(f"[삭제] 폴더 확인: {folder}")
    
    if not os.path.exists(folder):
        print("[삭제] 폴더 없음 → 자동 생성")
        os.makedirs(folder, exist_ok=True)
        return

    for file in os.listdir(folder):
        if file.lower().endswith(".xls"):
            full_path = os.path.join(folder, file)
            try:
                os.remove(full_path)
                print(f"[삭제 완료] {full_path}")
            except PermissionError:
                print(f"[삭제 실패: 권한 문제] {full_path}")
            except Exception as e:
                print(f"[삭제 오류] {full_path} → {e}")

def upload():    
    cfg = load_settings()
    items = cfg["ITEMS"]
    server_url = cfg["SERVER_URL"]
    
    if len(server_url) == 0:
        print(f"[UPLOAD] 실패: 서버URL이 없습니다.")
        return

    
    
    for item in items:
        upload_to_server(item, server_url)
        
                
def test():
    SERVER_URL = "http://localhost:8000"
    item_name = "Gold"
    local_path = os.path.join(DATA_DIR, item_name + ".xls")
    upload_to_server(local_path, SERVER_URL)

if __name__ == "__main__":
    #clear_xls_files("C:\부자차트")
    #clear_xls_files(DATA_DIR)
    main()
    #test()
