import os
import sys
import time
import json
import ctypes
import pyautogui
import win32gui
import win32con
import upload
import tkinter as tk
import setting


WINDOW_TITLE_KEYWORD = "BuJa Chart"

DATA_DIR = "C:\\bujachart"

FOR_TEST = False

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

        self.root.after(1000, self.root.destroy)


# ---------------------------------------------------------
# 마우스 제어
# ---------------------------------------------------------
def do_right_click(x, y):
    print(f"do_right_click: x={x}, y={y}")
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

    if not ctypes.windll.shell32.IsUserAnAdmin():
        raise Exception("관리자 권한이 아닙니다.")
    
    gui = OverlayGUI()
    # 설정 로드
    cfg = setting.load_settings()

    COLS = cfg["COLS"]
    ROWS = cfg["ROWS"]
    OFFSETX = cfg["OFFSETX"]
    OFFSETY = cfg["OFFSETY"]
    X = cfg["X"]
    Y = cfg["Y"]
    WIDTH = cfg["WIDTH"]
    HEIGHT = cfg["HEIGHT"]
    PATH = cfg["PATH"]
    os.makedirs(PATH, exist_ok=True)
    DATA_DIR = PATH

    SERVER_URL = cfg["SERVER_URL"]
    ITEMS = cfg["ITEMS"]

    if len(ITEMS) != COLS * ROWS:
        raise Exception("items 개수와 COLS*ROWS가 일치하지 않습니다.")
    
    if not FOR_TEST:
        clear_xls_files(PATH)
        gui.update_log(f"{PATH} 폴더 엑셀파일 삭제완료...")

    # BuJa Chart 창 찾기
    try:
        hwnd = find_target_window()
    except Exception as e:
        raise Exception("buja chart창을 찾을 수 없습니다.")

    # 창을 최상단으로
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)

    left, top, width, height = get_window_rect(hwnd)

    dpi_scale = get_dpi_scale()

    
    gui.update_log("작업 준비중...")

    time.sleep(2)

    if not FOR_TEST:
        idx = 0
        for y in range(ROWS):
            for x in range(COLS):

                if gui.is_stop:
                    return

                item_name = ITEMS[idx]

                click_x = left + X + ((WIDTH / COLS) * (x + 1) + OFFSETX) / dpi_scale
                click_y = top + Y + ((HEIGHT / ROWS) * y + OFFSETY) / dpi_scale

                gui.update_log(f"차트 {idx+1} 저장 (x={int(click_x)}, y={int(click_y)})")

                do_right_click(click_x, click_y)
                open_simulation()
                save_excel(item_name)
                upload.upload_to_server(item_name)
                idx += 1

    gui.finish()
    gui.root.mainloop()




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



if __name__ == "__main__":
    #clear_xls_files("C:\부자차트")
    #clear_xls_files(DATA_DIR)
    main()
    #test()
