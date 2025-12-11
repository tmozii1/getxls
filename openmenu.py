import os
import time
import pyautogui
import setting
import win32gui
import win32con
import ctypes

# ---------------------------------------------------------
# 2) 해상도별 DPI 스케일 적용 (pyautogui 기준)
# ---------------------------------------------------------
try:
    dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100.0
except:
    dpi_scale = 1.0

# ---------------------------------------------------------
# 3) 첫번째 차트 영역 클릭 지점 계산
# ---------------------------------------------------------

def find_target_window():
    def enum_handler(hwnd, result):
        title = win32gui.GetWindowText(hwnd)
        if "BuJa Chart" in title:
            result.append(hwnd)
    result = []
    win32gui.EnumWindows(enum_handler, result)
    
    if not result:
        raise Exception("❌ 'BuJa Chart'가 포함된 창을 찾을 수 없습니다.")
    return result[0]

def get_first_chart_click_point(hwnd):

    cfg = setting.load_settings()
    COLS = int(cfg["COLS"])
    ROWS = int(cfg["ROWS"])
    OFFSETX = int(cfg["OFFSETX"])
    OFFSETY = int(cfg["OFFSETY"])
    X = int(cfg["X"])
    Y = int(cfg["Y"])
    WIDTH = int(cfg["WIDTH"])
    HEIGHT = int(cfg["HEIGHT"])

    rect = win32gui.GetWindowRect(hwnd)
    left, top = rect[0], rect[1]

    click_x = left + X + ((WIDTH / COLS) * (0 + 1) + OFFSETX) / dpi_scale
    click_y = top + Y + ((HEIGHT / ROWS) * 0 + OFFSETY) / dpi_scale

    return int(click_x), int(click_y)

# ---------------------------------------------------------
# 4) 메뉴 열기 동작
# ---------------------------------------------------------
def main():

    # BuJa Chart 창 위치 찾기
    try:
        hwnd = find_target_window()
    except Exception as e:
        print(f"[WARN] {e} -> Buja Chart창을 찾을 수 없습니다.")
        return

    # 창을 최상단으로
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)

    x, y = get_first_chart_click_point(hwnd)

    print(f"[INFO] Opening menu at ({x},{y}) ...")
    
    # 차트 첫 위치 클릭
    pyautogui.moveTo(x, y, duration=0.2)
    pyautogui.click(button="right")   # 오른쪽 클릭
    time.sleep(0.3)

    # 메뉴에서 'M' 키 입력
    pyautogui.press('m')
    print("[INFO] Right-click menu opened (pressed M).")
    time.sleep(2.0)

# 실제 실행
if __name__ == "__main__":
    main()
