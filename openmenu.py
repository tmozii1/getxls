import os
import time
import pyautogui
from dotenv import dotenv_values

# ---------------------------------------------------------
# 1) settings.env 읽기
# ---------------------------------------------------------
cfg = dotenv_values(os.path.join("dist", "settings.env"))

COLS = int(cfg["COLS"])
ROWS = int(cfg["ROWS"])
OFFSETX = int(cfg["OFFSETX"])
OFFSETY = int(cfg["OFFSETY"])
X = int(cfg["X"])
Y = int(cfg["Y"])
WIDTH = int(cfg["WIDTH"])
HEIGHT = int(cfg["HEIGHT"])

# ---------------------------------------------------------
# 2) 해상도별 DPI 스케일 적용 (pyautogui 기준)
# ---------------------------------------------------------
try:
    import ctypes
    dpi_scale = ctypes.windll.shcore.GetScaleFactorForDevice(0) / 100.0
except:
    dpi_scale = 1.0

# ---------------------------------------------------------
# 3) 첫번째 차트 영역 클릭 지점 계산
# ---------------------------------------------------------
# main.py 기준 공식 동일
# x=0, y=0 기준 첫 차트
def get_first_chart_click_point():
    # 현재 모니터(전체 화면) 기준 좌표
    screen_width, screen_height = pyautogui.size()
    
    # 차트 전체 화면이 항상 왼쪽 상단이라는 가정
    left, top = 0, 0

    click_x = left + X + ((WIDTH / COLS) * (0 + 1) + OFFSETX) / dpi_scale
    click_y = top + Y + ((HEIGHT / ROWS) * 0 + OFFSETY) / dpi_scale

    return int(click_x), int(click_y)

# ---------------------------------------------------------
# 4) 메뉴 열기 동작
# ---------------------------------------------------------
def open_formula_manager_menu():
    x, y = get_first_chart_click_point()

    print(f"[INFO] Opening menu at ({x},{y}) ...")

    # 차트 첫 위치 클릭
    pyautogui.moveTo(x, y, duration=0.2)
    pyautogui.click(button="right")   # 오른쪽 클릭
    time.sleep(0.3)

    # 메뉴에서 'M' 키 입력
    pyautogui.press('m')
    print("[INFO] Right-click menu opened (pressed M).")
    time.sleep(0.5)

# 실제 실행
if __name__ == "__main__":
    open_formula_manager_menu()
    # 이후 S_MV 업데이트 로직 실행
