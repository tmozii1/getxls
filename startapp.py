# startapp.py

import os
import sys
import subprocess
import pyautogui
import time
import win32gui
import win32con

TASK_NAME = "영웅문Global"


def activate_window(title_keyword="영웅문"):
    def enum_handler(hwnd, result):
        title = win32gui.GetWindowText(hwnd)
        if title_keyword in title:
            result.append(hwnd)

    windows = []
    win32gui.EnumWindows(enum_handler, windows)

    if not windows:
        print("❌ 영웅문 창을 찾지 못함")
        return False

    hwnd = windows[0]

    # 최소화되어 있으면 복원
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    # 전면 포커스
    win32gui.SetForegroundWindow(hwnd)
    return True

def run_app():
    print(f"[작업 실행] 작업 스케줄러 Task 실행: {TASK_NAME}")

    # schtasks 실행
    try:
        result = subprocess.run(
            ["schtasks", "/run", "/tn", TASK_NAME],
            capture_output=True,
            text=True,
            shell=True
        )

        # 오류 감지
        if result.returncode != 0:
            print("[오류] 작업 스케줄러 실행 실패")
            print("메시지:", result.stderr.strip())
            sys.exit(1)

        print("[완료] 영웅문Global 실행 요청됨.")
        
        time.sleep(5)
        # 비밀번호 자동 입력
        PASSWORD = "down7270"
        pyautogui.typewrite(PASSWORD)
        print("[완료] 비밀번호입력 완료")
        time.sleep(0.1)
        pyautogui.press("enter")
        print("[로그인 완료 요청됨]")                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
        time.sleep(1)

    except Exception as e:
        print("[예외 발생] 프로그램 실행 중 오류:", e)
        sys.exit(1)


if __name__ == "__main__":
    run_app()
