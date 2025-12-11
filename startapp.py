# startapp.py

import os
import sys
import subprocess
import pyautogui
import time

TASK_NAME = "영웅문Global"


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
        
        time.sleep(3)
        # 비밀번호 자동 입력
        PASSWORD = "down7270"
        pyautogui.typewrite(PASSWORD)
        time.sleep(0.1)
        pyautogui.press("enter")

        print("[로그인 완료 요청됨]")

    except Exception as e:
        print("[예외 발생] 프로그램 실행 중 오류:", e)
        sys.exit(1)


if __name__ == "__main__":
    run_app()
