import sys

import time
import os
from pywinauto.application import Application

import main
import upload
import update
import setting
import openmenu


# -------------------------------
# main 단계: XLS 생성
# -------------------------------
def run_getxls():
    try:
        main.main()
    except Exception as e:
        print(f"[getxls] 오류: {e}")
        sys.exit(1)


# -------------------------------
# 업로드 단계
# -------------------------------
def run_upload():
    try:
        upload.main()
    except Exception as e:
        print(f"[upload] 오류: {e}")
        sys.exit(1)


# -------------------------------
# openmenu 단계: 수식관리자 열기
# -------------------------------
def run_openmenu():
    try:
        openmenu.main()
    except Exception as e:
        print(f"[openmenu] 오류: {e}")
        sys.exit(1)


# -------------------------------
# 수식관리자 창 대기
# -------------------------------
def wait_formula_window(timeout=15):
    app = Application(backend="uia")
    print("수식관리자 창을 기다리는 중...")

    for _ in range(timeout * 10):
        try:
            for w in app.windows():
                name = w.window_text()
                if "수식관리" in name or "사용자지표" in name:
                    print("수식관리자 창 확인됨!")
                    return True
        except:
            pass
        time.sleep(0.1)

    print("수식관리자 창을 찾지 못함.")
    return False


# -------------------------------
# update 단계
# -------------------------------
def run_update():
    try:
        update.main()
    except Exception as e:
        print(f"[update] 오류: {e}")
        sys.exit(1)


# -------------------------------
# 설정파일 생성
# -------------------------------
def run_setting():
    try:
        setting.main()
    except Exception as e:
        print(f"[setting] 오류: {e}")
        sys.exit(1)



# -------------------------------
# 도움말
# -------------------------------
HELP_TEXT = """
사용법: run.py [옵션]

옵션 목록:
  --help        : 도움말 출력
  --setting     : dist/setting.env 생성
  --getxls      : XLS 생성(main.py 실행)
  --upload      : XLS 서버 업로드(upload.py 실행)
  --auto        : 수식관리자 열기 → 업데이트(Update.py)

파라미터 없을 때:
  getxls → upload → auto 전체 자동 실행
"""


# -------------------------------
# 메인 실행
# -------------------------------
if __name__ == "__main__":
    args = sys.argv[1:]

    if not args:
        # 전체 자동 실행
        print("[전체 자동 실행: getxls → upload → auto]")
        run_getxls()
        #run_upload()
        run_openmenu()
        run_update()
        sys.exit(0)

    # 단일 옵션 실행
    cmd = args[0].lower()

    if cmd == "--help":
        print(HELP_TEXT)
    elif cmd == "--setting":
        run_setting()
    elif cmd == "--getxls":
        run_getxls()
    elif cmd == "--upload":
        run_upload()
    elif cmd == "--auto":
        run_openmenu()
        run_update()
    else:
        print("알 수 없는 옵션입니다. --help 참고.")
