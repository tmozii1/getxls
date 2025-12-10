import sys
import subprocess
import time
import os
from pywinauto.application import Application


# -------------------------------
# 공통 실행 함수
# -------------------------------
def run_python(script):
    """스크립트 실행 후 종료까지 대기"""
    print(f"[실행] {script}")
    proc = subprocess.Popen(["python", script])
    proc.wait()
    print(f"[완료] {script}\n")


# -------------------------------
# main 단계: XLS 생성
# -------------------------------
def run_getxls():
    run_python("main.py")


# -------------------------------
# 업로드 단계
# -------------------------------
def run_upload():
    run_python("upload.py")   # 팀장님 환경에 따라 파일명 조정 가능
    # 또는 main.py 내부 upload 함수만 따로 만든 경우 아래로 교체:
    # run_python("main.py --upload")


# -------------------------------
# openmenu 단계: 수식관리자 열기
# -------------------------------
def run_openmenu():
    run_python("openmenu.py")


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
    if not wait_formula_window():
        print("수식관리자 창이 없어 update.py 실행 불가.")
        return
    run_python("update.py")


# -------------------------------
# 설정파일 생성
# -------------------------------
def run_setting():
    print("[설정파일 생성]")
    dist = "dist"
    env_path = os.path.join(dist, "setting.env")

    if not os.path.exists(dist):
        os.makedirs(dist)
        print("dist 폴더 생성됨.")

    if not os.path.exists(env_path):
        with open(env_path, "w", encoding="utf-8") as f:
            f.write("# setting.env 자동 생성됨\n")
            f.write("COLS=5\nROWS=5\nOFFSETX=-60\nOFFSETY=150\nX=0\nY=0\n")
        print("setting.env 생성 완료.")
    else:
        print("이미 setting.env 존재함. 생성 스킵.")
    print()


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
        run_upload()
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
