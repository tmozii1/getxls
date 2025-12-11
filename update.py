from pywinauto import Desktop
from pywinauto.keyboard import send_keys
import time
import requests
import pyperclip
import pyautogui
import tkinter as tk

# =========================================================
# 0) 사용자 입력
# =========================================================
keys = ["Gold", "CrudeOil", "EuroFX"]

URL_SMV = [
    f"https://buja.tim.pe.kr/daydb?key={keys[0]}",
    f"https://buja.tim.pe.kr/daydb?key={keys[1]}",
    f"https://buja.tim.pe.kr/daydb?key={keys[2]}",
]

URL_WEEK_DB  = f"https://buja.tim.pe.kr/weekdb?key0={keys[0]}&key1={keys[1]}&key2={keys[2]}"
URL_MONTH_DB = f"https://buja.tim.pe.kr/mondb?key0={keys[0]}&key1={keys[1]}&key2={keys[2]}"

func_map = {
    "S_MV": URL_SMV[0],
    "S_MV_2": URL_SMV[1],
    "S_MV_3": URL_SMV[2],
    "S_Week_DB": URL_WEEK_DB,
    "S_Mon_DB": URL_MONTH_DB,
}


# =========================================================
# 1) URL → text 불러오기
# =========================================================
def fetch_formula(url: str) -> str:
    try:
        r = requests.get(url, timeout=5)
        r.raise_for_status()
        return r.text
    except Exception as e:
        print("❌ URL 요청 실패:", url, e)
        return ""


# =========================================================
# 2) BuJa Chart 및 수식관리 TreeView 초기화(1회만 실행)
# =========================================================
def init_treeview(main):
    print("=== 초기 TreeView 준비 ===")

    # 1) 일반함수 클릭
    btn_general = main.child_window(title="일반함수", control_type="Button")
    btn_general.click_input()
    time.sleep(0.4)

    # 2) TreeView 클릭
    tree = main.child_window(auto_id="1235", control_type="Tree")
    tree.click_input()
    time.sleep(0.3)

    # 3) 기본함수 접기 → 사용자함수 이동 → 펼치기
    send_keys("{HOME}")    # 기본함수
    time.sleep(0.2)
    send_keys("{LEFT}")    # 접기
    time.sleep(0.2)
    send_keys("{DOWN}")    # 사용자함수
    time.sleep(0.2)
    send_keys("{RIGHT}")   # 펼치기
    time.sleep(0.3)

    print("✔ 사용자함수까지 확장 완료")


# =========================================================
# 3) 특정 함수 항목 찾기 + 클릭 + 수식 붙여넣기
# =========================================================
def update_formula(main, func_name: str, formula_text: str):
    print(f"\n===== {func_name} 변경 시작 =====")

    tree = main.child_window(auto_id="1235", control_type="Tree")

    # 전체 항목 검색
    items = tree.descendants(control_type="TreeItem")
    target = None
    for it in items:
        if it.window_text() == func_name:
            target = it
            break

    if target is None:
        print(f"❌ 함수명 {func_name} 를 찾을 수 없음")
        return

    # 스크롤 이동 + select + double click
    target.set_focus()
    time.sleep(0.2)
    target.double_click_input()
    time.sleep(0.5)

    # 수식 입력칸 (auto_id=16077)
    edit = main.child_window(auto_id="16077", control_type="Edit")
    edit.click_input()
    time.sleep(0.1)

    pyperclip.copy(formula_text)
    send_keys("^a")
    time.sleep(0.1)
    send_keys("^v")
    time.sleep(0.2)

    # 저장 버튼 클릭
    save_btn = main.child_window(title="작업저장", control_type="Button")
    save_btn.click_input()

    print(f"✔ {func_name} 저장 완료")


# =========================================================
# 4) 메인 실행
# =========================================================
def main():
    print("=== Connecting to BuJa Chart ===")
    main = Desktop(backend="uia").window(title_re=".*BuJa Chart.*")
    main.set_focus()

    # 초기 설정
    init_treeview(main)

    # 반복 업데이트
    for func_name, url in func_map.items():
        text = fetch_formula(url)
        if text.strip() == "":
            print(f"❌ URL 내용 없음: {func_name}")
            continue

        update_formula(main, func_name, text)

    # 닫기버튼
    # ToolBar 컨트롤 찾기
    toolbar = main.child_window(auto_id="59392", control_type="ToolBar")

    # 자식 버튼 나열
    buttons = toolbar.children()

    # 마지막 버튼이 닫기 버튼(Button9)
    close_btn = buttons[-1]   # 버튼 리스트의 마지막 항목
    close_btn.click_input()

    print("\n=== 전체 작업 완료 ===")
    tk.messagebox.showinfo("일일업데이트", "일일업데이트 완료")

if __name__ == "__main__":
    main()
