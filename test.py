from pywinauto import Application, Desktop
import time

# 수식 문자열
NEW_FORMULA = """여기에 팀장님 수식 넣기"""

# 1) 수식관리자 창 연결
win = Desktop(backend="uia").window(title_re=".*BuJa Chart.*")
win.set_focus()
time.sleep(0.5)

# 2) 일반함수 클릭
btn_general = win.child_window(title="일반함수", control_type="Button")
btn_general.click_input()
time.sleep(0.5)

# 3) TreeView 찾기
tree = win.child_window(auto_id="1235", control_type="Tree")

# 4) TreeItem 목록 수집
items = tree.descendants(control_type="TreeItem")
print("TreeItems:", [i.window_text() for i in items])

user_item = None
for item in items:
    if "사용자함수" in item.window_text():
        user_item = item
        break

if not user_item:
    print("❌ 사용자함수 항목을 찾을 수 없음")
    exit()

# 5) 사용자함수 TreeItem 좌표 얻기
rect = user_item.rectangle()
center_x = rect.left + (rect.width() // 2)
center_y = rect.top + (rect.height() // 2)

# 6) 좌표 double click (가장 확실한 expand 방식)
print("✔ 사용자함수 더블클릭 (좌표 기반)")
user_item.click_input(button="left", coords=(center_x, center_y))
time.sleep(0.2)
user_item.click_input(button="left", coords=(center_x, center_y))
time.sleep(0.5)

# 7) 다시 Tree 목록 갱신 후 S_MV 찾기
items2 = tree.descendants(control_type="TreeItem")
s_mv = None
for it in items2:
    if it.window_text() == "S_MV":
        s_mv = it
        break

if not s_mv:
    print("❌ S_MV 안 나타남. 사용자함수 트리가 안 열렸다는 뜻.")
    exit()

# 8) S_MV 더블클릭
rect2 = s_mv.rectangle()
cx = rect2.left + rect2.width() // 2
cy = rect2.top + rect2.height() // 2

s_mv.click_input(coords=(cx, cy))
time.sleep(0.1)
s_mv.click_input(coords=(cx, cy))
time.sleep(0.5)

# 9) 수식 입력창에 수식 덮어쓰기
edit = win.child_window(auto_id="16077", control_type="Edit")
edit.click_input()
time.sleep(0.1)
edit.set_edit_text(NEW_FORMULA)

# 10) 저장
btn_save = win.child_window(title="작업저장", control_type="Button")
btn_save.click_input()

print("=== 완성: S_MV 업데이트 성공 ===")
