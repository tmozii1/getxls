from pywinauto import Application, Desktop

# 1) BuJa Chart 윈도우 찾기
main = Desktop(backend="uia").window(title_re=".*BuJa Chart.*")

print("=== BuJa Chart Window Found ===")
print("Title:", main.window_text())
print("Control Type:", main.element_info.control_type)

print("\n=== Listing child controls ===")
main.print_control_identifiers()
