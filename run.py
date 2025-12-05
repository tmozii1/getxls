import sys
import main      # 자동화 실행 파일
import setting     # 설정창 UI 파일

def print_help():
    filename = "run.exe"
    help_text = f"""
사용법: {filename} [옵션]

옵션 목록:
  --setting, -s, setting     설정창(UI) 실행
  --upload,  -u, upload      업로드 기능 실행
  --help,    -h, help        도움말 출력

예시:
  {filename} --setting
  {filename} -u
  {filename} -h
"""
    print(help_text)

def start():
    # 파라미터가 있을 경우
    if len(sys.argv) > 1:
        arg = sys.argv[1].lower()
        
        if arg in ("--help", "-h", "help"):
            print_help()
            return

        if arg in ("--setting", "-s", "setting"):
            # 설정 UI 실행
            setting.SettingsWindow(main.main)
            return
        
        if arg in ("--upload", "-u", "upload"):
            main.upload()
            return

    # 기본 실행 → 자동작업 실행
    main.main()


if __name__ == "__main__":
    start()
