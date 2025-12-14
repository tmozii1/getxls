import os
import requests
import setting

def normalize_server_url(url: str):
    url = url.replace('"', '').replace('"', '')
    url = url.strip().rstrip("/")          # 마지막 / 제거
    if not url.endswith("/upload"):
        url = url + "/upload"              # upload 추가
    return url

def upload_to_server(item_name):
    cfg = setting.load_settings()
    server_url = cfg["SERVER_URL"]
    data_dir = cfg.get("PATH", "C:\\bujachart")
    upload_to_server2(item_name, server_url, data_dir)

def upload_to_server2(item_name, server_url, data_dir):
    
    # 서버가 비워져 있으면 전송 안함 
    if len(server_url) == 0:
        return

    local_file_path = os.path.join(data_dir, item_name + ".xls")    
    try: 
        files = {"file": open(local_file_path, "rb")}
        url = normalize_server_url(server_url)
        r = requests.post(url, files=files, timeout=10)

        if r.status_code == 200:
            print(f"[UPLOAD] 성공 → {local_file_path}")
        else:
            print(f"[UPLOAD] 실패 → {r.text}")
    except Exception as e:
        print(f"[UPLOAD] 오류: {e}")

def main():    
    cfg = setting.load_settings()
    items = cfg["ITEMS"]
    server_url = cfg["SERVER_URL"]
    
    # settings.py logic ensures PATH is clean
    data_dir = cfg.get("PATH", "C:\\bujachart")
    if not data_dir:
        data_dir = "C:\\bujachart"

    if len(server_url) == 0:
        print(f"[UPLOAD] 실패: 서버URL이 없습니다.")
        return

    for item in items:
        upload_to_server(item, server_url, data_dir)

if __name__ == "__main__":
    main()