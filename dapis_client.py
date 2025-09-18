#!/usr/bin/env python3
"""
dapis_client.py
Python版クライアント
FastAPIサーバーに対象パスを送信し、ブラウザで検索画面を開く
"""

import sys
import os
import json
import uuid
import subprocess
import platform
import urllib.request

if len(sys.argv) < 2:
    print(f"Usage: {sys.argv[0]} path1 [path2 ...]")
    sys.exit(1)

# 絶対パスに変換
paths = [os.path.abspath(p) for p in sys.argv[1:]]

# UUID生成
sid = str(uuid.uuid4())

# JSON生成
data = {"session_id": sid, "paths": paths}
json_data = json.dumps(data).encode("utf-8")

# POSTリクエスト送信
req = urllib.request.Request(
    "http://127.0.0.1:8000/submit_targets",
    data=json_data,
    headers={"Content-Type": "application/json"},
    method="POST",
)

try:
    with urllib.request.urlopen(req) as resp:
        resp_json = json.load(resp)
        if not resp_json.get("ok"):
            print("Failed to submit targets:", resp_json)
            sys.exit(1)
except Exception as e:
    print("Failed to submit targets:", e)
    sys.exit(1)

# デフォルトブラウザで検索画面を開く
url = f"http://127.0.0.1:8000/?session_id={sid}"

if platform.system() == "Darwin":  # macOS
    subprocess.run(["open", url])
elif platform.system() == "Windows":
    subprocess.run(["start", url], shell=True)
else:  # Linuxなど
    subprocess.run(["xdg-open", url])

print(f"Targets submitted. Browser opened for session_id={sid}")

