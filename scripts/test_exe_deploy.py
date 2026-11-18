# -*- coding: utf-8 -*-
"""单文件 EXE 真实部署测试：
1) 拷贝 exe 到独立部署目录（模拟真实使用）
2) 启动 → 全量端到端测试
3) 停服重启 → 验证登录态/内容不丢、模板回写仍工作、data 落在 exe 旁

前置：先用 PyInstaller 打出 dist_exe/OfflineSheet.exe
    pyinstaller --onefile --add-data "static;static" --name OfflineSheet server.py
"""
import json
import os
import shutil
import socket
import subprocess
import sys
import time
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
EXE_SRC = os.path.join(ROOT, "dist_exe", "OfflineSheet.exe")
DEPLOY = os.path.join(HERE, "deploy_test")
EXE = os.path.join(DEPLOY, "OfflineSheet.exe")
E2E = os.path.join(HERE, "test_e2e.py")
OUT = os.path.join(HERE, "out")
PY = sys.executable

os.makedirs(DEPLOY, exist_ok=True)
shutil.copy2(EXE_SRC, EXE)
shutil.rmtree(os.path.join(DEPLOY, "data"), ignore_errors=True)


def start():
    return subprocess.Popen([EXE, "--port", "8123"], cwd=DEPLOY,
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT)


def wait_up(proc, timeout=120):
    t0 = time.time()
    while time.time() - t0 < timeout:
        if proc.poll() is not None:
            raise RuntimeError("exe 提前退出: " + proc.stdout.read().decode(errors="ignore"))
        try:
            socket.create_connection(("127.0.0.1", 8123), timeout=1).close()
            return
        except OSError:
            time.sleep(0.5)
    raise RuntimeError("等待服务启动超时")


def stop(proc):
    proc.terminate()
    try:
        proc.wait(10)
    except Exception:
        proc.kill()


fails = []

p = start()
wait_up(p)
print("== ROUND 1: 全量端到端（对 EXE）==")
r = subprocess.run([PY, E2E])
if r.returncode != 0:
    fails.append("round1-e2e")

print("== ROUND 2: 停服重启持久化 ==")
stop(p)
time.sleep(2)
p = start()
wait_up(p)

with open(os.path.join(OUT, "owner_token.txt")) as f:
    t = f.read().strip()
doc = json.loads(urllib.request.urlopen(f"http://127.0.0.1:8123/api/doc?t={t}").read())
ok_login = doc["me"]["role"] == "owner"
print("PASS" if ok_login else "FAIL", "重启后 token 有效、角色保持 owner")
if not ok_login:
    fails.append("restart-login")

ok_data = (doc["title"] == "template_test"
           and doc["sheets"][0]["rows"][2][0].strip() == "网络分析仪")
print("PASS" if ok_data else "FAIL", "重启后填写内容不丢")
if not ok_data:
    fails.append("restart-data")

edit = urllib.request.Request(
    f"http://127.0.0.1:8123/api/edit?t={t}",
    data=json.dumps({"op": "set", "sheet": 0, "r": 2, "c": 2, "v": "李工"}).encode(),
    method="POST")
urllib.request.urlopen(edit)
out_docx = os.path.join(OUT, "export_after_restart.docx")
with open(out_docx, "wb") as f:
    f.write(urllib.request.urlopen(f"http://127.0.0.1:8123/api/export?fmt=docx&t={t}").read())
from docx import Document
v = Document(out_docx).tables[0].cell(2, 2).text.strip()
ok_rw = v == "李工"
print("PASS" if ok_rw else "FAIL", f"重启后导出仍回写模板（cell={v!r}）")
if not ok_rw:
    fails.append("restart-writeback")

ok_dir = os.path.isdir(os.path.join(DEPLOY, "data"))
print("PASS" if ok_dir else "FAIL", "data 目录生成在 exe 旁")
if not ok_dir:
    fails.append("data-beside-exe")

stop(p)
print()
if fails:
    print("FAILED:", fails)
    sys.exit(1)
print("EXE DEPLOY TEST: ALL PASS")
