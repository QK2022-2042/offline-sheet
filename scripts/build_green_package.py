# -*- coding: utf-8 -*-
"""组装 offline-sheet 绿色便携包（Windows 64 位，目标机零安装）

前置条件（与本仓库同级的 build/ 目录）：
    build/python-embed.zip   python-3.11.9-embed-amd64.zip
                             https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip
    build/get-pip.py         https://bootstrap.pypa.io/pip/get-pip.py
    build/wheels/            openpyxl / python-docx / lxml 等离线轮子
                             pip download openpyxl python-docx -d build/wheels ^
                               --python-version 3.11 --platform win_amd64 --only-binary=:all:
产出：dist/offline-sheet/（内嵌便携 Python，拷到目标机双击 start.bat 即用）
"""
import os
import shutil
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                       # 仓库根目录
BUILD = os.path.join(ROOT, "build")
DIST = os.path.join(ROOT, "dist", "offline-sheet")

# 1. 清理并重建 dist
if os.path.exists(DIST):
    shutil.rmtree(DIST)
os.makedirs(DIST)

# 2. 解压 embeddable python
pydir = os.path.join(DIST, "python")
with zipfile.ZipFile(os.path.join(BUILD, "python-embed.zip")) as z:
    z.extractall(pydir)

# 3. 启用 site-packages（python311._pth）
pth = os.path.join(pydir, "python311._pth")
with open(pth, encoding="utf-8") as f:
    content = f.read()
if "#import site" in content:
    content = content.replace("#import site", "import site")
with open(pth, "w", encoding="utf-8") as f:
    f.write(content)

# 4. 拷贝代码
shutil.copy2(os.path.join(ROOT, "server.py"), os.path.join(DIST, "server.py"))
shutil.copytree(os.path.join(ROOT, "static"), os.path.join(DIST, "static"))

# 5. start.bat（UTF-8 无 BOM；chcp 65001 后 cmd 按 UTF-8 解析后续行，中文不乱码）
start_bat = """@echo off
chcp 65001 >nul
title 离线协同表格
cd /d %~dp0
set PYTHONIOENCODING=utf-8
set PORT=8081
if exist port.txt set /p PORT=<port.txt
echo.
echo   离线协同表格服务启动中（端口 %PORT%）...
echo   请勿关闭本窗口；编辑结束后直接关窗即可停止服务。
echo   如需换端口：在本目录新建 port.txt，写入端口号后重新启动。
echo.
"python\\python.exe" server.py --port %PORT%
pause
"""
with open(os.path.join(DIST, "start.bat"), "w", encoding="utf-8", newline="\r\n") as f:
    f.write(start_bat)

# 6. 防火墙放行.bat（需管理员）
fw_bat = """@echo off
netsh advfirewall firewall add rule name="OfflineSheet 8081" dir=in action=allow protocol=TCP localport=8081
echo Done. If access denied, run as Administrator.
pause
"""
with open(os.path.join(DIST, "防火墙放行.bat"), "w", encoding="gbk", newline="\r\n") as f:
    f.write(fw_bat)

print("assembled:", DIST)
print("下一步：")
print('  1) "%s" build/get-pip.py --no-input  （给便携 Python 装 pip）' % os.path.join(DIST, "python", "python.exe"))
print("  2) 便携 python -m pip install --no-index --find-links build/wheels openpyxl python-docx")
print("  3) 拷贝 docs/使用说明.txt 到包内（如需）")
