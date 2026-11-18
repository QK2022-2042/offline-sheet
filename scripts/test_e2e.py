# -*- coding: utf-8 -*-
"""端到端验证：xlsx/docx 导入 → 编辑 → 模板回写导出（需先启动服务，默认 8123 端口）"""
import io
import json
import os
import sys
import urllib.request

BASE = "http://127.0.0.1:8123"
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
os.makedirs(OUT, exist_ok=True)
FAILS = []


def check(name, cond, detail=""):
    tag = "PASS" if cond else "FAIL"
    print(f"[{tag}] {name}" + (f"  -- {detail}" if detail and not cond else ""))
    if not cond:
        FAILS.append(name)


def req(method, path, data=None, headers=None):
    r = urllib.request.Request(BASE + path, data=data, method=method,
                               headers=headers or {})
    with urllib.request.urlopen(r) as resp:
        return resp.read()


def make_test_xlsx():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "登记表"
    ws["A1"] = "设备登记表"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:D1")
    headers = ["名称", "数量", "单价", "金额"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=i, value=h)
        c.fill = PatternFill("solid", fgColor="FFD966")
        c.font = Font(bold=True)
    ws["A3"], ws["B3"], ws["C3"] = "示波器", 2, 8800
    ws["D3"] = "=B3*C3"                      # 公式
    ws["A4"], ws["B4"], ws["C4"] = "万用表", 5, 350
    ws["D4"] = "=B4*C4"
    ws.column_dimensions["A"].width = 24     # 列宽
    ws["A10"] = "合计"
    ws["B10"] = "=SUM(D3:D4)"
    wb.save(TXLSX := os.path.join(OUT, "template_test.xlsx"))
    return TXLSX


def make_test_docx():
    from docx import Document
    doc = Document()
    doc.add_paragraph("设备领用登记表（正文段落，导出时应原样保留）")
    tbl = doc.add_table(rows=3, cols=3)
    tbl.style = "Table Grid"
    # 表头横向合并 A1:C1
    a = tbl.cell(0, 0).text = "领用清单"
    tbl.cell(0, 0).merge(tbl.cell(0, 2))
    tbl.cell(1, 0).text = "设备名称"
    tbl.cell(1, 1).text = "领用人"
    tbl.cell(1, 2).text = "日期"
    tbl.cell(2, 0).text = "（待填写）"
    doc.save(TDOCX := os.path.join(OUT, "template_test.docx"))
    return TDOCX


TXLSX = make_test_xlsx()
TDOCX = make_test_docx()

# ---------- 1. 登录为上传人 ----------
owner = json.loads(req("POST", "/api/login", json.dumps({"name": "老板"}).encode()))
check("登录成为上传人", owner["role"] == "owner", str(owner))
T = owner["token"]
with open(os.path.join(OUT, "owner_token.txt"), "w") as f:
    f.write(T)

# ---------- 2. xlsx 链路 ----------
req("PUT", f"/api/upload?t={T}", open(TXLSX, "rb").read(),
    {"X-Filename": "template_test.xlsx"})
doc = json.loads(req("GET", f"/api/doc?t={T}"))
check("xlsx 导入解析", doc["title"] == "template_test" and doc["sheets"][0]["name"] == "登记表")
check("xlsx 合并单元格解析", [0, 0, 0, 3] in doc["sheets"][0]["merges"])
# 注：openpyxl 现造的文件没有公式缓存值，data_only 导入显示空串；
# 真实 Excel 另存的文件自带缓存，会显示计算结果。导出侧公式保留见下。
check("公式格导入不报错（无缓存显示空）", doc["sheets"][0]["rows"][2][3] == "",
      repr(doc["sheets"][0]["rows"][2][3]))

req("POST", f"/api/edit?t={T}", json.dumps(
    {"op": "set", "sheet": 0, "r": 4, "c": 0, "v": " 频谱仪"}).encode())
req("POST", f"/api/edit?t={T}", json.dumps(
    {"op": "set", "sheet": 0, "r": 2, "c": 1, "v": "3"}).encode())  # 改 B3 数量 2->3
req("POST", f"/api/edit?t={T}", json.dumps(
    {"op": "set", "sheet": 0, "r": 0, "c": 0, "v": " 设备登记表（已填写）"}).encode())

xlsx_out = os.path.join(OUT, "export_test.xlsx")
open(xlsx_out, "wb").write(req("GET", f"/api/export?fmt=xlsx&t={T}"))

from openpyxl import load_workbook
wb = load_workbook(xlsx_out)          # data_only=False 看公式
ws = wb["登记表"]
check("回写：改动值写入", ws["B3"].value == 3, repr(ws["B3"].value))
check("回写：新增行写入", ws["A5"].value == " 频谱仪", repr(ws["A5"].value))
check("回写：标题锚点格更新", str(ws["A1"].value).strip() == "设备登记表（已填写）",
      repr(ws["A1"].value))
check("保真：公式保留", ws["D3"].value == "=B3*C3", repr(ws["D3"].value))
check("保真：列宽保留", abs(ws.column_dimensions["A"].width - 24) < 0.1,
      str(ws.column_dimensions["A"].width))
check("保真：填充色保留", ws["B2"].fill.fgColor.rgb == "00FFD966",
      str(ws["B2"].fill.fgColor.rgb))
check("保真：合并单元格保留", "A1:D1" in [str(m) for m in ws.merged_cells.ranges])

# ---------- 3. docx 链路 ----------
req("PUT", f"/api/upload?t={T}", open(TDOCX, "rb").read(),
    {"X-Filename": "template_test.docx"})
doc = json.loads(req("GET", f"/api/doc?t={T}"))
sh = doc["sheets"][0]
check("docx 导入为表格 sheet", sh["name"] == "表格1")
check("docx 横向合并解析", [0, 0, 0, 2] in sh["merges"], str(sh["merges"]))
check("docx 表头文本", sh["rows"][0][0] == "领用清单" and sh["rows"][1][2] == "日期")

req("POST", f"/api/edit?t={T}", json.dumps(
    {"op": "set", "sheet": 0, "r": 2, "c": 0, "v": "网络分析仪"}).encode())
req("POST", f"/api/edit?t={T}", json.dumps(
    {"op": "set", "sheet": 0, "r": 2, "c": 1, "v": "王工"}).encode())

docx_out = os.path.join(OUT, "export_test.docx")
open(docx_out, "wb").write(req("GET", f"/api/export?fmt=docx&t={T}"))

from docx import Document
from docx.oxml.ns import qn
d2 = Document(docx_out)
check("docx 保真：正文段落保留",
      d2.paragraphs[0].text == "设备领用登记表（正文段落，导出时应原样保留）",
      d2.paragraphs[0].text)
t2 = d2.tables[0]
xml = t2._tbl.xml
check("docx 保真：gridSpan 保留", 'w:gridSpan' in xml)
check("docx 保真：表格样式保留", "Table Grid" not in xml and t2.style is not None)
check("docx 回写：锚点格", t2.cell(2, 0).text.strip() == "网络分析仪", t2.cell(2, 0).text)
check("docx 回写：同行第二列", t2.cell(2, 1).text.strip() == "王工", t2.cell(2, 1).text)
check("docx 回写：未动单元格不变", t2.cell(1, 2).text.strip() == "日期")

# ---------- 4. 权限 ----------
guest = json.loads(req("POST", "/api/login", json.dumps({"name": "小王"}).encode()))
ok_guest = guest["role"] == "guest"
try:
    req("GET", f"/api/export?fmt=xlsx&t={guest['token']}")
    ok_export = False
except Exception as e:
    ok_export = getattr(e, "code", None) == 403
check("协作者无上传人角色", ok_guest)
check("协作者导出被 403 拦截", ok_export)

print()
if FAILS:
    print("FAILED:", len(FAILS), FAILS)
    sys.exit(1)
print("ALL PASS")
