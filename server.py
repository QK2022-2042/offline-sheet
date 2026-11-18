#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
离线局域网 Excel 协同编辑服务（零第三方依赖，除 openpyxl/python-docx 用于导入导出）

    python3 server.py --port 8081
    浏览器访问 http://<本机IP>:8081

角色：
    owner  上传人 —— 可导入、可导出下载、可移除他人
    guest  其他同事 —— 只能在线编辑，导出接口直接 403

数据：data/doc.json（每次修改原子落盘），重启不丢。
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import os
import queue
import re
import secrets
import sys
import threading
import time
import traceback
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

# PyInstaller 打包后：资源（static）在临时解压目录 _MEIPASS，
# 数据（data/）必须落在 exe 旁边，否则重启丢数据。
if getattr(sys, "frozen", False):
    RESOURCE = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    BASE = os.path.dirname(os.path.abspath(sys.executable))
else:
    RESOURCE = BASE = os.path.dirname(os.path.abspath(__file__))
STATIC = os.path.join(RESOURCE, "static")
DATA = os.path.join(BASE, "data")
os.makedirs(DATA, exist_ok=True)
DOC_PATH = os.path.join(DATA, "doc.json")
TEMPLATE_XLSX = os.path.join(DATA, "template.xlsx")
TEMPLATE_DOCX = os.path.join(DATA, "template.docx")

MAX_ROWS, MAX_COLS = 5000, 200
LOCK = threading.RLock()
CLIENTS: dict[str, "Client"] = {}      # token -> Client
COLORS = ["#2563eb", "#db2777", "#059669", "#d97706", "#7c3aed",
          "#0891b2", "#dc2626", "#65a30d", "#c026d3", "#475569"]


def now() -> float:
    return time.time()


def blank_doc() -> dict:
    return {
        "title": "未命名表格",
        "owner": None,
        "version": 0,
        "seq": 0,
        "template": None,
        "baseline": [],
        "sheets": [{"name": "Sheet1", "rows": [["" for _ in range(12)] for _ in range(40)],
                    "cols": 12, "merges": []}],
        "updated_at": now(),
    }


class Client:
    __slots__ = ("token", "name", "role", "color", "q", "last")

    def __init__(self, token: str, name: str, role: str, color: str):
        self.token = token
        self.name = name
        self.role = role
        self.color = color
        self.q: queue.Queue = queue.Queue(maxsize=200)
        self.last = now()

    def roster(self) -> dict:
        return {"token": self.token, "name": self.name, "role": self.role,
                "color": self.color, "idle": round(now() - self.last, 1)}

    def send(self, payload: dict) -> None:
        try:
            self.q.put_nowait(payload)
        except queue.Full:
            pass


# ---------------------------------------------------------------- 文档读写
def load_doc() -> dict:
    if os.path.exists(DOC_PATH):
        try:
            with open(DOC_PATH, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            traceback.print_exc()
    return blank_doc()


DOC = load_doc()
DOC.setdefault("members", {})
DOC.setdefault("template", None)
DOC.setdefault("baseline", [])


def save_doc() -> None:
    tmp = DOC_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(DOC, f, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, DOC_PATH)


def bump() -> int:
    DOC["version"] += 1
    DOC["seq"] += 1
    DOC["updated_at"] = now()
    return DOC["version"]


def broadcast(payload: dict, exclude: str | None = None) -> None:
    for token, c in list(CLIENTS.items()):
        if token != exclude:
            c.send(payload)


def roster() -> list[dict]:
    return [c.roster() for c in sorted(CLIENTS.values(), key=lambda x: x.last)]


def push_roster() -> None:
    broadcast({"type": "roster", "users": roster()})


def resolve(token: str) -> Client | None:
    """token -> Client；服务重启后凭 members 记录自动重建，不必重新登录。"""
    c = CLIENTS.get(token)
    if c:
        c.last = now()
        return c
    m = DOC.get("members", {}).get(token)
    if not m:
        return None
    c = Client(token, m["name"], m["role"], m.get("color") or color_for(m["name"]))
    CLIENTS[token] = c
    return c


def color_for(name: str) -> str:
    return COLORS[sum(name.encode()) % len(COLORS)]


# 兼容：老数据文件只记了 owner token，没有 members 记录
if DOC.get("owner") and DOC["owner"] not in DOC["members"]:
    DOC["members"][DOC["owner"]] = {"name": "上传人", "role": "owner", "color": COLORS[0]}
    save_doc()


# ---------------------------------------------------------------- 导入 / 导出
NUMERIC = re.compile(r"^-?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$")


def xlsx_to_state(blob: bytes) -> list[dict]:
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(blob), data_only=True)
    sheets = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        nrows = min(ws.max_row or 0, MAX_ROWS)
        ncols = min(ws.max_column or 0, MAX_COLS)
        rows = []
        for r in range(1, max(nrows, 40) + 1):
            row = []
            for c in range(1, max(ncols, 12) + 1):
                v = ws.cell(row=r, column=c).value if (r <= nrows and c <= ncols) else None
                if v is None:
                    row.append("")
                elif isinstance(v, float) and v.is_integer():
                    row.append(str(int(v)))
                else:
                    row.append(str(v))
            rows.append(row)
        merges = []
        for rng in ws.merged_cells.ranges:
            merges.append([rng.min_row - 1, rng.min_col - 1, rng.max_row - 1, rng.max_col - 1])
        sheets.append({"name": ws.title, "rows": rows,
                       "cols": max(ncols, 12), "merges": merges})
    return sheets or blank_doc()["sheets"]


def state_to_xlsx() -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    wb = Workbook()
    wb.remove(wb.active)
    for sh in DOC["sheets"]:
        ws = wb.create_sheet(title=(sh["name"] or "Sheet")[:31])
        for ri, row in enumerate(sh["rows"]):
            if ri >= MAX_ROWS:
                break
            for ci, val in enumerate(row[:MAX_COLS]):
                if val == "":
                    continue
                cell = ws.cell(row=ri + 1, column=ci + 1)
                if NUMERIC.match(str(val).strip()):
                    try:
                        cell.value = int(val) if re.fullmatch(r"-?\d+", str(val).strip()) else float(str(val).replace(",", ""))
                        continue
                    except Exception:
                        pass
                cell.value = str(val)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        for m in sh.get("merges", []):
            try:
                ws.merge_cells(start_row=m[0] + 1, start_column=m[1] + 1,
                               end_row=m[2] + 1, end_column=m[3] + 1)
            except Exception:
                pass
        ws.freeze_panes = "A1"
        for ci in range(1, min(len(sh["rows"][0]) if sh["rows"] else 1, MAX_COLS) + 1):
            ws.column_dimensions[ws.cell(row=1, column=ci).column_letter].width = 16
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def state_to_docx() -> bytes:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.section import WD_ORIENT
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    doc.add_heading(DOC["title"], level=1)
    for sh in DOC["sheets"]:
        rows, width = [], 0
        for r in sh["rows"]:
            vals = list(r)
            while vals and not str(vals[-1]).strip():
                vals.pop()
            if vals:
                rows.append(vals)
                width = max(width, len(vals))
        if not rows:
            continue
        doc.add_heading(sh["name"], level=2)
        table = doc.add_table(rows=len(rows), cols=width)
        table.style = "Table Grid"
        for ri, row in enumerate(rows):
            for ci in range(width):
                cell = table.cell(ri, ci)
                cell.text = str(row[ci]) if ci < len(row) else ""
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(8)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ---------------------------------------------------------------- Word 表格型表单
def _tc_text(tc) -> str:
    """读取 w:tc 内所有段落文本，段落间以换行拼接。"""
    from docx.oxml.ns import qn
    out = []
    for p in tc.findall(qn("w:p")):
        line = "".join((t.text or "") for t in p.findall(".//" + qn("w:t")))
        out.append(line)
    return "\n".join(out).strip()


def _tc_span(tc) -> tuple[int, str | None]:
    """返回 (横向跨度 gridSpan, 纵向合并标记 vMerge)。vMerge: None/restart/continue"""
    from docx.oxml.ns import qn
    tcPr = tc.find(qn("w:tcPr"))
    h, vm = 1, None
    if tcPr is not None:
        gs = tcPr.find(qn("w:gridSpan"))
        if gs is not None:
            h = max(1, int(gs.get(qn("w:val")) or 1))
        v = tcPr.find(qn("w:vMerge"))
        if v is not None:
            vm = v.get(qn("w:val")) or "restart"
    return h, vm


def _table_grid(tbl):
    """遍历 w:tbl，展开 gridSpan/vMerge 到规则网格。
    返回 (anchors, merges, merge_map, nrows, ncols)：
      anchors[(r,c)] = 锚点 w:tc 元素（可写入的单元格）
      merges        = [[r1,c1,r2,c2], ...] 零基合并区域
      merge_map     = {"r:c": [ar,ac]} 被覆盖格 -> 锚点坐标
    """
    from docx.oxml.ns import qn
    trs = tbl.findall(qn("w:tr"))
    nrows = len(trs)
    ncols = 0
    for tr in trs:
        w = sum(_tc_span(tc)[0] for tc in tr.findall(qn("w:tc")))
        ncols = max(ncols, w)
    anchors: dict = {}
    merges: list = []
    merge_map: dict = {}
    covered = [[False] * ncols for _ in range(nrows)]
    vopen: dict = {}   # 起始列 -> {"r1","c1","h"}
    for ri, tr in enumerate(trs):
        ci = 0
        for tc in tr.findall(qn("w:tc")):
            h, vm = _tc_span(tc)
            while ci < ncols and covered[ri][ci]:
                ci += 1
            if ci >= ncols:
                break
            c2 = min(ci + h - 1, ncols - 1)
            if vm == "continue":
                anchor = None
                for sc, o in vopen.items():
                    if o["c1"] <= ci <= o["c1"] + o["h"] - 1:
                        anchor = o
                        break
                if anchor:
                    for cc in range(ci, c2 + 1):
                        covered[ri][cc] = True
                        merge_map[f"{ri}:{cc}"] = [anchor["r1"], anchor["c1"]]
                else:  # 容错：没有找到开启的纵向合并，按普通格处理
                    anchors[(ri, ci)] = tc
                    for cc in range(ci + 1, c2 + 1):
                        covered[ri][cc] = True
                        merge_map[f"{ri}:{cc}"] = [ri, ci]
                ci = c2 + 1
                continue
            # 新锚点：同列上有未闭合的纵向合并则在此行之前结束
            if ci in vopen:
                o = vopen.pop(ci)
                r2 = max(o["r1"], ri - 1)
                merges.append([o["r1"], o["c1"], r2, o["c1"] + o["h"] - 1])
            anchors[(ri, ci)] = tc
            for cc in range(ci + 1, c2 + 1):
                covered[ri][cc] = True
                merge_map[f"{ri}:{cc}"] = [ri, ci]
            if vm == "restart" or c2 > ci:
                vopen[ci] = {"r1": ri, "c1": ci, "h": c2 - ci + 1}
            ci = c2 + 1
    for o in vopen.values():  # 表格结束时仍未闭合的纵向合并
        merges.append([o["r1"], o["c1"], nrows - 1, o["c1"] + o["h"] - 1])
    return anchors, merges, merge_map, nrows, ncols


def docx_to_state(blob: bytes) -> list[dict]:
    """把 docx 里的每个表格转成一个 sheet；正文段落不支持（已确认边界）。"""
    from docx import Document
    doc = Document(io.BytesIO(blob))
    sheets = []
    for ti, tbl in enumerate(doc.tables, 1):
        anchors, merges, merge_map, nrows, ncols = _table_grid(tbl._tbl)
        grid = [["" for _ in range(ncols)] for _ in range(nrows)]
        for (r, c), tc in anchors.items():
            grid[r][c] = _tc_text(tc)
        while len(grid) < 40:
            grid.append(["" for _ in range(ncols)])
        if ncols < 12:
            for row in grid:
                row.extend([""] * (12 - ncols))
        sheets.append({"name": f"表格{ti}", "rows": grid,
                       "cols": max(ncols, 12), "merges": merges,
                       "merge_map": merge_map})
    return sheets or blank_doc()["sheets"]


def _set_tc_text(tc, text: str) -> None:
    """把文本写入 w:tc：保留第一个段落的段落属性，清掉多余段落与旧 run。"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    ps = tc.findall(qn("w:p"))
    if not ps:
        p = OxmlElement("w:p")
        tc.append(p)
        ps = [p]
    p = ps[0]
    for extra in ps[1:]:
        tc.remove(extra)
    for r in p.findall(qn("w:r")):
        p.remove(r)
    run = OxmlElement("w:r")
    for i, ln in enumerate(str(text).split("\n")):
        if i:
            run.append(OxmlElement("w:br"))
        t = OxmlElement("w:t")
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = ln
        run.append(t)
    p.append(run)


def writeback_xlsx() -> bytes:
    """加载原始 xlsx 模板，仅回写与基线不同的单元格，保留公式/样式/列宽/条件格式。"""
    from openpyxl import load_workbook
    wb = load_workbook(TEMPLATE_XLSX)
    base = DOC.get("baseline") or []
    changed = 0
    for si, sh in enumerate(DOC["sheets"]):
        if si >= len(wb.worksheets):
            break
        ws = wb.worksheets[si]
        b = base[si]["rows"] if si < len(base) else []
        covered = set()
        for m in sh.get("merges", []):
            for r in range(m[0], m[2] + 1):
                for c in range(m[1], m[3] + 1):
                    if (r, c) != (m[0], m[1]):
                        covered.add((r, c))
        for ri, row in enumerate(sh["rows"]):
            for ci, val in enumerate(row):
                if (ri, ci) in covered:
                    continue
                bval = b[ri][ci] if (ri < len(b) and ci < len(b[ri])) else ""
                if str(val) == str(bval):
                    continue
                cell = ws.cell(row=ri + 1, column=ci + 1)
                sval = str(val).strip()
                if sval == "":
                    cell.value = None
                elif NUMERIC.match(sval):
                    try:
                        cell.value = (int(sval) if re.fullmatch(r"-?\d+", sval)
                                      else float(sval.replace(",", "")))
                    except Exception:
                        cell.value = str(val)
                else:
                    cell.value = str(val)
                changed += 1
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def writeback_docx() -> bytes:
    """加载原始 docx 模板，按表回写与基线不同的单元格，保留模板格式。"""
    from docx import Document
    doc = Document(TEMPLATE_DOCX)
    base = DOC.get("baseline") or []
    for si, sh in enumerate(DOC["sheets"]):
        if si >= len(doc.tables):
            break
        tbl = doc.tables[si]._tbl
        anchors, _, _, nrows, _ = _table_grid(tbl)
        b = base[si] if si < len(base) else {}
        brows = b.get("rows", [])
        mmap = b.get("merge_map", {})
        for ri, row in enumerate(sh["rows"]):
            if ri >= nrows:
                break  # 模板里没有这一行，无法回写
            for ci, val in enumerate(row):
                ar, ac = mmap.get(f"{ri}:{ci}", [ri, ci])
                tc = anchors.get((ar, ac))
                if tc is None:
                    continue
                bval = (brows[ri][ci] if ri < len(brows) and ci < len(brows[ri]) else "")
                if str(val) == str(bval):
                    continue
                _set_tc_text(tc, val)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ---------------------------------------------------------------- HTTP
class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    server_version = "OfflineSheet/1.0"

    def log_message(self, fmt, *args):
        if "--verbose" in os.sys.argv:
            super().log_message(fmt, *args)

    # -- 工具 --
    def _json(self, code: int, obj) -> None:
        body = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _file(self, code: int, body: bytes, ctype: str, filename: str) -> None:
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Content-Disposition",
                         f"attachment; filename*=UTF-8''{urllib.parse.quote(filename)}")
        self.end_headers()
        self.wfile.write(body)

    def _static(self, path: str) -> None:
        rel = path.lstrip("/")
        full = os.path.normpath(os.path.join(STATIC, rel))
        if not full.startswith(STATIC) or not os.path.isfile(full):
            self._json(404, {"error": "not found"})
            return
        ctype = {".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
                 ".css": "text/css; charset=utf-8", ".png": "image/png",
                 ".svg": "image/svg+xml", ".ico": "image/x-icon"}.get(
            os.path.splitext(full)[1], "application/octet-stream")
        with open(full, "rb") as f:
            body = f.read()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        self.wfile.write(body)

    def _token(self) -> Client | None:
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        token = (q.get("t") or [None])[0]
        if not token:
            return None
        return resolve(token)

    # -- GET --
    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        try:
            if path in ("/", "/index.html"):
                self._static("index.html")
            elif path.startswith("/static/"):
                self._static(path[len("/static/"):])
            elif path == "/api/doc":
                self._doc()
            elif path == "/api/events":
                self._events()
            elif path == "/api/export":
                self._export()
            elif path == "/api/ping":
                c = self._token()
                if c:
                    c.last = now()
                self._json(200, {"ok": bool(c)})
            else:
                self._json(404, {"error": "not found"})
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception:
            traceback.print_exc()
            try:
                self._json(500, {"error": "server error"})
            except Exception:
                pass

    def _doc(self):
        c = self._token()
        if not c:
            self._json(401, {"error": "未登录"})
            return
        c.last = now()
        self._json(200, {
            "title": DOC["title"], "version": DOC["version"],
            "sheets": DOC["sheets"], "me": c.roster(),
            "users": roster(), "owner": DOC["owner"],
        })

    def _events(self):
        c = self._token()
        if not c:
            self._json(401, {"error": "未登录"})
            return
        c.last = now()
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.send_header("Connection", "keep-alive")
        self.send_header("X-Accel-Buffering", "no")
        self.end_headers()
        self.wfile.write(f": hello {c.name}\n\n".encode())
        self.wfile.flush()
        push_roster()
        try:
            while True:
                try:
                    payload = c.q.get(timeout=15)
                    self.wfile.write(f"data: {json.dumps(payload, ensure_ascii=False)}\n\n".encode())
                except queue.Empty:
                    self.wfile.write(b": ping\n\n")
                self.wfile.flush()
                c.last = now()
        except (BrokenPipeError, ConnectionResetError, OSError):
            pass
        finally:
            CLIENTS.pop(c.token, None)
            push_roster()

    def _export(self):
        c = self._token()
        if not c:
            self._json(401, {"error": "未登录"})
            return
        if c.role != "owner":
            self._json(403, {"error": "只有上传人可以导出下载"})
            return
        q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        fmt = (q.get("fmt") or ["xlsx"])[0]
        stamp = time.strftime("%Y%m%d_%H%M")
        title = re.sub(r"[\\/:*?\"<>|]", "_", DOC["title"]) or "表格"
        tmpl = DOC.get("template")
        try:
            if fmt == "docx":
                body = writeback_docx() if tmpl == "docx" else state_to_docx()
                self._file(200, body,
                           "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           f"{title}_{stamp}.docx")
            else:
                body = writeback_xlsx() if tmpl == "xlsx" else state_to_xlsx()
                self._file(200, body,
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           f"{title}_{stamp}.xlsx")
        except Exception:
            traceback.print_exc()
            # 模板回写失败时退回按值重写，保证能拿到文件
            body = state_to_docx() if fmt == "docx" else state_to_xlsx()
            self._file(200, body,
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                       if fmt == "docx" else
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       f"{title}_{stamp}.{fmt}")
        broadcast({"type": "notice", "text": f"{c.name} 导出了文件", "level": "info"})

    # -- POST --
    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        length = int(self.headers.get("Content-Length") or 0)
        try:
            if path == "/api/login":
                self._login()
            elif path == "/api/upload":
                self._upload()
            elif path == "/api/edit":
                self._edit()
            elif path == "/api/leave":
                c = self._token()
                if c:
                    CLIENTS.pop(c.token, None)
                    push_roster()
                self._json(200, {"ok": True})
            else:
                self._json(404, {"error": "not found"})
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception:
            traceback.print_exc()
            try:
                self._json(500, {"error": "server error"})
            except Exception:
                pass

    def _body(self) -> bytes:
        n = int(self.headers.get("Content-Length") or 0)
        return self.rfile.read(n) if n else b""

    def _login(self):
        raw = self._body()
        try:
            data = json.loads(raw or b"{}")
        except Exception:
            data = {}
        name = (data.get("name") or "").strip()[:20] or "匿名"
        existing = (data.get("token") or "").strip()
        if existing:
            c = resolve(existing)
            if c:
                self._json(200, {"token": c.token, "name": c.name, "role": c.role,
                                 "color": c.color})
                return
        role = "owner" if DOC["owner"] is None else "guest"
        token = secrets.token_urlsafe(18)
        c = Client(token, name, role, color_for(name))
        with LOCK:
            if role == "owner":
                DOC["owner"] = token
            DOC["members"][token] = {"name": name, "role": role, "color": c.color}
            save_doc()
        CLIENTS[token] = c
        self._json(200, {"token": token, "name": name, "role": role, "color": c.color})
        push_roster()

    def _upload(self):
        c = self._token()
        if not c:
            self._json(401, {"error": "未登录"})
            return
        if c.role != "owner":
            self._json(403, {"error": "只有上传人可以导入/替换文档"})
            return
        blob = self._body()
        if not blob:
            self._json(400, {"error": "空文件"})
            return
        filename = urllib.parse.unquote(self.headers.get("X-Filename") or "upload.xlsx")
        name = os.path.basename(filename)
        ext = os.path.splitext(name)[1].lower()
        if ext == ".xlsx":
            kind, tpath = "xlsx", TEMPLATE_XLSX
        elif ext == ".docx":
            kind, tpath = "docx", TEMPLATE_DOCX
        else:
            self._json(400, {"error": "仅支持 .xlsx / .docx"})
            return
        try:
            sheets = xlsx_to_state(blob) if kind == "xlsx" else docx_to_state(blob)
        except Exception as e:
            traceback.print_exc()
            self._json(400, {"error": f"解析失败：{e}"})
            return
        baseline = []
        for sh in sheets:
            b = {"rows": [list(r) for r in sh["rows"]]}
            if kind == "docx":
                b["merge_map"] = sh.get("merge_map", {})
            baseline.append(b)
        with LOCK:
            try:
                tmp = tpath + ".tmp"
                with open(tmp, "wb") as f:
                    f.write(blob)
                os.replace(tmp, tpath)
                other = TEMPLATE_DOCX if kind == "xlsx" else TEMPLATE_XLSX
                if os.path.exists(other):
                    os.remove(other)
            except OSError:
                traceback.print_exc()
                self._json(500, {"error": "模板文件保存失败"})
                return
            DOC["sheets"] = sheets
            DOC["baseline"] = baseline
            DOC["template"] = kind
            DOC["title"] = os.path.splitext(name)[0]
            v = bump()
            save_doc()
        broadcast({"type": "reload", "version": v, "sheets": sheets, "title": DOC["title"]})
        self._json(200, {"ok": True, "version": v, "sheets": len(sheets)})

    def _edit(self):
        c = self._token()
        if not c:
            self._json(401, {"error": "未登录"})
            return
        try:
            op = json.loads(self._body() or b"{}")
        except Exception:
            self._json(400, {"error": "bad json"})
            return
        op["by"] = {"name": c.name, "color": c.color, "token": c.token}
        with LOCK:
            si = int(op.get("sheet", 0))
            if not (0 <= si < len(DOC["sheets"])):
                self._json(400, {"error": "sheet out of range"})
                return
            sh = DOC["sheets"][si]
            kind = op.get("op")
            applied = True
            if kind == "set":
                r, cc = int(op["r"]), int(op["c"])
                if not (0 <= r < MAX_ROWS and 0 <= cc < MAX_COLS):
                    self._json(400, {"error": "cell out of range"})
                    return
                while len(sh["rows"]) <= r:
                    sh["rows"].append(["" for _ in range(sh["cols"])])
                while len(sh["rows"][r]) <= cc:
                    sh["rows"][r].append("")
                sh["rows"][r][cc] = str(op.get("v", ""))[:2000]
            elif kind == "addrow":
                at = int(op.get("at", len(sh["rows"])))
                sh["rows"].insert(at, ["" for _ in range(sh["cols"])])
            elif kind == "addcol":
                at = int(op.get("at", sh["cols"]))
                for row in sh["rows"]:
                    row.insert(at, "")
                sh["cols"] += 1
            elif kind == "cursor":
                applied = False
            else:
                self._json(400, {"error": "unknown op"})
                return
            v = bump() if applied else DOC["version"]
            if applied and DOC["version"] % 5 == 0:
                save_doc()
        # 光标不改变文档内容（applied=False），但必须广播，否则别人看不到"谁正停在哪一格"
        op["type"] = "patch"
        op["version"] = v
        broadcast(op, exclude=None)
        self._json(200, {"ok": True, "version": v})

    # -- PUT：导入用裸字节，避免 multipart 解析 --
    def do_PUT(self):
        self.do_POST()

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Content-Length", "0")
        self.end_headers()


def main():
    ap = argparse.ArgumentParser(description="离线局域网 Excel 协同编辑")
    ap.add_argument("--port", type=int, default=8081)
    ap.add_argument("--host", default="0.0.0.0")
    args = ap.parse_args()
    httpd = ThreadingHTTPServer((args.host, args.port), Handler)
    httpd.daemon_threads = True
    ip = "127.0.0.1"
    ips = []
    try:
        import socket
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))  # 不发包，只取默认路由出口地址
        ips.append(s.getsockname()[0])
        s.close()
    except Exception:
        pass
    try:
        import socket
        hostname_ips = [a for a in socket.gethostbyname_ex(socket.gethostname())[2]
                        if not a.startswith("127.")]
        ips.extend(a for a in hostname_ips if a not in ips)
    except Exception:
        pass
    ip = ips[0] if ips else "127.0.0.1"
    print("=" * 56)
    print(" 离线协同表格已启动")
    print(f"   本机访问   http://127.0.0.1:{args.port}")
    if len(ips) > 1:
        for a in ips:
            print(f"   局域网访问 http://{a}:{args.port}")
    else:
        print(f"   局域网访问 http://{ip}:{args.port}")
    print(f"   数据文件   {DOC_PATH}")
    print("   第一个登录的人自动成为上传人（可导入/导出）")
    print("=" * 56, flush=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nbye")


if __name__ == "__main__":
    main()
