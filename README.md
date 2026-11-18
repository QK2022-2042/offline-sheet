# 离线协同表格 · Offline Sheet

**断网局域网里，把 Excel / Word 表单变成网页，多人同时填写，导出回写原模板、格式不丢。**

Turn an Excel/Word form into a shared web page on an air-gapped LAN — everyone fills it in the browser, and only the uploader exports it back to the original template with styles, formulas and merges intact.

---

## ✨ 特性 / Features

- **零安装运行**：单文件 EXE（PyInstaller 打包）或绿色文件夹，目标机不需要 Python、不需要联网
- **Excel + Word 表单**：`.xlsx` 和表格型 `.docx`（含横向 / 纵向合并单元格）都能导入成网页表格
- **多人实时协同**：原生 JS + SSE，无任何前端依赖；他人编辑实时闪动提示、光标跟随
- **模板回写导出**：上传时留存原始文件与基线值，导出时只回写变更单元格——公式、字体、列宽、条件格式、表格样式全部保留
- **上传人权限模型**：第一位进入者自动成为上传人（唯一可导入 / 导出者），其他协作者仅可编辑，导出接口返回 403
- **防误传播**：协作者页面无下载入口、拦截 Ctrl+S/Ctrl+P、铺姓名水印（防截屏拍照则超出软件能力）

## 🚀 快速开始 / Quick Start

### 方式一：源码运行（本机有 Python 3.9+）

```bash
pip install openpyxl python-docx
python server.py --port 8081
```

按提示把 `http://<本机IP>:8081` 发到局域网群里即可。

### 方式二：打包绿色文件夹（目标机零依赖）

```bash
# 1. 准备 build/python-embed.zip（python-3.11.9-embed-amd64）与 build/wheels/（离线轮子）
python scripts/build_green_package.py
# 2. 进入 dist/offline-sheet，双击 start.bat
```

### 方式三：打包单文件 EXE

```bash
pip install pyinstaller
pyinstaller --onefile --add-data "static;static" --name OfflineSheet server.py
```

> Windows 10/11 64 位。换端口：EXE 加参数 `--port 8082`，绿色包在目录内建 `port.txt`。

## 📖 使用流程

1. **上传人**打开网页、输入姓名（第一个进入者自动成为上传人）→「导入 Excel / Word」选文件
2. 其他人浏览器打开同一地址，各自填写；改动实时同步（对方格子闪动、彩色光标）
3. 定稿后**只有上传人**能点「导出 Excel / 导出 Word」，下载的文件回写原模板、格式一致

## 🔧 工作原理

```
上传 .xlsx/.docx ──► 解析成网格（值 + 合并单元格）──► data/doc.json（实时状态）
        │                                             ▲
        └─ 原始文件留存 data/template.* + 记录基线值      │ SSE 实时广播
                                                      │
导出 ◄── 加载原模板，仅回写「与基线不同的单元格」◄── 浏览器编辑
```

- 合并单元格：Excel 按 openpyxl merge range 展示；Word 按 `w:gridSpan` / `w:vMerge` 展开成规则网格并记录锚点映射，导出按锚点回写
- 断线容错：SSE 漏包自动整份重拉；数据每 5 次修改原子落盘，重启不丢

## ⚠️ 已知边界

| 事项 | 说明 |
|---|---|
| Excel 导出 | 公式、样式、列宽、条件格式、合并区保留；图表和图片不保留（openpyxl 限制） |
| Word | 仅支持表格型表单；正文段落不出现在网页里（导出原样保留）；表格内新增行无法回写 |
| 公式显示 | 导入时公式显示为 Excel 缓存值（真实 Excel 文件均有缓存） |
| 并发策略 | 单元格级后写覆盖，适合分区编辑 |
| 大表性能 | 建议单表 ≤ 2000 行 × 40 列 |

## ✅ 测试

```bash
python scripts/test_e2e.py        # 20 项端到端（导入/编辑/回写/权限），需先启动服务
python scripts/test_exe_deploy.py # EXE 部署测试：全量 e2e + 重启持久化
```

当前记录：20 项端到端 + EXE 真实部署（含停服重启持久化）全部通过。

## 📁 目录结构

```
├── server.py               服务端（HTTP + SSE + 导入导出 + 模板回写）
├── static/                 前端（原生 JS，无依赖）
├── scripts/
│   ├── build_green_package.py   组装绿色文件夹（内嵌便携 Python）
│   ├── test_e2e.py              端到端测试
│   └── test_exe_deploy.py       EXE 部署测试
├── docs/使用说明.txt        面向使用者的简明手册
└── README.md
```

## 📄 License

[MIT](LICENSE)
