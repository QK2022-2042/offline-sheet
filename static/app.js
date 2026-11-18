/* 离线协同表格 — 前端（原生 JS，无任何外部依赖） */
'use strict';

const API = (p, opt) => fetch(p + (p.includes('?') ? '&' : '?') + 't=' + encodeURIComponent(S.token), opt);

const S = {
  token: localStorage.getItem('os_token') || '',
  me: null,
  doc: null,
  sheet: 0,
  es: null,
  cursors: {},          // token -> {sheet,r,c,color,name}
  pending: new Set(),
};

const $ = (id) => document.getElementById(id);
const colName = (i) => {
  let s = '';
  i += 1;
  while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - 1 - m) / 26; }
  return s;
};

/* ---------------------------------------------------------------- 登录 */
async function login(name) {
  const r = await fetch('/api/login', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, token: S.token }),
  });
  const j = await r.json();
  if (!r.ok) throw new Error(j.error || '登录失败');
  S.token = j.token;
  localStorage.setItem('os_token', j.token);
  S.me = j;
  return j;
}

async function boot() {
  if (!S.token) return showLogin();
  const r = await API('/api/doc');
  if (r.status === 401) { localStorage.removeItem('os_token'); S.token = ''; return showLogin(); }
  const doc = await r.json();
  S.doc = doc;
  S.me = doc.me;
  enterApp();
}

function showLogin() {
  $('login').hidden = false;
  $('app').hidden = true;
  $('nameInput').focus();
}

async function enterApp() {
  $('login').hidden = true;
  $('app').hidden = false;
  $('roleTag').textContent = S.me.role === 'owner' ? '上传人' : '协作者';
  $('roleTag').className = 'role ' + (S.me.role === 'owner' ? 'owner' : 'guest');
  const isOwner = S.me.role === 'owner';
  ['importBtn', 'exportXlsx', 'exportDocx'].forEach((id) => { $(id).hidden = !isOwner; });
  $('title').readOnly = !isOwner;
  document.body.classList.toggle('is-guest', !isOwner);
  buildWatermark();
  renderUsers(S.doc.users || []);
  render();
  connect();
  setInterval(() => API('/api/ping').catch(() => {}), 20000);
}

/* ---------------------------------------------------------------- 渲染 */
function render() {
  const doc = S.doc;
  if (!doc) return;
  $('ver').textContent = 'v' + doc.version;
  if (document.activeElement !== $('title')) $('title').value = doc.title;

  // tabs
  const tabs = $('tabs');
  tabs.innerHTML = '';
  doc.sheets.forEach((sh, i) => {
    const b = document.createElement('button');
    b.className = 'tab' + (i === S.sheet ? ' active' : '');
    b.textContent = sh.name;
    b.onclick = () => { S.sheet = i; render(); };
    tabs.appendChild(b);
  });

  const sh = doc.sheets[S.sheet];
  if (!sh) return;
  const covered = new Set();
  const spans = new Map();
  (sh.merges || []).forEach(([r1, c1, r2, c2]) => {
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        if (r === r1 && c === c1) continue;
        covered.add(r + ':' + c);
      }
    }
    spans.set(r1 + ':' + c1, [r2 - r1 + 1, c2 - c1 + 1]);
  });

  const rows = Math.min(Math.max(sh.rows.length, 20), 500);
  const cols = Math.min(Math.max(sh.cols, 8), 40);

  const table = $('grid');
  table.innerHTML = '';

  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  const corner = document.createElement('th');
  corner.className = 'corner';
  hr.appendChild(corner);
  for (let c = 0; c < cols; c++) {
    const th = document.createElement('th');
    th.textContent = colName(c);
    th.className = 'colhead';
    hr.appendChild(th);
  }
  thead.appendChild(hr);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  for (let r = 0; r < rows; r++) {
    const tr = document.createElement('tr');
    const rn = document.createElement('th');
    rn.className = 'rowhead';
    rn.textContent = r + 1;
    tr.appendChild(rn);
    for (let c = 0; c < cols; c++) {
      const key = r + ':' + c;
      if (covered.has(key)) continue;
      const td = document.createElement('td');
      td.dataset.r = r;
      td.dataset.c = c;
      const span = spans.get(key);
      if (span) { td.rowSpan = span[0]; td.colSpan = span[1]; td.classList.add('merged'); }
      const div = document.createElement('div');
      div.className = 'cell';
      div.contentEditable = 'true';
      div.spellcheck = false;
      div.dataset.r = r;
      div.dataset.c = c;
      const v = (sh.rows[r] && sh.rows[r][c] != null) ? sh.rows[r][c] : '';
      div.textContent = v;
      if (v !== '' && /^-?[\d,]+(\.\d+)?$/.test(String(v).trim())) div.classList.add('num');
      bindCell(div, r, c);
      td.appendChild(div);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  paintCursors();
}

function bindCell(el, r, c) {
  el.addEventListener('focus', () => {
    sendEdit({ op: 'cursor', sheet: S.sheet, r, c });
  });
  el.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); el.blur(); moveFocus(r, c, 0, 1); }
    else if (e.key === 'Tab') {
      e.preventDefault(); el.blur(); moveFocus(r, c, e.shiftKey ? -1 : 1, 0);
    } else if (e.key === 'Escape') { el.blur(); }
  });
  el.addEventListener('blur', () => {
    const v = (el.textContent || '').replace(/\u00a0/g, ' ').trim();
    const old = (S.doc.sheets[S.sheet].rows[r] || [])[c] ?? '';
    el.classList.toggle('num', v !== '' && /^-?[\d,]+(\.\d+)?$/.test(v));
    if (v === String(old)) return;
    setLocal(r, c, v);
    sendEdit({ op: 'set', sheet: S.sheet, r, c, v });
  });
}

function moveFocus(r, c, dc, dr) {
  const t = document.querySelector(`.cell[data-r="${r + dr}"][data-c="${c + dc}"]`);
  if (t) t.focus();
}

function setLocal(r, c, v) {
  const sh = S.doc.sheets[S.sheet];
  while (sh.rows.length <= r) sh.rows.push([]);
  while (sh.rows[r].length <= c) sh.rows[r].push('');
  sh.rows[r][c] = v;
}

/* ---------------------------------------------------------------- 网络 */
async function sendEdit(op) {
  try {
    const r = await API('/api/edit', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(op),
    });
    if (r.status === 401) { toast('登录已失效，请刷新页面', 'warn'); return; }
    const j = await r.json();
    if (j.version) {
      S.doc.version = j.version;
      $('ver').textContent = 'v' + j.version;
      $('saved').textContent = '已保存 ' + new Date().toLocaleTimeString('zh-CN', { hour12: false });
    }
  } catch (e) { toast('保存失败：' + e.message, 'warn'); }
}

let everConnected = false;

// 断线或漏包后重新拉一份完整文档，避免本地停留在旧内容
async function resync() {
  try {
    const r = await API('/api/doc');
    if (!r.ok) return;
    const doc = await r.json();
    S.doc = doc; S.me = doc.me;
    if (S.sheet >= doc.sheets.length) S.sheet = 0;
    renderUsers(doc.users || []);
    render();
    toast('已同步到最新内容', 'info');
  } catch (e) { /* 网络还没恢复，等下一次 */ }
}

function connect() {
  if (S.es) S.es.close();
  const es = new EventSource('/api/events?t=' + encodeURIComponent(S.token));
  S.es = es;
  es.onopen = () => {
    $('connDot').className = 'dot on';
    if (everConnected) resync();  // 重连后补齐断线期间漏掉的改动
    everConnected = true;
  };
  es.onerror = () => { $('connDot').className = 'dot off'; };
  es.onmessage = (ev) => {
    let m;
    try { m = JSON.parse(ev.data); } catch { return; }
    handle(m);
  };
}

function handle(m) {
  if (m.type === 'roster') return renderUsers(m.users);
  if (m.type === 'notice') return toast(m.text, m.level || 'info');
  if (m.type === 'reload') {
    S.doc.sheets = m.sheets; S.doc.title = m.title; S.doc.version = m.version;
    S.sheet = 0; render(); toast('上传人更新了文档', 'info'); return;
  }
  if (m.type === 'patch') {
    // 版本号跳号说明中间漏了消息（网络抖动/重连），直接整份重拉，避免两边内容悄悄不一致
    if (m.version && m.version > S.doc.version + 1) {
      S.doc.version = m.version - 1;
      resync();
      return;
    }
    if (m.by && m.by.token === S.token) return;
    const sh = S.doc.sheets[m.sheet];
    if (!sh) return;
    if (m.op === 'cursor') {
      S.cursors[m.by.token] = { sheet: m.sheet, r: m.r, c: m.c, color: m.by.color, name: m.by.name };
      paintCursors(); return;
    }
    if (m.op === 'set') {
      while (sh.rows.length <= m.r) sh.rows.push([]);
      while (sh.rows[m.r].length <= m.c) sh.rows[m.r].push('');
      sh.rows[m.r][m.c] = m.v;
      const el = document.querySelector(`.cell[data-r="${m.r}"][data-c="${m.c}"]`);
      if (el && document.activeElement !== el) {
        el.textContent = m.v;
        el.classList.toggle('num', m.v !== '' && /^-?[\d,]+(\.\d+)?$/.test(m.v));
        el.classList.add('flash');
        setTimeout(() => el.classList.remove('flash'), 700);
      }
    } else if (m.op === 'addrow' || m.op === 'addcol') {
      S.doc.version = m.version; render();
    }
    S.doc.version = m.version || S.doc.version;
    $('ver').textContent = 'v' + S.doc.version;
  }
}

function renderUsers(users) {
  const box = $('users');
  box.innerHTML = '';
  (users || []).forEach((u) => {
    const s = document.createElement('span');
    s.className = 'avatar';
    s.style.background = u.color;
    s.title = u.name + (u.role === 'owner' ? '（上传人）' : '');
    s.textContent = (u.name || '?').slice(0, 1);
    box.appendChild(s);
  });
}

function paintCursors() {
  document.querySelectorAll('.cell.cursor').forEach((el) => {
    el.classList.remove('cursor'); el.style.outlineColor = '';
  });
  Object.values(S.cursors).forEach((cu) => {
    if (cu.sheet !== S.sheet) return;
    const el = document.querySelector(`.cell[data-r="${cu.r}"][data-c="${cu.c}"]`);
    if (el) { el.classList.add('cursor'); el.style.outlineColor = cu.color; }
  });
}

/* ---------------------------------------------------------------- 交互 */
function toast(text, level) {
  const d = document.createElement('div');
  d.className = 'toast ' + (level || 'info');
  d.textContent = text;
  $('toasts').appendChild(d);
  setTimeout(() => d.remove(), 4000);
}

function buildWatermark() {
  const label = (S.me.name || '') + '  ' + new Date().toLocaleDateString('zh-CN');
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="320" height="180">
      <text x="20" y="100" font-size="16" fill="#000" fill-opacity="0.06"
        transform="rotate(-28 20 100)" font-family="sans-serif">${label.replace(/[<>&]/g, '')}</text>
    </svg>`;
  $('watermark').style.backgroundImage =
    'url("data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg) + '")';
}

function bindUI() {
  $('enterBtn').onclick = async () => {
    const name = $('nameInput').value.trim();
    if (!name) return $('nameInput').focus();
    try {
      await login(name);
      const doc = await (await API('/api/doc')).json();
      S.doc = doc; S.me = doc.me; enterApp();
    } catch (e) { toast(e.message, 'warn'); }
  };
  $('nameInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') $('enterBtn').click(); });

  document.querySelectorAll('[data-act]').forEach((b) => {
    b.onclick = () => {
      const sh = S.doc.sheets[S.sheet];
      const act = b.dataset.act;
      if (act === 'addrow') {
        const at = sh.rows.length;
        sh.rows.push([]);
        sendEdit({ op: 'addrow', sheet: S.sheet, at }).then(render);
        render();
      } else if (act === 'addcol') {
        sendEdit({ op: 'addcol', sheet: S.sheet, at: sh.cols });
        sh.cols += 1;
        render();
      }
    };
  });

  $('importBtn').onclick = () => $('fileInput').click();
  $('fileInput').onchange = async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    if (!/\.(xlsx|docx)$/i.test(f.name)) { toast('仅支持 .xlsx / .docx 文件', 'warn'); return; }
    if (!confirm(`导入「${f.name}」将替换当前所有内容，确定？`)) return;
    const buf = await f.arrayBuffer();
    const r = await fetch('/api/upload?t=' + encodeURIComponent(S.token), {
      method: 'PUT', headers: { 'X-Filename': encodeURIComponent(f.name) }, body: buf,
    });
    const j = await r.json();
    if (!r.ok) return toast(j.error || '导入失败', 'warn');
    toast('导入成功：' + j.sheets + ' 个工作表', 'info');
    $('fileInput').value = '';
  };

  const dl = (fmt) => { window.location.href = `/api/export?fmt=${fmt}&t=${encodeURIComponent(S.token)}`; };
  $('exportXlsx').onclick = () => dl('xlsx');
  $('exportDocx').onclick = () => dl('docx');

  // 只读保护（软性，防止误操作与他人顺手另存）
  document.addEventListener('contextmenu', (e) => {
    if (document.body.classList.contains('is-guest')) { e.preventDefault(); toast('编辑期间不允许下载/另存', 'warn'); }
  });
  document.addEventListener('keydown', (e) => {
    const k = (e.key || '').toLowerCase();
    if ((e.metaKey || e.ctrlKey) && ['s', 'p', 'u'].includes(k)) { e.preventDefault(); toast('编辑期间不允许下载/另存', 'warn'); }
  });
  document.addEventListener('dragstart', (e) => { if (e.target.closest('.grid')) e.preventDefault(); });
}

bindUI();
boot();
