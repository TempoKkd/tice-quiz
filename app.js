/* =====================================================================
 * 题册 · 表格导入刷题题库
 * 纯前端实现：Excel/CSV 导入 → localStorage 持久化 → 练习/考试/错题本
 * ===================================================================== */

'use strict';

/* ---------------- 持久化 ---------------- */
const LS = { q: 'tice_questions_v1', w: 'tice_wrong_v1', r: 'tice_records_v1' };
function load(k, d) { try { const v = JSON.parse(localStorage.getItem(k)); return v ?? d; } catch (e) { return d; } }
function save(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* 存储满等异常忽略 */ } }

let questions = load(LS.q, []);                 // 题库
let wrongBook = load(LS.w, {});                 // 错题本: id -> {c:错误次数, t:最近时间}
let records = load(LS.r, { answered: 0, correct: 0, lastImport: null });

/* ---------------- 常量与工具 ---------------- */
const TYPE = {
  single:   { name: '单选题', short: '单选' },
  multiple: { name: '多选题', short: '多选' },
  judge:    { name: '判断题', short: '判断' },
};
const OPT_KEYS = 'ABCDEFGH'.split('');

const esc = s => String(s ?? '').replace(/[&<>"']/g,
  m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));
const shuffle = a => { for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
const eqSet = (set, arr) => set.size === arr.length && arr.every(x => set.has(x));
const uid = i => 'q' + Date.now().toString(36) + '_' + i + '_' + Math.random().toString(36).slice(2, 6);
const fmtMS = ms => {
  const s = Math.max(0, Math.round(ms / 1000));
  const m = Math.floor(s / 60), ss = s % 60;
  return (m < 10 ? '0' + m : m) + ':' + (ss < 10 ? '0' + ss : ss);
};
const fmtDur = ms => {
  const s = Math.round(ms / 1000);
  if (s < 60) return s + ' 秒';
  return Math.floor(s / 60) + ' 分 ' + (s % 60) + ' 秒';
};

function answerText(q) { return q.type === 'judge' ? (q.answer[0] === 'A' ? '正确' : '错误') : q.answer.join('、'); }
function pickedText(q, set) {
  if (!set || !set.size) return '未作答';
  if (q.type === 'judge') { const v = [...set][0]; return v === 'A' ? '正确' : v === 'B' ? '错误' : '未作答'; }
  return [...set].sort().join('、');
}
function optChip(q, key) { return q.type === 'judge' ? (key === 'A' ? '✓' : '✗') : key; }
function fingerprint(q) { return q.type + '|' + q.question + '|' + q.answer.join(''); }

/* ---------------- Toast 与 弹窗 ---------------- */
const toastRoot = () => document.getElementById('toastRoot');
function toast(msg, kind = '') {
  const el = document.createElement('div');
  el.className = 'toast ' + kind;
  el.textContent = msg;
  toastRoot().appendChild(el);
  setTimeout(() => { el.style.opacity = '0'; el.style.transition = 'opacity .3s'; setTimeout(() => el.remove(), 320); }, 2400);
}

function openModal(html) {
  const root = document.getElementById('modalRoot');
  root.innerHTML = `<div class="modal-mask"><div class="modal" role="dialog" aria-modal="true">${html}</div></div>`;
  root.querySelector('.modal-mask').addEventListener('click', e => {
    if (e.target.classList.contains('modal-mask')) closeModal();
  });
}
function closeModal() { document.getElementById('modalRoot').innerHTML = ''; }

function confirmDlg({ title, body, okText = '确定', danger = false }) {
  return new Promise(resolve => {
    openModal(`
      <div class="modal-title">${esc(title)}</div>
      <div class="modal-body">${body}</div>
      <div class="modal-actions">
        <button class="btn btn-ghost" data-m="no">取消</button>
        <button class="btn ${danger ? 'btn-danger' : ''}" data-m="yes">${esc(okText)}</button>
      </div>`);
    document.getElementById('modalRoot').onclick = e => {
      const m = e.target.dataset.m;
      if (m === 'yes') { closeModal(); resolve(true); }
      if (m === 'no') { closeModal(); resolve(false); }
      if (e.target.classList.contains('modal-mask')) { closeModal(); resolve(false); }
    };
  });
}

/* =====================================================================
 * 表格解析：CSV / Excel → 题目对象
 * 模板列: 题型 | 题目 | 选项A | 选项B | ... | 选项H | 答案 | 解析
 * ===================================================================== */

const HEADER_RULES = [
  [/^(题型|类型|题目类型)$/, 'type'],
  [/^(题目|题干|问题|题目内容)$/, 'question'],
  ...OPT_KEYS.map(k => [new RegExp('^(选项)?' + k + '$', 'i'), k]),
  [/^(答案|正确答案|参考答案)$/, 'answer'],
  [/^(解析|答案解析|说明|解释)$/, 'analysis'],
];
function matchHeader(cell) {
  const c = String(cell || '').trim();
  for (const [re, key] of HEADER_RULES) if (re.test(c)) return key;
  return null;
}

function parseCSV(text) {
  const rows = []; let row = [], field = '', inQ = false;
  text = String(text).replace(/^\uFEFF/, '');
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false; }
      else field += c;
    } else {
      if (c === '"') inQ = true;
      else if (c === ',') { row.push(field); field = ''; }
      else if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
      else if (c !== '\r') field += c;
    }
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows.filter(r => r.some(cell => String(cell).trim() !== ''));
}

async function readTextSmart(file) {
  const buf = await file.arrayBuffer();
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); }
  catch (e) { return new TextDecoder('gbk').decode(buf); } // 兼容国内 Excel 导出的 GBK CSV
}

function parseType(s) {
  s = String(s || '').trim();
  if (/单选/.test(s) || /^single$/i.test(s)) return 'single';
  if (/多选/.test(s) || /^multiple$/i.test(s)) return 'multiple';
  if (/判断/.test(s) || /^(judge|tf)$/i.test(s)) return 'judge';
  return null;
}
function normJudgeAnswer(s) {
  s = String(s || '').trim().toLowerCase();
  if (['a', '对', '正确', 'true', 't', '√', 'yes', 'y', '1'].includes(s)) return 'A';
  if (['b', '错', '错误', 'false', 'f', '×', 'x', 'no', 'n', '0'].includes(s)) return 'B';
  return null;
}

/* 解析一行 → {q} 或 {error} */
function parseRow(get) {
  const stem = String(get('question') ?? '').trim();
  if (!stem) return { error: '题干为空' };
  const typeRaw = String(get('type') ?? '').trim();
  const ansRaw = String(get('answer') ?? '').trim();
  const analysis = String(get('analysis') ?? '').trim();
  let type = parseType(typeRaw);

  const opts = [];
  for (const k of OPT_KEYS) {
    const v = String(get(k) ?? '').trim();
    if (v) opts.push({ key: k, text: v });
  }

  // 判断题：无选项，或题型明确为判断
  if (type === 'judge' || (type === null && opts.length === 0)) {
    const a = normJudgeAnswer(ansRaw);
    if (!a) return { error: `判断题答案「${ansRaw || '空'}」无法识别，应填 对/错（或 正确/错误）` };
    return { q: { type: 'judge', question: stem, options: [{ key: 'A', text: '正确' }, { key: 'B', text: '错误' }], answer: [a], analysis } };
  }

  // 选择题
  if (opts.length < 2) return { error: '选项不足，选择题至少需要 2 个选项' };
  const letters = [...new Set((ansRaw.toUpperCase().match(/[A-H]/g) || []))].sort();
  if (!letters.length) return { error: '答案为空或格式不正确（应填选项字母，如 A 或 ABD）' };
  const keys = opts.map(o => o.key);
  const out = letters.filter(l => !keys.includes(l));
  if (out.length) return { error: `答案 ${out.join('')} 超出选项范围（仅 ${keys[0]}~${keys[keys.length - 1]}）` };
  if (!type) type = letters.length > 1 ? 'multiple' : 'single';
  if (type === 'single' && letters.length !== 1) return { error: `单选题答案应为 1 个字母，当前为 ${letters.join('')}` };
  if (type === 'multiple' && letters.length < 2) return { error: '多选题答案应至少包含 2 个字母' };
  return { q: { type, question: stem, options: opts, answer: letters, analysis } };
}

/* 二维数组 → 解析结果。返回 {oks, dupBank, dupFile, errs} */
function rowsToQuestions(rows, bankFP) {
  const res = { oks: [], dupBank: 0, dupFile: 0, errs: [] };
  const fileFP = new Set();
  const first = rows[0].map(c => String(c ?? '').trim());
  const map = {};
  first.forEach((cell, i) => { const k = matchHeader(cell); if (k && map[k] === undefined) map[k] = i; });

  let get, start;
  if (map.question !== undefined) {           // 有表头
    get = (r, k) => (map[k] === undefined ? '' : r[map[k]]);
    start = 1;
  } else {                                     // 无表头：按模板列顺序
    const pos = { type: 0, question: 1, answer: 10, analysis: 11 };
    OPT_KEYS.forEach((k, i) => pos[k] = 2 + i);
    get = (r, k) => r[pos[k]];
    start = 0;
  }

  rows.slice(start).forEach((r, idx) => {
    const line = idx + start + 1;
    const out = parseRow(k => get(r, k));
    if (out.error) { res.errs.push({ line, reason: out.error }); return; }
    const fp = fingerprint(out.q);
    if (fileFP.has(fp)) { res.dupFile++; return; }
    fileFP.add(fp);
    if (bankFP.has(fp)) { res.dupBank++; out.q._dupBank = true; res.oks.push(out.q); return; }
    res.oks.push(out.q);
  });
  return res;
}

async function handleFile(file) {
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  let rows;
  try {
    if (ext === 'csv' || ext === 'txt') {
      rows = parseCSV(await readTextSmart(file));
    } else if (ext === 'xlsx' || ext === 'xls') {
      if (!window.XLSX) { toast('Excel 解析组件未加载成功，请联网刷新，或改用 CSV 文件', 'err'); return; }
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    } else {
      toast('仅支持 .xlsx / .xls / .csv 格式的表格文件', 'err'); return;
    }
  } catch (e) { toast('文件读取失败：' + e.message, 'err'); return; }

  rows = (rows || []).filter(r => r && r.some(c => String(c).trim() !== ''));
  if (!rows.length) { toast('未在文件中读到有效数据行', 'err'); return; }
  showImportPreview(file.name, rowsToQuestions(rows, new Set(questions.map(fingerprint))));
}

function showImportPreview(fname, res) {
  const fresh = res.oks.filter(q => !q._dupBank);
  const dist = t => res.oks.filter(q => q.type === t).length;
  if (!res.oks.length && !res.errs.length) { toast('文件中没有可导入的题目', 'err'); return; }
  const hasBank = questions.length > 0;
  openModal(`
    <div class="modal-title">导入预览</div>
    <div class="modal-body">
      文件 <b>${esc(fname)}</b> 解析完成：
      <div class="imp-stats">
        <span class="imp-stat is-ok">可导入 ${fresh.length} 题</span>
        ${res.dupBank ? `<span class="imp-stat is-dup">与现有题库重复 ${res.dupBank} 题</span>` : ''}
        ${res.dupFile ? `<span class="imp-stat is-dup">文件内重复 ${res.dupFile} 题</span>` : ''}
        ${res.errs.length ? `<span class="imp-stat is-err">无法识别 ${res.errs.length} 行</span>` : ''}
      </div>
      <div style="font-size:13px">题型分布：单选 ${dist('single')} · 多选 ${dist('multiple')} · 判断 ${dist('judge')}</div>
      ${res.errs.length ? `<div class="imp-errs">${res.errs.slice(0, 8).map(e =>
        `<div><b>第${e.line}行</b>${esc(e.reason)}</div>`).join('')}
        ${res.errs.length > 8 ? `<div>… 共 ${res.errs.length} 行存在问题</div>` : ''}</div>` : ''}
    </div>
    <div class="modal-actions">
      <button class="btn btn-ghost" data-m="cancel">取消</button>
      ${hasBank && res.dupBank + fresh.length ? `<button class="btn btn-ghost" data-m="replace">替换现有题库（${res.oks.length} 题）</button>` : ''}
      ${fresh.length ? `<button class="btn btn-accent" data-m="append">${hasBank ? `追加导入 ${fresh.length} 题` : `导入 ${fresh.length} 题`}</button>` : ''}
    </div>`);
  document.getElementById('modalRoot').onclick = e => {
    const m = e.target.dataset.m;
    if (m === 'append') { doImport(fresh); closeModal(); }
    else if (m === 'replace') { doImport(res.oks, true); closeModal(); }
    else if (m === 'cancel' || e.target.classList.contains('modal-mask')) closeModal();
  };
}

function doImport(list, replace = false) {
  if (replace) { questions = []; wrongBook = {}; save(LS.w, wrongBook); }
  list.forEach((q, i) => { delete q._dupBank; q.id = uid(i); questions.push(q); });
  save(LS.q, questions);
  records.lastImport = { time: Date.now(), added: list.length, total: questions.length };
  save(LS.r, records);
  toast(`成功导入 ${list.length} 道题目`, 'ok');
  render();
}

/* ---------------- 模板下载 & 示例 ---------------- */
const TPL_HEAD = ['题型', '题目', '选项A', '选项B', '选项C', '选项D', '选项E', '选项F', '答案', '解析'];
const SAMPLE_ROWS = [
  ['单选', '中国的首都是哪里？', '上海', '北京', '广州', '深圳', '', '', 'B', '北京是中华人民共和国的首都。'],
  ['单选', '下列哪个属于操作系统？', 'Photoshop', 'Windows 11', 'Chrome', 'Excel', '', '', 'B', ''],
  ['单选', '水的化学式是？', 'CO₂', 'O₂', 'H₂O', 'NaCl', '', '', 'C', ''],
  ['单选', 'HTTP 协议的默认端口是？', '21', '443', '80', '8080', '', '', 'C', 'HTTP 默认端口为 80，HTTPS 为 443。'],
  ['多选', '下列属于编程语言的有？', 'Python', 'HTML', 'Java', 'CSS', '', '', 'AC', 'HTML 是标记语言，CSS 是样式表语言，二者均非编程语言。'],
  ['多选', '下列属于中国四大名著的有？', '《红楼梦》', '《聊斋志异》', '《西游记》', '《水浒传》', '', '', 'ACD', '四大名著为《红楼梦》《西游记》《水浒传》《三国演义》。'],
  ['多选', '下列哪些属于可再生能源？', '太阳能', '煤炭', '风能', '石油', '', '', 'AC', ''],
  ['判断', '地球是太阳系中距离太阳第三近的行星。', '', '', '', '', '', '', '对', ''],
  ['判断', '光在真空中的传播速度约为每秒 30 万公里。', '', '', '', '', '', '', '对', ''],
  ['判断', '鲸鱼属于鱼类。', '', '', '', '', '', '', '错', '鲸鱼是哺乳动物，用肺呼吸。'],
];

function downloadTemplate() {
  const rows = [TPL_HEAD, ...SAMPLE_ROWS.slice(0, 3), SAMPLE_ROWS[4], SAMPLE_ROWS[7]];
  const csv = '﻿' + rows.map(r => r.map(c => {
    c = String(c);
    return /[",\n]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c;
  }).join(',')).join('\r\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  a.download = '题库导入模板.csv';
  a.click();
  URL.revokeObjectURL(a.href);
  toast('模板已下载，用 Excel 或 WPS 打开即可编辑', 'ok');
}

function loadSample() {
  const res = rowsToQuestions([TPL_HEAD, ...SAMPLE_ROWS], new Set(questions.map(fingerprint)));
  const fresh = res.oks.filter(q => !q._dupBank);
  if (!fresh.length) { toast('示例题目已全部在题库中'); return; }
  doImport(fresh);
}

/* =====================================================================
 * 答题会话（练习 / 考试 / 错题复习共用引擎）
 * ===================================================================== */
let session = null;   // {mode, list, idx, results[], start, timer, deadline, active, finished}
let currentView = 'home';
const setupState = {
  practice: { type: 'all', order: 'seq', count: 'all' },
  exam: { count: 20, mins: 30 },
};

function buildList({ type = 'all', order = 'seq', count = 'all' }) {
  let list = questions.filter(q => type === 'all' || q.type === type);
  if (order === 'rand') list = shuffle([...list]);
  if (count !== 'all') list = list.slice(0, Math.min(count, list.length));
  return list;
}

function startSession(mode, list, mins = 0) {
  session = {
    mode, list, idx: 0,
    results: new Array(list.length),
    start: Date.now(), deadline: 0, timer: null,
    active: true, finished: false,
  };
  if (mode === 'exam' && mins > 0) {
    session.deadline = Date.now() + mins * 60000;
    session.timer = setInterval(tickTimer, 500);
  }
  currentView = 'quiz';
  render();
}

function tickTimer() {
  const el = document.getElementById('examTimer');
  const left = session ? session.deadline - Date.now() : 0;
  if (!session || left <= 0) {
    if (session && session.active) { endTimer(); toast('时间到，已自动交卷'); finishExam(); }
    return;
  }
  if (el) {
    el.textContent = '剩余 ' + fmtMS(left);
    el.classList.toggle('urgent', left < 60000);
  }
}
function endTimer() { if (session && session.timer) { clearInterval(session.timer); session.timer = null; } }
function endSession() { endTimer(); session = null; }

const curResult = () => session.results[session.idx] ??= { picked: new Set(), submitted: false };

function pickOption(i) {
  if (!session || session.finished) return;
  const q = session.list[session.idx];
  const r = curResult();
  const key = q.options[i].key;

  if (session.mode === 'exam') {
    if (q.type === 'multiple') { r.picked.has(key) ? r.picked.delete(key) : r.picked.add(key); render(); }
    else {
      r.picked = new Set([key]); render();
      if (session.idx < session.list.length - 1) setTimeout(() => { if (session && !session.finished) { session.idx++; render(); } }, 220);
    }
    return;
  }
  // 练习 / 错题复习：提交后不可再改
  if (r.submitted) return;
  if (q.type === 'multiple') {
    r.picked.has(key) ? r.picked.delete(key) : r.picked.add(key);
    render();
  } else {
    r.picked = new Set([key]);
    submitPractice();
  }
}

function submitPractice() {
  const s = session, r = curResult();
  if (r.submitted) return;
  if (!r.picked.size) { toast('请先选择答案'); return; }
  r.submitted = true;
  const q = s.list[s.idx];
  r.correct = eqSet(r.picked, q.answer);
  records.answered++; if (r.correct) records.correct++;
  save(LS.r, records);
  if (r.correct) {
    if (s.mode === 'wrong' && wrongBook[q.id]) { delete wrongBook[q.id]; save(LS.w, wrongBook); r.removed = true; }
  } else {
    (wrongBook[q.id] ??= { c: 0 }).c++;
    wrongBook[q.id].t = Date.now();
    save(LS.w, wrongBook);
  }
  render();
}

function practiceNext() {
  if (session.idx < session.list.length - 1) { session.idx++; render(); }
  else finishPractice();
}
function finishPractice() {
  session.finished = true; session.active = false; endTimer();
  currentView = 'result'; render();
}

async function submitExamClick() {
  let empty = 0;
  for (let i = 0; i < session.list.length; i++) {
    const r = session.results[i];
    if (!r || !r.picked.size) empty++;
  }
  const ok = await confirmDlg({
    title: '确认交卷？',
    body: empty ? `还有 <b>${empty}</b> 道题未作答，未作答按答错计分。` : '全部题目均已作答，交卷后不可修改。',
    okText: '交卷',
  });
  if (ok) finishExam();
}
function finishExam() {
  if (!session || session.finished) return;
  endTimer();
  session.list.forEach((q, i) => {
    const r = session.results[i] ??= { picked: new Set(), submitted: true };
    r.submitted = true;
    r.correct = r.picked.size > 0 && eqSet(r.picked, q.answer);
    if (r.picked.size) { records.answered++; if (r.correct) records.correct++; }
    if (!r.correct) { (wrongBook[q.id] ??= { c: 0 }).c++; wrongBook[q.id].t = Date.now(); }
  });
  save(LS.r, records); save(LS.w, wrongBook);
  session.finished = true; session.active = false;
  session.used = Date.now() - session.start;
  currentView = 'result'; render();
}

/* =====================================================================
 * 视图渲染
 * ===================================================================== */
const view = () => document.getElementById('view');
const wrongCount = () => Object.keys(wrongBook).length;
const byType = t => questions.filter(q => q.type === t).length;

const ICONS = {
  upload: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 16V4"/><path d="m6 10 6-6 6 6"/><path d="M4 20h16"/></svg>',
  pen: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>',
  timer: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="13" r="8"/><path d="M12 9v4l2.5 2.5"/><path d="M9 2h6"/></svg>',
  book: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20V4a2 2 0 0 0-2-2H6.5A2.5 2.5 0 0 0 4 4.5Z"/><path d="m9.5 8 5 5M14.5 8l-5 5"/></svg>',
  trash: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6M14 11v6"/></svg>',
  search: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="m21 21-4-4"/></svg>',
  list: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M8 6h13M8 12h13M8 18h13"/><circle cx="4" cy="6" r="1.2" fill="currentColor" stroke="none"/><circle cx="4" cy="12" r="1.2" fill="currentColor" stroke="none"/><circle cx="4" cy="18" r="1.2" fill="currentColor" stroke="none"/></svg>',
  spark: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3v3M12 18v3M3 12h3M18 12h3M5.6 5.6l2.2 2.2M16.2 16.2l2.2 2.2M18.4 5.6l-2.2 2.2M7.8 16.2l-2.2 2.2"/></svg>',
  read: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M2 4h6a4 4 0 0 1 4 4v13a3 3 0 0 0-3-3H2Z"/><path d="M22 4h-6a4 4 0 0 0-4 4v13a3 3 0 0 1 3-3h7Z"/></svg>',
  filter: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M4 5h16l-6 8v5l-4 2v-7Z"/></svg>',
  sort: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M7 4v16M7 20l-3-3M7 20l3-3M17 20V4M17 4l-3 3M17 4l3 3"/></svg>',
  hash: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M9 3 7 21M17 3l-2 18M4 8h17M3 16h17"/></svg>',
};

const badge = t => `<span class="badge badge-${t}">${TYPE[t].short}</span>`;

function emptyState(title, desc, actions = '') {
  return `<div class="card empty anim-in">
    <div class="empty-mark">${ICONS.spark}</div>
    <div class="empty-title">${title}</div>
    <div class="empty-desc">${desc}</div>
    <div class="empty-actions">${actions}</div>
  </div>`;
}

/* ---------------- 概览 ---------------- */
function renderHome() {
  const n = questions.length;
  const acc = records.answered ? Math.round(records.correct / records.answered * 100) : 0;
  const li = records.lastImport;
  const p = t => n ? (byType(t) / n * 100) : 0;

  view().innerHTML = `
    <div class="page-head">
      <div class="page-title">今天也要好好刷题</div>
      <div class="page-sub">${li ? `最近导入：${new Date(li.time).toLocaleString('zh-CN')}，新增 ${li.added} 题` : '从导入一张表格开始，建立你的专属题库'}</div>
    </div>

    <div class="stat-grid">
      <div class="card stat"><div class="stat-label">${ICONS.list} 题库总量</div><div class="stat-num">${n}</div><div class="stat-hint">单选 ${byType('single')} · 多选 ${byType('multiple')} · 判断 ${byType('judge')}</div></div>
      <div class="card stat"><div class="stat-label">${ICONS.book} 待复习错题</div><div class="stat-num" style="color:${wrongCount() ? 'var(--red)' : 'inherit'}">${wrongCount()}</div><div class="stat-hint">${wrongCount() ? '答对一次即可移出错题本' : '保持住，暂无错题'}</div></div>
      <div class="card stat"><div class="stat-label">${ICONS.pen} 累计作答</div><div class="stat-num">${records.answered}</div><div class="stat-hint">答对 ${records.correct} 题</div></div>
      <div class="card stat"><div class="stat-label">${ICONS.spark} 正确率</div><div class="stat-num">${records.answered ? acc + '%' : '—'}</div><div class="stat-hint">${records.answered ? '继续加油' : '还没有作答记录'}</div></div>
    </div>

    ${n ? `<div class="card type-bar-wrap">
      <div style="font-weight:700;font-size:15px">题型构成</div>
      <div class="type-bar"><span class="tb-single" style="width:${p('single')}%"></span><span class="tb-multiple" style="width:${p('multiple')}%"></span><span class="tb-judge" style="width:${p('judge')}%"></span></div>
      <div class="type-legend"><span><i class="tb-single"></i>单选 ${byType('single')}</span><span><i class="tb-multiple"></i>多选 ${byType('multiple')}</span><span><i class="tb-judge"></i>判断 ${byType('judge')}</span></div>
    </div>` : ''}

    ${n ? `<div class="action-grid">
      <button class="card action-card" data-action="goto" data-view="practice"><div class="action-icon ai-blue">${ICONS.pen}</div><div class="action-name">开始练习</div><div class="action-desc">逐题作答，即时反馈答案与解析</div></button>
      <button class="card action-card" data-action="goto" data-view="recite"><div class="action-icon ai-green">${ICONS.read}</div><div class="action-name">背题模式</div><div class="action-desc">答案直接展示，快速通览记忆</div></button>
      <button class="card action-card" data-action="goto" data-view="exam"><div class="action-icon ai-amber">${ICONS.timer}</div><div class="action-name">模拟考试</div><div class="action-desc">限时作答，交卷后统一评分</div></button>
      <button class="card action-card" data-action="goto" data-view="wrong"><div class="action-icon ai-red">${ICONS.book}</div><div class="action-name">错题复习</div><div class="action-desc">${wrongCount() ? `还有 ${wrongCount()} 道错题待攻克` : '当前没有错题'}</div></button>
      <button class="card action-card" data-action="goto" data-view="import"><div class="action-icon ai-teal">${ICONS.upload}</div><div class="action-name">继续导入</div><div class="action-desc">追加新的表格题目到题库</div></button>
    </div>`
    : emptyState('题库还是空的', '支持 Excel（.xlsx / .xls）与 CSV 表格导入，也可以先加载示例题体验',
        `<button class="btn btn-accent" data-action="goto" data-view="import">${ICONS.upload} 导入表格</button>
         <button class="btn btn-ghost" data-action="load-sample">加载示例题目</button>`)}
  `;
}

/* ---------------- 导入页 ---------------- */
function renderImport() {
  view().innerHTML = `
    <div class="page-head">
      <div class="page-title">导入题目</div>
      <div class="page-sub">把整理好的题目表格（Excel / CSV）拖进来，即可开始刷题</div>
    </div>

    <div class="dropzone" id="dz">
      <div class="dz-icon">${ICONS.upload}</div>
      <div class="dz-title">点击选择，或将表格文件拖到这里</div>
      <div class="dz-sub">支持 .xlsx / .xls / .csv，单文件建议不超过 5000 行</div>
      <input type="file" id="fileInput" accept=".csv,.txt,.xlsx,.xls" hidden>
    </div>

    <div class="imp-head" style="margin-top:18px">
      <button class="btn btn-ghost" data-action="tpl-download">下载导入模板（CSV）</button>
      <button class="btn btn-ghost" data-action="load-sample">加载示例题目</button>
    </div>

    <div class="card fmt-note">
      <h3>表格格式说明</h3>
      <table class="fmt-table">
        <tr><th>列名</th><th>说明</th></tr>
        <tr><td><b>题型</b></td><td>填「单选」「多选」或「判断」；留空时按答案和选项自动判断</td></tr>
        <tr><td><b>题目</b></td><td>题干，必填</td></tr>
        <tr><td><b>选项A ~ 选项H</b></td><td>选择题至少填 2 个选项；判断题留空即可</td></tr>
        <tr><td><b>答案</b></td><td>单选填 <code>A</code>；多选填 <code>ABD</code>；判断填 <code>对</code> 或 <code>错</code></td></tr>
        <tr><td><b>解析</b></td><td>选填，答错后展示，帮助记忆</td></tr>
      </table>
      <div class="fmt-tips">
        · 第一行必须是表头（列名见上表），也兼容无表头按模板列顺序排列的文件<br>
        · 导入时会自动跳过：重复题目、题干为空、答案超出选项范围等异常行，并在预览中列出<br>
        · Excel 用户建议「另存为 → CSV UTF-8」以获得最佳兼容性；直接上传 .xlsx 也可以
      </div>
    </div>
  `;
  const dz = document.getElementById('dz');
  const fi = document.getElementById('fileInput');
  dz.addEventListener('click', () => fi.click());
  fi.addEventListener('change', () => { if (fi.files[0]) handleFile(fi.files[0]); fi.value = ''; });
  ['dragover', 'dragenter'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove('drag'); }));
  dz.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) handleFile(f); });
}

/* ---------------- 练习 / 考试设置 ---------------- */
function chipGroup(group, key, options) {
  return `<div class="chips">${options.map(([v, label]) =>
    `<button class="chip ${setupState[group][key] === v ? 'on' : ''}" data-action="chip" data-group="${group}" data-key="${key}" data-val="${v}">${label}</button>`).join('')}</div>`;
}

function renderPracticeSetup() {
  const st = setupState.practice;
  const list = buildList(st);
  view().innerHTML = `
    <div class="page-head"><div class="page-title">开始练习</div><div class="page-sub">逐题作答、即时反馈，答错的题会自动收入错题本</div></div>
    ${questions.length ? `<div class="setup">
      <div class="card setup-card">
        <div class="setup-label">${ICONS.filter} 题目范围</div>
        ${chipGroup('practice', 'type', [['all', '全部题型'], ['single', '仅单选'], ['multiple', '仅多选'], ['judge', '仅判断']])}
      </div>
      <div class="card setup-card">
        <div class="setup-label">${ICONS.sort} 出题顺序</div>
        ${chipGroup('practice', 'order', [['seq', '顺序出题'], ['rand', '随机打乱']])}
      </div>
      <div class="card setup-card">
        <div class="setup-label">${ICONS.hash} 题量</div>
        ${chipGroup('practice', 'count', [['all', '全部'], [10, '10 题'], [20, '20 题'], [50, '50 题']])}
      </div>
      <button class="btn btn-accent setup-start" data-action="practice-start" ${list.length ? '' : 'disabled'}>开始作答（${list.length} 题）</button>
      ${list.length ? '' : '<div class="setup-tip">当前筛选条件下没有题目</div>'}
    </div>` : emptyState('还没有题目可以练习', '先去导入一张题目表格，或加载示例题快速体验',
      `<button class="btn btn-accent" data-action="goto" data-view="import">去导入</button>
       <button class="btn btn-ghost" data-action="load-sample">加载示例题目</button>`)}
  `;
}

function renderExamSetup() {
  const st = setupState.exam;
  const n = st.count === 'all' ? questions.length : Math.min(st.count, questions.length);
  view().innerHTML = `
    <div class="page-head"><div class="page-title">模拟考试</div><div class="page-sub">随机抽题、限时作答，交卷后统一评分（每题 1 分，60% 及格）</div></div>
    ${questions.length ? `<div class="setup">
      <div class="card setup-card">
        <div class="setup-label">${ICONS.hash} 抽题数量</div>
        ${chipGroup('exam', 'count', [[10, '10 题'], [20, '20 题'], [30, '30 题'], [50, '50 题'], ['all', `全部（${questions.length}）`]])}
      </div>
      <div class="card setup-card">
        <div class="setup-label">${ICONS.timer} 考试时长</div>
        ${chipGroup('exam', 'mins', [[15, '15 分钟'], [30, '30 分钟'], [45, '45 分钟'], [60, '60 分钟'], [0, '不限时']])}
      </div>
      <button class="btn btn-accent setup-start" data-action="exam-start">开始考试（${n} 题）</button>
      <div class="setup-tip">考试结束后，答错与未作答的题会自动收入错题本</div>
    </div>` : emptyState('还没有题目可以组卷', '先去导入一张题目表格，或加载示例题快速体验',
      `<button class="btn btn-accent" data-action="goto" data-view="import">去导入</button>
       <button class="btn btn-ghost" data-action="load-sample">加载示例题目</button>`)}
  `;
}

/* ---------------- 答题页 ---------------- */
function optionClass(q, r, key, isExam) {
  if (isExam || !r.submitted) return r.picked.has(key) ? 'sel' : '';
  const inAns = q.answer.includes(key);
  if (inAns && r.picked.has(key)) return 'correct';
  if (!inAns && r.picked.has(key)) return 'wrong';
  if (inAns) return 'miss';
  return '';
}

function renderQuiz() {
  const s = session;
  if (!s) { currentView = 'home'; render(); return; }
  const q = s.list[s.idx];
  const r = curResult();
  const isExam = s.mode === 'exam';
  const done = s.results.filter(x => x && x.picked.size).length;
  const modeName = { practice: '练习模式', exam: '模拟考试', wrong: '错题复习' }[s.mode];

  const optsHtml = q.options.map((o, i) => {
    const cls = optionClass(q, r, o.key, isExam);
    const disabled = !isExam && r.submitted;
    return `<button class="opt ${q.type === 'judge' ? 'judge' : ''} ${cls}" data-action="opt" data-i="${i}" ${disabled ? 'disabled' : ''}
      aria-pressed="${r.picked.has(o.key)}">
      <span class="opt-key">${optChip(q, o.key)}</span>
      <span class="opt-text">${esc(o.text)}</span>
    </button>`;
  }).join('');

  let fbHtml = '';
  if (!isExam && r.submitted) {
    const ok = r.correct;
    fbHtml = `<div class="fb ${ok ? 'fb-ok' : 'fb-bad'}">
      <div class="fb-title">${ok ? '✓ 回答正确' : '✗ 回答错误'}
        ${r.removed ? '<span class="fb-tag">已移出错题本</span>' : ''}
        ${!ok ? '<span class="fb-tag">已加入错题本</span>' : ''}</div>
      ${!ok ? `<div class="fb-line">你的答案：<b>${esc(pickedText(q, r.picked))}</b>　正确答案：<b class="ok">${esc(answerText(q))}</b></div>` : ''}
      ${q.analysis ? `<div class="fb-ana"><b>解析</b>${esc(q.analysis)}</div>` : ''}
    </div>`;
  }

  let footHtml;
  if (isExam) {
    footHtml = `
      <button class="btn btn-ghost" data-action="qprev" ${s.idx === 0 ? 'disabled' : ''}>← 上一题</button>
      <span class="spacer"></span>
      ${s.idx < s.list.length - 1
        ? `<button class="btn" data-action="qnext">下一题 →</button>`
        : `<button class="btn btn-accent" data-action="exam-submit">交卷</button>`}`;
  } else if (!r.submitted) {
    footHtml = q.type === 'multiple'
      ? `<span class="spacer"></span><button class="btn btn-accent" data-action="submit" ${r.picked.size ? '' : 'disabled'}>提交答案</button>`
      : `<span class="spacer"></span><span class="count-note">点击选项直接作答</span>`;
  } else {
    footHtml = `<span class="spacer"></span>${s.idx < s.list.length - 1
      ? `<button class="btn btn-accent" data-action="qnext">下一题 →</button>`
      : `<button class="btn btn-accent" data-action="practice-finish">查看结果</button>`}`;
  }

  const qnavHtml = isExam ? `
    <div class="qnav">${s.list.map((_, i) => {
      const rr = s.results[i];
      return `<button class="${rr && rr.picked.size ? 'answered' : ''} ${i === s.idx ? 'cur' : ''}" data-action="qgoto" data-i="${i}">${i + 1}</button>`;
    }).join('')}</div>
    <div class="qnav-legend"><span><i style="background:var(--teal)"></i>已作答</span><span><i style="background:var(--accent)"></i>当前题</span><span><i style="background:#fff;border:1px solid var(--line)"></i>未作答</span></div>` : '';

  view().innerHTML = `
    <div class="quiz-wrap anim-in">
      <div class="quiz-top">
        <button class="quiz-exit" data-action="exit-quiz">✕ 退出</button>
        <div class="quiz-prog">
          <div class="quiz-prog-text"><span>${modeName} · 第 ${s.idx + 1} / ${s.list.length} 题</span>${isExam ? `<span>已答 ${done} 题</span>` : ''}</div>
          <div class="pbar"><i style="width:${(s.idx + ((r.submitted || (isExam && r.picked.size)) ? 1 : 0)) / s.list.length * 100}%"></i></div>
        </div>
        ${isExam && s.deadline ? `<div class="quiz-timer" id="examTimer">剩余 ${fmtMS(s.deadline - Date.now())}</div>` : ''}
        ${isExam ? `<button class="btn btn-accent btn-sm" data-action="exam-submit">交卷</button>` : ''}
      </div>

      <div class="card qcard">
        <div class="q-meta">${badge(q.type)}<span class="src">快捷键：A~${String.fromCharCode(64 + Math.min(q.options.length, 8))} 选择，Enter 下一题</span></div>
        <div class="q-stem">${esc(q.question)}</div>
        ${q.type === 'multiple' ? '<div class="q-hint">本题为多选题，请选择全部正确选项后点击「提交答案」</div>' : ''}
        <div class="opts">${optsHtml}</div>
        ${fbHtml}
        ${qnavHtml}
      </div>
      <div class="quiz-foot">${footHtml}</div>
    </div>`;
}

/* ---------------- 结果页 ---------------- */
function renderResult() {
  const s = session;
  if (!s || !s.finished) { currentView = 'home'; render(); return; }
  const isExam = s.mode === 'exam';
  const total = s.list.length;
  const correct = s.results.filter(r => r.correct).length;
  const empty = isExam ? s.results.filter(r => !r.picked.size).length : 0;
  const wrong = total - correct - empty;
  const pct = total ? Math.round(correct / total * 100) : 0;
  const used = Date.now() - s.start;

  const reviewList = s.list.map((q, i) => ({ q, r: s.results[i] })).filter(x => !x.r.correct);
  const reviewHtml = reviewList.length ? `
    <div class="review-title">错题回顾（${reviewList.length}）</div>
    ${reviewList.map(({ q, r }) => `
      <div class="card rv-item">
        <div class="rv-head">${badge(q.type)}<div class="rv-q">${esc(q.question)}</div></div>
        <div class="rv-ans">你的答案：<b>${esc(pickedText(q, r.picked))}</b>　正确答案：<b class="ok">${esc(answerText(q))}</b></div>
        ${q.analysis ? `<div class="rv-ana"><b>解析：</b>${esc(q.analysis)}</div>` : ''}
      </div>`).join('')}` :
    `<div class="card empty" style="padding:36px"><div class="empty-title" style="font-size:19px">全部答对，一道错题都没有</div><div class="empty-desc">保持这个状态，去挑战更高题量吧</div></div>`;

  const C = 2 * Math.PI * 52;
  view().innerHTML = `
    <div class="result-wrap anim-in">
      <div class="card result-hero">
        <div class="ring-wrap">
          <svg class="ring" viewBox="0 0 120 120">
            <circle class="ring-bg" cx="60" cy="60" r="52"></circle>
            <circle class="ring-val" cx="60" cy="60" r="52" style="stroke-dasharray:0 ${C}"></circle>
          </svg>
          <div class="ring-txt"><div class="ring-score">${pct}%</div><div class="ring-sub">正确率</div></div>
        </div>
        <div class="result-title">${isExam ? (pct >= 60 ? '考试通过，继续保持' : '未达及格线，再接再厉') : '本次练习完成'}</div>
        <div class="result-sub">${s.mode === 'wrong' ? '答对的错题已自动移出错题本' : '答错的题目已自动收入错题本'}</div>
        <div class="result-stats">
          <div class="rs"><div class="rs-num g">${correct}</div><div class="rs-label">答对</div></div>
          <div class="rs"><div class="rs-num r">${wrong}</div><div class="rs-label">答错</div></div>
          ${isExam ? `<div class="rs"><div class="rs-num n">${empty}</div><div class="rs-label">未作答</div></div>` : ''}
          <div class="rs"><div class="rs-num">${fmtDur(used)}</div><div class="rs-label">用时</div></div>
        </div>
        <div class="result-actions">
          ${reviewList.length ? `<button class="btn btn-accent" data-action="goto" data-view="wrong">去错题本复习</button>` : ''}
          <button class="btn btn-ghost" data-action="restart">${isExam ? '再考一次' : '再练一次'}</button>
          <button class="btn btn-ghost" data-action="finish-home">返回概览</button>
        </div>
      </div>
      ${reviewHtml}
    </div>`;
  requestAnimationFrame(() => requestAnimationFrame(() => {
    const ring = document.querySelector('.ring-val');
    if (ring) ring.style.strokeDasharray = `${pct / 100 * C} ${C}`;
  }));
}

/* ---------------- 背题模式 ---------------- */
const reciteState = { type: 'all', hide: false, revealed: new Set() };
function renderRecite() {
  const list = questions.filter(q => reciteState.type === 'all' || q.type === reciteState.type);
  const chip = (v, label) => `<button class="chip ${reciteState.type === v ? 'on' : ''}" data-action="recite-type" data-val="${v}">${label}</button>`;

  view().innerHTML = `
    <div class="page-head"><div class="page-title">背题模式</div><div class="page-sub">所有答案直接展示，也可遮住答案自测，点击即可揭示</div></div>
    ${questions.length ? `
      <div class="list-head">
        ${chip('all', `全部 ${questions.length}`)}${chip('single', `单选 ${byType('single')}`)}${chip('multiple', `多选 ${byType('multiple')}`)}${chip('judge', `判断 ${byType('judge')}`)}
        <span style="flex:1"></span>
        <button class="btn ${reciteState.hide ? 'btn-accent' : 'btn-ghost'} btn-sm" data-action="recite-toggle">${reciteState.hide ? '显示全部答案' : '遮住答案自测'}</button>
      </div>
      <div class="count-note" style="margin-bottom:12px">共 ${list.length} 题${reciteState.hide ? ' · 已遮住答案，点击卡片底部揭示' : ''}</div>
      ${list.map((q, i) => {
        const hidden = reciteState.hide && !reciteState.revealed.has(q.id);
        return `<div class="card recite-card">
          <div class="q-meta">
            <span class="recite-num">${String(i + 1).padStart(2, '0')}</span>
            ${badge(q.type)}
            ${hidden ? '' : `<span class="recite-ans">答案：${esc(answerText(q))}</span>`}
          </div>
          <div class="q-stem" style="font-size:17px">${esc(q.question)}</div>
          <div class="opts">${q.options.map(o => `
            <div class="opt static ${q.type === 'judge' ? 'judge' : ''} ${!hidden && q.answer.includes(o.key) ? 'correct' : ''}">
              <span class="opt-key">${optChip(q, o.key)}</span>
              <span class="opt-text">${esc(o.text)}</span>
            </div>`).join('')}</div>
          ${hidden
            ? `<button class="ans-mask" data-action="recite-reveal" data-id="${q.id}">答案已遮住 · 点击显示</button>`
            : (q.analysis ? `<div class="rv-ana"><b>解析：</b>${esc(q.analysis)}</div>` : '')}
        </div>`;
      }).join('')}`
    : emptyState('题库还是空的', '导入题目后，即可在背题模式下通览全部题目与答案',
      `<button class="btn btn-accent" data-action="goto" data-view="import">去导入</button>
       <button class="btn btn-ghost" data-action="load-sample">加载示例题目</button>`)}
  `;
}

/* ---------------- 错题本 ---------------- */
function renderWrong() {
  const ids = Object.keys(wrongBook);
  const list = ids.map(id => questions.find(q => q.id === id)).filter(Boolean)
    .sort((a, b) => (wrongBook[b.id].t || 0) - (wrongBook[a.id].t || 0));

  view().innerHTML = `
    <div class="page-head"><div class="page-title">错题本</div><div class="page-sub">练习或考试中答错、未答的题会汇总在这里，答对一次即移出</div></div>
    ${list.length ? `
      <div class="list-head">
        <span class="count-note">共 ${list.length} 道错题</span>
        <span style="flex:1"></span>
        <button class="btn btn-accent" data-action="wrong-start">开始错题复习</button>
        <button class="btn btn-ghost btn-sm" data-action="wrong-clear">清空错题本</button>
      </div>
      ${list.map(q => `
        <div class="card mi">
          <span class="badge-flex">${badge(q.type)}</span>
          <div class="mi-body">
            <div class="mi-q">${esc(q.question)}</div>
            <div class="mi-meta">正确答案：<b>${esc(answerText(q))}</b></div>
          </div>
          <span class="mi-wrong">错 ${wrongBook[q.id].c} 次</span>
          <button class="mi-del" data-action="wrong-remove" data-id="${q.id}" title="移出错题本">${ICONS.trash}</button>
        </div>`).join('')}`
    : emptyState('错题本是空的', '刷题时答错的题会自动出现在这里，方便针对性复习',
      `<button class="btn btn-accent" data-action="goto" data-view="practice">去刷题</button>`)}
  `;
}

/* ---------------- 题库管理 ---------------- */
const manageState = { kw: '', type: 'all' };
function renderManage() {
  const kw = manageState.kw.trim().toLowerCase();
  const list = questions.filter(q =>
    (manageState.type === 'all' || q.type === manageState.type) &&
    (!kw || q.question.toLowerCase().includes(kw) || (q.analysis || '').toLowerCase().includes(kw)));

  view().innerHTML = `
    <div class="page-head"><div class="page-title">题库管理</div><div class="page-sub">查看、检索题库中的全部题目</div></div>
    ${questions.length ? `
      <div class="list-head">
        <div class="search">${ICONS.search}<input id="mngKw" placeholder="搜索题干或解析…" value="${esc(manageState.kw)}"></div>
        <button class="chip ${manageState.type === 'all' ? 'on' : ''}" data-action="mng-type" data-val="all">全部 ${questions.length}</button>
        <button class="chip ${manageState.type === 'single' ? 'on' : ''}" data-action="mng-type" data-val="single">单选 ${byType('single')}</button>
        <button class="chip ${manageState.type === 'multiple' ? 'on' : ''}" data-action="mng-type" data-val="multiple">多选 ${byType('multiple')}</button>
        <button class="chip ${manageState.type === 'judge' ? 'on' : ''}" data-action="mng-type" data-val="judge">判断 ${byType('judge')}</button>
        <span style="flex:1"></span>
        <button class="btn btn-ghost btn-sm" data-action="mng-clear" style="color:var(--red)">清空题库</button>
      </div>
      <div class="count-note" style="margin-bottom:12px">匹配 ${list.length} 题</div>
      ${list.length ? list.map(q => `
        <div class="card mi">
          <span class="badge-flex">${badge(q.type)}</span>
          <div class="mi-body">
            <div class="mi-q">${esc(q.question)}</div>
            <div class="mi-meta">答案：<b>${esc(answerText(q))}</b>${q.analysis ? '　·　含解析' : ''}${wrongBook[q.id] ? `　·　<span style="color:var(--red)">在错题本中</span>` : ''}</div>
          </div>
          <button class="mi-del" data-action="mng-del" data-id="${q.id}" title="删除此题">${ICONS.trash}</button>
        </div>`).join('')
      : `<div class="card empty" style="padding:40px"><div class="empty-title" style="font-size:18px">没有匹配的题目</div><div class="empty-desc">换个关键词或题型筛选试试</div></div>`}`
    : emptyState('题库还是空的', '导入表格后，所有题目都可以在这里检索和管理',
      `<button class="btn btn-accent" data-action="goto" data-view="import">去导入</button>
       <button class="btn btn-ghost" data-action="load-sample">加载示例题目</button>`)}
  `;
  const kwInput = document.getElementById('mngKw');
  if (kwInput) {
    let composing = false;
    const refocus = () => { const el = document.getElementById('mngKw'); el.focus(); el.setSelectionRange(el.value.length, el.value.length); };
    kwInput.addEventListener('compositionstart', () => composing = true);
    kwInput.addEventListener('compositionend', () => { composing = false; manageState.kw = kwInput.value; renderManage(); refocus(); });
    kwInput.addEventListener('input', () => {
      if (composing) return;
      manageState.kw = kwInput.value;
      renderManage(); refocus();
    });
  }
}

/* ---------------- 路由 ---------------- */
const RENDER = {
  home: renderHome, import: renderImport, practice: renderPracticeSetup,
  exam: renderExamSetup, quiz: renderQuiz, result: renderResult,
  wrong: renderWrong, manage: renderManage, recite: renderRecite,
};

function render() {
  (RENDER[currentView] || renderHome)();
  document.querySelectorAll('.nav-item').forEach(b => {
    b.classList.toggle('on', b.dataset.view === currentView ||
      (currentView === 'quiz' && ((session.mode === 'practice' && b.dataset.view === 'practice') ||
        (session.mode === 'exam' && b.dataset.view === 'exam') || (session.mode === 'wrong' && b.dataset.view === 'wrong'))) ||
      (currentView === 'result' && false));
  });
  const wb = document.getElementById('wrongBadge');
  const wc = wrongCount();
  wb.hidden = !wc; wb.textContent = wc;
  document.querySelector('.main').scrollTop = 0;
  window.scrollTo(0, 0);
}

async function go(v) {
  if (session && session.active && currentView === 'quiz' && v !== 'quiz') {
    const ok = await confirmDlg({ title: '退出当前答题？', body: '本次答题进度将不会保存。', okText: '退出', danger: true });
    if (!ok) return;
    endSession();
  }
  if (currentView !== 'result' && session && session.finished && v !== 'result') endSession();
  currentView = v;
  render();
}

/* ---------------- 全局事件 ---------------- */
document.getElementById('nav').addEventListener('click', e => {
  const b = e.target.closest('.nav-item');
  if (b) go(b.dataset.view);
});

view().addEventListener('click', async e => {
  const el = e.target.closest('[data-action]');
  if (!el || el.disabled) return;
  const a = el.dataset.action;
  const s = session;

  switch (a) {
    case 'goto': go(el.dataset.view); break;
    case 'load-sample': loadSample(); break;
    case 'tpl-download': downloadTemplate(); break;

    case 'chip': {
      const g = el.dataset.group, k = el.dataset.key;
      let v = el.dataset.val;
      if (v !== 'all' && !isNaN(Number(v)) && v !== '') v = Number(v);
      setupState[g][k] = v;
      render(); break;
    }
    case 'practice-start': {
      const st = setupState.practice;
      const list = buildList(st);
      if (!list.length) { toast('当前筛选条件下没有题目'); break; }
      startSession('practice', list); break;
    }
    case 'exam-start': {
      const st = setupState.exam;
      const list = buildList({ order: 'rand', count: st.count });
      if (!list.length) { toast('题库为空'); break; }
      startSession('exam', list, st.mins); break;
    }

    case 'opt': pickOption(Number(el.dataset.i)); break;
    case 'submit': submitPractice(); break;
    case 'qnext': if (s.idx < s.list.length - 1) { s.idx++; render(); } break;
    case 'qprev': if (s.idx > 0) { s.idx--; render(); } break;
    case 'qgoto': s.idx = Number(el.dataset.i); render(); break;
    case 'practice-finish': finishPractice(); break;
    case 'exam-submit': submitExamClick(); break;
    case 'exit-quiz': {
      const ok = await confirmDlg({ title: '退出当前答题？', body: '本次答题进度将不会保存。', okText: '退出', danger: true });
      if (ok) { endSession(); go('home'); }
      break;
    }

    case 'restart': {
      const mode = s.mode;
      if (mode === 'exam') { const st = setupState.exam; endSession(); startSession('exam', buildList({ order: 'rand', count: st.count }), st.mins); }
      else if (mode === 'wrong') { const ids = new Set(Object.keys(wrongBook)); const list = questions.filter(q => ids.has(q.id)); endSession(); list.length ? startSession('wrong', list) : go('wrong'); }
      else { const st = setupState.practice; endSession(); startSession('practice', buildList(st)); }
      break;
    }
    case 'finish-home': endSession(); go('home'); break;

    case 'wrong-start': {
      const ids = new Set(Object.keys(wrongBook));
      const list = questions.filter(q => ids.has(q.id));
      if (!list.length) { toast('错题本是空的'); break; }
      startSession('wrong', list); break;
    }
    case 'wrong-remove': {
      delete wrongBook[el.dataset.id]; save(LS.w, wrongBook);
      toast('已移出错题本'); render(); break;
    }
    case 'wrong-clear': {
      const ok = await confirmDlg({ title: '清空错题本？', body: `将移除全部 ${wrongCount()} 道错题记录。`, okText: '清空', danger: true });
      if (ok) { wrongBook = {}; save(LS.w, wrongBook); render(); }
      break;
    }

    case 'recite-type': reciteState.type = el.dataset.val; renderRecite(); break;
    case 'recite-toggle':
      reciteState.hide = !reciteState.hide;
      reciteState.revealed.clear();
      renderRecite(); break;
    case 'recite-reveal': reciteState.revealed.add(el.dataset.id); renderRecite(); break;

    case 'mng-type': manageState.type = el.dataset.val; renderManage(); break;
    case 'mng-del': {
      const id = el.dataset.id;
      const ok = await confirmDlg({ title: '删除这道题？', body: '删除后不可恢复。', okText: '删除', danger: true });
      if (ok) {
        questions = questions.filter(q => q.id !== id);
        delete wrongBook[id];
        save(LS.q, questions); save(LS.w, wrongBook);
        toast('已删除'); renderManage();
      }
      break;
    }
    case 'mng-clear': {
      const ok = await confirmDlg({ title: '清空整个题库？', body: `将删除全部 ${questions.length} 道题目及错题记录，且不可恢复。`, okText: '清空', danger: true });
      if (ok) {
        questions = []; wrongBook = {};
        save(LS.q, questions); save(LS.w, wrongBook);
        toast('题库已清空'); render();
      }
      break;
    }
  }
});

/* 键盘快捷键：答题页 A~H 选择、Enter 下一步 */
window.addEventListener('keydown', e => {
  if (currentView !== 'quiz' || !session) return;
  if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
  const k = e.key.toUpperCase();
  const q = session.list[session.idx];
  const i = OPT_KEYS.indexOf(k);
  if (i >= 0 && i < q.options.length) { pickOption(i); return; }
  if (e.key === 'Enter') {
    e.preventDefault();
    const r = curResult();
    if (session.mode === 'exam') { if (session.idx < session.list.length - 1) { session.idx++; render(); } }
    else if (!r.submitted) { if (q.type === 'multiple') submitPractice(); }
    else practiceNext();
  }
});

render();
