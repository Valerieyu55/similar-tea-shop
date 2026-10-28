/**
 * 比例手搖飲經營遊戲（相似形）· 後台紀錄服務
 * 指導老師：余佳霖老師
 *
 * 安裝步驟（只要做一次）：
 * 1. 新增一份 Google 試算表 → 擴充功能 → Apps Script，把這整份程式貼上、存檔。
 * 2. 上方函式選單選「setup」→ 執行。第一次會要求授權，按允許。
 *    （會建立「學生總表」「作答紀錄」兩個工作表，並把教師密碼設為 55555）
 * 3. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *      執行身分：我　／　誰可以存取：任何人
 *    部署後複製「網頁應用程式網址」（…/exec 結尾）。
 * 4. 把網址貼到遊戲 HTML 開頭的 const LOG_URL='' 引號裡，再上傳到 GitHub Pages。
 *
 * 想改教師密碼：專案設定 → 指令碼屬性 → TEACHER_KEY。
 * 程式更新後要「部署 → 管理部署作業 → 編輯 → 版本選新版本」，網址才會沿用。
 */

const SUM = '學生總表';
const LOG = '作答紀錄';
const ROS = '學生名單';
const ROS_HEAD = ['班級','座號','姓名'];
const ZONE_COLS = [['lab','研發部'],['draft','杯身設計室'],['brand','品牌部'],['bar','吧台'],['office','財務室']];
const SUM_HEAD = ['班級','座號','姓名','店名','最後上線','目前進度','等級','經驗值','營收'].concat(ZONE_COLS.map(z => z[1]), ['總星數','大杯定價','隱藏卡','累計作答','第一次答對率','一次答對題數','遊戲時間(分)','分數','名單核對']);
const LOG_HEAD = ['收到時間','事件ID','班級','座號','姓名','店名','事件','題號','題目','學生作答（依序）','正確答案','第一次就答對','最後答對','作答次數','看提示','作答秒數','迷思診斷','備註'];
const EV = {login:'登入', shop:'改店名', start:'進入關卡', answer:'作答', clear:'過關', leave:'中途離開', reset:'重設進度'};
const MAX_LOG_RETURN = 12000;

function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TEACHER_KEY')) props.setProperty('TEACHER_KEY', '55555');
  sheet_(SUM, SUM_HEAD);
  sheet_(LOG, LOG_HEAD);
  sheet_(ROS, ROS_HEAD);
}

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#fff3d6');
    sh.setFrozenRows(1);
    // 班級、座號存成純文字，避免 05 被轉成 5
    if (name === SUM || name === ROS) sh.getRange('A:B').setNumberFormat('@');
    if (name === LOG) { sh.getRange('C:D').setNumberFormat('@'); sh.getRange('H:H').setNumberFormat('@'); }
  }
  else {
    // 欄位更新時，自動把標題列換成新版
    const cur = sh.getRange(1, 1, 1, head.length).getDisplayValues()[0];
    if (cur.join('|') !== head.join('|')) sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#fff3d6');
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return json_({ok: true, service: 'similar-tea-shop-log'});
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ok: false, error: 'badjson'}); }
  if (d.action === 'teacher') return teacher_(String(d.key || ''));
  if (d.action === 'board') return board_(d);
  if (d.action === 'verify') return verify_(d);
  if (d.action === 'roster') return setRoster_(d);
  return ingest_(d);
}

/* 刪除班級填「測試」的資料（老師自己試玩後可執行一次） */
function removeTestRows() {
  [SUM, LOG].forEach(name => {
    const sh = sheet_(name, name === SUM ? SUM_HEAD : LOG_HEAD);
    const col = name === SUM ? 1 : 3;
    for (let r = sh.getLastRow(); r >= 2; r--) {
      if (String(sh.getRange(r, col).getValue()).trim() === '測試') sh.deleteRow(r);
    }
  });
  CacheService.getScriptCache().remove('board');
}

/* ---------- 學生名單 ---------- */
const normName_ = n => String(n || '').replace(/\s+/g, '');
const key_ = (c, s) => String(c || '').trim() + '-' + String(+s);
function roster_() {
  const sh = sheet_(ROS, ROS_HEAD);
  const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues() : [];
  const map = {}, classes = {};
  v.forEach(r => { if (String(r[0]).trim() === '') return; map[key_(r[0], r[1])] = normName_(r[2]); classes[String(r[0]).trim()] = 1; });
  return {map, classes, rows: v.filter(r => String(r[0]).trim() !== ''), size: Object.keys(map).length};
}
function check_(R, cls, seat, name) {
  if (!R.size) return '';
  if (!R.classes[String(cls || '').trim()]) return 'cls';
  const k = key_(cls, seat);
  if (!(k in R.map)) return 'seat';
  if (R.map[k] !== normName_(name)) return 'name';
  return 'ok';
}
// 學生登入時核對（只回傳對不對，不會透露正確姓名）
function verify_(d) {
  const R = roster_();
  if (!R.size) return json_({ok: true, roster: false});
  const r = check_(R, d.cls, d.seat, d.name);
  return json_({ok: r === 'ok', roster: true, why: r === 'ok' ? '' : r});
}
// 老師上傳名單（需要教師密碼，會整份取代）
function setRoster_(d) {
  const cache = CacheService.getScriptCache(), fails = +(cache.get('tfail') || 0);
  if (fails >= 8) return json_({ok: false, error: 'locked'});
  const real = PropertiesService.getScriptProperties().getProperty('TEACHER_KEY');
  if (!real || String(d.key || '') !== real) { cache.put('tfail', String(fails + 1), 600); return json_({ok: false, error: 'bad'}); }
  const rows = (Array.isArray(d.rows) ? d.rows : []).map(r => [String(r[0] || '').trim(), String(+r[1]), normName_(r[2])])
    .filter(r => r[0] && r[1] !== 'NaN' && r[2]);
  if (!rows.length) return json_({ok: false, error: 'empty'});
  const sh = sheet_(ROS, ROS_HEAD);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  sh.getRange('A:B').setNumberFormat('@');
  sh.getRange(2, 1, rows.length, 3).setValues(rows);
  // 重新標記學生總表的「名單核對」
  const R = roster_(), sum = sheet_(SUM, SUM_HEAD), ci = SUM_HEAD.indexOf('名單核對') + 1;
  if (sum.getLastRow() > 1) {
    const v = sum.getRange(2, 1, sum.getLastRow() - 1, 3).getValues();
    sum.getRange(2, ci, v.length, 1).setValues(v.map(r => [check_(R, r[0], r[1], r[2]) === 'ok' ? '✔' : '名單外']));
  }
  return json_({ok: true, n: rows.length});
}

/* ---------- 學生上傳 ---------- */
function ingest_(d) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return json_({ok: false, error: 'busy'}); }
  try {
    const cache = CacheService.getScriptCache();
    const events = Array.isArray(d.events) ? d.events : [];
    const rows = [];
    const now = new Date();
    events.forEach(ev => {
      if (!ev || !ev.id || !ev.st) return;
      const ck = 'e_' + ev.id;
      if (cache.get(ck)) return;          // 重複上傳的事件略過
      cache.put(ck, '1', 21600);
      const yn = v => (v === true ? '是' : v === false ? '否' : '');
      const isAns = ev.type === 'answer';
      rows.push([
        ev.t ? new Date(ev.t) : now, ev.id,
        String(ev.st.cls || ''), String(ev.st.seat || ''), String(ev.st.name || ''), String(ev.shop || ''),
        EV[ev.type] || ev.type, ev.qno ? "'" + ev.qno : '', ev.question || '', ev.given || '', ev.correct || '',
        isAns ? yn(ev.first) : '', isAns ? yn(ev.ok) : '', isAns ? (ev.tries || '') : '',
        isAns ? yn(ev.hint) : '', isAns ? (ev.sec || '') : '', ev.diag || '', ev.note || ''
      ]);
    });
    if (rows.length) {
      const lg = sheet_(LOG, LOG_HEAD);
      lg.getRange(lg.getLastRow() + 1, 1, rows.length, LOG_HEAD.length).setValues(rows);
    }
    if (d.summary && d.summary.st) upsert_(d.summary, now);
    return json_({ok: true, n: rows.length});
  } finally {
    lock.releaseLock();
  }
}

function upsert_(s, now) {
  const sh = sheet_(SUM, SUM_HEAD);
  const cls = String(s.st.cls || ''), seat = String(s.st.seat || '');
  const zs = s.zs || {};
  const stat = s.stat || {n: 0, f: 0};
  const row = [
    cls, seat, String(s.st.name || ''), String(s.shop || ''), now, s.cur || '',
    s.lv || 1, s.xp || 0, s.rev || 0
  ].concat(ZONE_COLS.map(z => zs[z[0]] ? zs[z[0]] + '★' : ''), [
    s.stars || 0, s.price || '', s.bonus ? '有' : '',
    stat.n || 0, stat.n ? Math.round(stat.f / stat.n * 100) + '%' : '',
    s.first || 0, Math.round((s.ms || 0) / 6000) / 10, s.score || 0
  ]);
  const ck = check_(roster_(), cls, seat, s.st.name);
  row.push(ck === '' ? '' : ck === 'ok' ? '✔' : '名單外');
  const last = sh.getLastRow();
  if (last > 1) {
    const keys = sh.getRange(2, 1, last - 1, 2).getDisplayValues();
    for (let i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === cls && String(+keys[i][1]) === String(+seat)) {
        sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
        return;
      }
    }
  }
  sh.getRange(last + 1, 1, 1, row.length).setValues([row]);
}

/* ---------- 學生排行榜（公開，不需密碼；姓名只顯示頭尾字，不回傳座號） ---------- */
function mask_(n) {
  n = String(n || '').trim();
  if (n.length <= 1) return n;
  if (n.length === 2) return n[0] + '○';
  return n[0] + '○'.repeat(n.length - 2) + n[n.length - 1];
}
function board_(d) {
  const cache = CacheService.getScriptCache();
  let rows = null;
  const hit = cache.get('board');
  if (hit) rows = JSON.parse(hit);
  else {
    const sh = sheet_(SUM, SUM_HEAD);
    const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, SUM_HEAD.length).getValues() : [];
    const ix = n => SUM_HEAD.indexOf(n);
    rows = v.filter(r => String(r[0]).trim() !== '').map(r => ({
      k: String(r[0]).trim() + '-' + String(+r[1]),
      cls: String(r[0]).trim(), name: mask_(r[2]), shop: String(r[ix('店名')] || ''),
      stars: +r[ix('總星數')] || 0, first: +r[ix('一次答對題數')] || 0,
      score: +r[ix('分數')] || 0, min: +r[ix('遊戲時間(分)')] || 0
    })).filter(r => r.score > 0).sort((a, b) => b.score - a.score || a.min - b.min);
    const txt = JSON.stringify(rows);
    if (txt.length < 95000) cache.put('board', txt, 20);
  }
  const me = String(d.cls || '').trim() + '-' + String(+d.seat);
  const meIdx = rows.findIndex(r => r.k === me);
  const out = rows.slice(0, 300).map((r, i) => ({cls: r.cls, name: r.name, shop: r.shop, stars: r.stars, first: r.first, score: r.score, min: r.min, me: i === meIdx}));
  return json_({ok: true, rows: out, me: meIdx >= 0 ? {rank: meIdx + 1} : null});
}

/* ---------- 教師讀取 ---------- */
function teacher_(key) {
  const cache = CacheService.getScriptCache();
  const fails = +(cache.get('tfail') || 0);
  if (fails >= 8) return json_({ok: false, error: 'locked'});
  const real = PropertiesService.getScriptProperties().getProperty('TEACHER_KEY');
  if (!real) return json_({ok: false, error: 'nokey'});
  if (key !== real) { cache.put('tfail', String(fails + 1), 600); return json_({ok: false, error: 'bad'}); }
  cache.remove('tfail');
  const sum = sheet_(SUM, SUM_HEAD), lg = sheet_(LOG, LOG_HEAD);
  const sv = sum.getLastRow() > 1 ? sum.getRange(2, 1, sum.getLastRow() - 1, SUM_HEAD.length).getValues() : [];
  const total = lg.getLastRow() - 1;
  const n = Math.min(Math.max(total, 0), MAX_LOG_RETURN);
  const lv = n > 0 ? lg.getRange(lg.getLastRow() - n + 1, 1, n, LOG_HEAD.length).getValues() : [];
  return json_({ok: true, at: Date.now(), sumHead: SUM_HEAD, sum: sv, logHead: LOG_HEAD, log: lv, truncated: total > MAX_LOG_RETURN, roster: roster_().rows});
}
