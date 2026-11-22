/**
 * 수강생 발표 & 피드백 메모 보드 · made by 유환일
 *
 * 스프레드시트 하나가 여러 수업의 관리 대장입니다.
 *   '수업'   : 수업별 명단, 시간 설정, 발표 진행 상황
 *   '메모'   : 학생들이 남긴 익명 메모 (수업id로 구분)
 *   '좋아요' : 발표 좋아요 · 메모 공감 · 질문 공감 기록
 *   '질문'   : 특강 질문 받기에서 학생들이 남긴 익명 질문
 *
 * 설치: 새 구글 스프레드시트 → 확장 프로그램 → Apps Script
 *       → Code.gs 붙여넣기, 파일 추가(HTML) 이름 'Index' 로 Index.html 붙여넣기
 *       → 아래 ADMIN_KEY 변경 → 배포 → 새 배포 → 웹 앱 (실행: 나 / 액세스: 모든 사용자)
 * 교수자 콘솔: 배포 주소 뒤에 ?admin=관리키
 *
 * 주의: 시트 탭 이름('수업','메모','좋아요','질문')과 열 순서는 바꾸지 마세요. 파일 이름은 바꿔도 됩니다.
 */

const ADMIN_KEY = 'change-this-key';   // 교수자 관리용 키. 배포 전에 꼭 바꿔 주세요.

const MAX_LEN = 300;
const DEFAULTS = { p1: 120, p2: 60, fb: 60, l1: '발표', l2: '기획 설명', cats: ['좋았던 점', '제안', '차용하고 싶은 점'] };
const SH = { CLASSES: '수업', MEMOS: '메모', LIKES: '좋아요', QUESTIONS: '질문' };
const HEAD = {
  '수업': ['id', '수업명', '명단', '설정', '진행', '생성', '수정'],
  '메모': ['id', '시각', '수업id', '발표자', '분류', '내용', '숨김', '작성기기', '원문'],
  '좋아요': ['시각', '수업id', '종류', '대상', '기기'],
  '질문': ['id', '시각', '수업id', '내용', '숨김', '답변', '작성기기', '원문']
};

/* ---------- web app ---------- */
function doGet(e) {
  const prm = (e && e.parameter) || {};
  if (prm.ping) return jsonOut_({ ok: true, t: Date.now() });
  const isAdmin = !!prm.admin && prm.admin === ADMIN_KEY;
  const cls = prm.cls ? findClass_(String(prm.cls)) : null;   // 'c'는 Apps Script 예약 파라미터라 쓰지 않습니다
  let mode = 'home';
  if (isAdmin && !cls) mode = 'console';
  else if (isAdmin && cls && prm.m === 'stage') mode = 'stage';
  else if (cls) mode = 'board';
  const cfg = {
    mode: mode,
    base: ScriptApp.getService().getUrl(),
    admin: isAdmin ? ADMIN_KEY : '',
    c: cls ? cls.id : '',
    p: cls && cls.roster.indexOf(String(prm.p || '')) >= 0 ? String(prm.p) : '',
    view: String(prm.view || '')
  };
  const body = HtmlService.createHtmlOutputFromFile('Index').getContent();
  const boot = '<script>window.BOARD_CONFIG=' + JSON.stringify(cfg).replace(/</g, '\\u003c') + ';</script>';
  return HtmlService.createHtmlOutput(boot + body)
    .setTitle('수강생 발표 & 피드백 메모 보드')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ---------- sheet helpers ---------- */
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEAD[name]);
    sh.setFrozenRows(1);
  } else {
    // 예전에 만든 시트에 새 열(예: '원문')이 없으면 제목만 오른쪽 끝에 붙입니다. 기존 열은 건드리지 않습니다.
    const have = sh.getLastColumn(), want = HEAD[name].length;
    if (have < want) sh.getRange(1, have + 1, 1, want - have).setValues([HEAD[name].slice(have)]);
  }
  return sh;
}
function rows_(sh) {
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function clean_(s) {
  s = String(s || '').replace(/\r\n?/g, '\n').trim();
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // 스프레드시트 수식으로 해석되지 않게
  return s;
}
/** 한국어가 아닌 글이면 한국어로 자동 번역합니다(Google 번역). 한국어 글이거나 번역에 실패하면 null. */
function translateKo_(text) {
  const letters = String(text || '').replace(/[\s0-9.,!?'"“”‘’()\[\]{}:;~\-_/\\@#$%^&*+=<>|`…·]/g, '');
  if (!letters) return null;
  const ko = (letters.match(/[\uAC00-\uD7A3\u3131-\u318E]/g) || []).length;
  if (ko / letters.length >= 0.5) return null;
  try {
    const t = LanguageApp.translate(String(text), '', 'ko');
    if (t && t.trim() && t.trim() !== String(text).trim()) return t.trim();
  } catch (e) {}
  return null;
}
function json_(s, fallback) { try { return s ? JSON.parse(s) : fallback; } catch (e) { return fallback; } }
function checkAdmin_(key) { if (!key || key !== ADMIN_KEY) throw new Error('관리 권한이 없습니다. 관리자 주소로 다시 열어 주세요.'); }
function bust_(c) { CacheService.getScriptCache().remove('b:' + c); }

/** 명단 텍스트 → 이름 배열. 학번, (괄호), 번호를 걷어내고 중복 이름에는 (2)를 붙입니다. */
function rosterFromText_(text) {
  const out = [], seen = {};
  String(text || '').split(/\r?\n/).forEach(function (raw) {
    let line = raw.trim();
    if (!line || line.charAt(0) === '#') return;
    line = line.replace(/^\d{1,3}[.)]?\s+/, '').replace(/[(\[（][^)\]）]*[)\]）]/g, ' ').trim();
    const fields = line.split(/[\t,;|]/).map(function (s) { return s.replace(/\d{5,}/g, '').trim(); }).filter(String);
    if (!fields.length) return;
    let name = fields[0];
    const toks = name.split(/\s+/);
    if (toks.length > 1 && /[가-힣]/.test(toks[0])) name = toks[0];
    if (/^(학번|이름|성명|순번|번호|name)$/i.test(name)) return;
    name = name.slice(0, 30);
    if (seen[name]) { seen[name]++; name = name + ' (' + seen[name] + ')'; } else seen[name] = 1;
    out.push(name);
  });
  return out;
}
function settings_(s) {
  s = s || {};
  const num = function (v, d, lo, hi) { v = Math.round(Number(v)); return isFinite(v) ? Math.min(hi, Math.max(lo, v)) : d; };
  let cats = Array.isArray(s.cats) ? s.cats : String(s.cats || '').split(',');
  cats = cats.map(function (x) { return String(x).trim().slice(0, 20); }).filter(String).slice(0, 5);
  if (!cats.length) cats = DEFAULTS.cats.slice();
  return {
    p1: num(s.p1, DEFAULTS.p1, 10, 1800),
    p2: num(s.p2, DEFAULTS.p2, 0, 1800),
    fb: num(s.fb, DEFAULTS.fb, 0, 900),
    l1: String(s.l1 || DEFAULTS.l1).slice(0, 20),
    l2: String(s.l2 || DEFAULTS.l2).slice(0, 20),
    cats: cats,
    type: s.type === 'qa' ? 'qa' : 'talk',        // talk = 발표·피드백, qa = 특강 질문 받기
    title: String(s.title || '').trim().slice(0, 60) // 특강 질문 화면 제목(학생에게 보임)
  };
}
function freshProgress_(roster) { return { pool: roster.slice(), order: [], absent: [] }; }
/** 명단이 바뀌었을 때: 이미 발표·결석 처리된 사람은 두고, 나머지만 뽑기 대상으로 */
function reconcile_(roster, pg) {
  pg = pg || freshProgress_(roster);
  const inRoster = {}; roster.forEach(function (n) { inRoster[n] = 1; });
  const order = (pg.order || []).filter(function (o) { return inRoster[o.name]; });
  const absent = (pg.absent || []).filter(function (n) { return inRoster[n]; });
  const used = {}; order.forEach(function (o) { used[o.name] = 1; }); absent.forEach(function (n) { used[n] = 1; });
  return { pool: roster.filter(function (n) { return !used[n]; }), order: order, absent: absent };
}

function findClass_(id) {
  const rows = rows_(sheet_(SH.CLASSES));
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === id) {
      const roster = rosterFromText_(rows[i][2]);
      return {
        row: i + 2, id: id, name: String(rows[i][1]), rosterText: String(rows[i][2]), roster: roster,
        settings: settings_(json_(rows[i][3], {})),
        progress: reconcile_(roster, json_(rows[i][4], null)),
        created: rows[i][5], updated: rows[i][6]
      };
    }
  }
  return null;
}
function newId_() {
  const chars = 'abcdefghijkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 7; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

/** 처음 한 번 실행하면 관리 대장 이름을 정하고 시트 세 개를 미리 만들어 둡니다(선택). */
function setup() {
  SpreadsheetApp.getActiveSpreadsheet().rename('수강생 발표 & 피드백 메모 보드');
  sheet_(SH.CLASSES); sheet_(SH.MEMOS); sheet_(SH.LIKES); sheet_(SH.QUESTIONS);
}

/* ---------- public API (students) ---------- */
function getBoard(c) {
  c = String(c || '');
  const cache = CacheService.getScriptCache();
  const hit = cache.get('b:' + c);
  if (hit) return hit;
  const cls = findClass_(c);
  if (!cls) throw new Error('보드를 찾지 못했습니다. QR이나 주소를 다시 확인해 주세요.');
  if (cls.settings.type === 'qa') {
    const out = JSON.stringify(qaPayload_(cls));
    try { cache.put('b:' + c, out, 4); } catch (err) {}
    return out;
  }

  const memos = rows_(sheet_(SH.MEMOS))
    .filter(function (r) { return r[0] && String(r[2]) === c && !isTrue_(r[6]); })
    .map(function (r) {
      const m = { id: String(r[0]), ts: new Date(r[1]).getTime(), pid: String(r[3]), cat: String(r[4]), text: String(r[5]).replace(/^'/, '') };
      if (r[8]) m.orig = String(r[8]).replace(/^'/, '');
      return m;
    });
  const memoLikes = {}, talkLikes = {};
  rows_(sheet_(SH.LIKES)).forEach(function (r) {
    if (String(r[1]) !== c) return;
    const kind = String(r[2]), t = String(r[3]);
    if (kind === 'memo') memoLikes[t] = (memoLikes[t] || 0) + 1;
    else if (kind === 'talk') talkLikes[t] = (talkLikes[t] || 0) + 1;
  });
  memos.forEach(function (m) { m.likes = memoLikes[m.id] || 0; });

  const out = JSON.stringify({ roster: cls.roster, cats: cls.settings.cats, memos: memos, talkLikes: talkLikes, updated: Date.now() });
  try { cache.put('b:' + c, out, 5); } catch (err) {}
  return out;
}

function addMemo(c, pid, cat, text, token) {
  c = String(c || ''); pid = String(pid || ''); cat = String(cat || '');
  let raw = String(text || '').trim().slice(0, MAX_LEN);
  const tr = translateKo_(raw);
  text = clean_(tr || raw); token = String(token || '').slice(0, 64);
  const orig = tr ? clean_(raw) : '';
  const cls = findClass_(c);
  if (!cls) throw new Error('보드를 찾지 못했습니다.');
  if (cls.roster.indexOf(pid) < 0) throw new Error('발표자를 다시 선택해 주세요.');
  if (cls.settings.cats.indexOf(cat) < 0) throw new Error('메모 분류를 골라 주세요.');
  if (!text || text === "'") throw new Error('메모 내용을 입력해 주세요.');
  if (text.length > MAX_LEN) text = text.slice(0, MAX_LEN);
  const cache = CacheService.getScriptCache();
  if (token && cache.get('rate:' + token)) throw new Error('메모는 5초에 하나씩 남길 수 있습니다. 잠시 후 다시 눌러 주세요.');
  withLock_(function () { sheet_(SH.MEMOS).appendRow([Utilities.getUuid(), new Date(), c, pid, cat, text, false, token, orig]); });
  if (token) cache.put('rate:' + token, '1', 5);
  bust_(c);
  return getBoard(c);
}

function toggleLike(c, kind, target, token) {
  c = String(c || ''); kind = String(kind || ''); target = String(target || ''); token = String(token || '').slice(0, 64);
  if ((kind !== 'memo' && kind !== 'talk' && kind !== 'q') || !target || !token) throw new Error('좋아요를 처리하지 못했습니다. 새로고침 후 다시 눌러 주세요.');
  const cls = findClass_(c);
  if (!cls) throw new Error('보드를 찾지 못했습니다.');
  if (kind === 'talk' && cls.roster.indexOf(target) < 0) throw new Error('발표자를 찾지 못했습니다.');
  withLock_(function () {
    const sh = sheet_(SH.LIKES), rows = rows_(sh);
    for (let i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][1]) === c && String(rows[i][2]) === kind && String(rows[i][3]) === target && String(rows[i][4]) === token) { sh.deleteRow(i + 2); return; }
    }
    sh.appendRow([new Date(), c, kind, target, token]);
  });
  bust_(c);
  return getBoard(c);
}

/* ---------- admin API (교수자) ---------- */
function hideMemo(c, id, key) {
  checkAdmin_(key);
  withLock_(function () {
    const sh = sheet_(SH.MEMOS), rows = rows_(sh);
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id)) { sh.getRange(i + 2, 7).setValue(true); return; }
    }
  });
  bust_(String(c));
  return getBoard(String(c));
}

function listClasses(key) {
  checkAdmin_(key);
  const memoCount = {}, qCount = {};
  rows_(sheet_(SH.MEMOS)).forEach(function (r) { if (r[0] && !isTrue_(r[6])) memoCount[r[2]] = (memoCount[r[2]] || 0) + 1; });
  rows_(sheet_(SH.QUESTIONS)).forEach(function (r) { if (r[0] && !isTrue_(r[4])) qCount[r[2]] = (qCount[r[2]] || 0) + 1; });
  const list = rows_(sheet_(SH.CLASSES)).filter(function (r) { return r[0]; }).map(function (r) {
    const roster = rosterFromText_(r[2]);
    const pg = reconcile_(roster, json_(r[4], null));
    const st = settings_(json_(r[3], {}));
    return {
      id: String(r[0]), name: String(r[1]), count: roster.length, type: st.type, title: st.title, questions: qCount[r[0]] || 0,
      done: pg.order.filter(function (o) { return o.status === 'done'; }).length,
      absent: pg.absent.length, memos: memoCount[r[0]] || 0,
      created: r[5] ? new Date(r[5]).getTime() : 0, updated: r[6] ? new Date(r[6]).getTime() : 0
    };
  });
  list.sort(function (a, b) { return b.updated - a.updated; });
  return JSON.stringify(list);
}

function getClass(key, id) {
  checkAdmin_(key);
  const cls = findClass_(String(id));
  if (!cls) throw new Error('수업을 찾지 못했습니다.');
  return JSON.stringify({ id: cls.id, name: cls.name, rosterText: cls.rosterText, roster: cls.roster, settings: cls.settings, progress: cls.progress });
}

function saveClass(key, data) {
  checkAdmin_(key);
  data = data || {};
  const name = String(data.name || '').trim().slice(0, 60);
  const st = settings_(data.settings);
  const rosterText = st.type === 'qa' ? '' : String(data.rosterText || '').slice(0, 20000);
  const roster = rosterFromText_(rosterText);
  if (!name) throw new Error('수업 이름을 입력해 주세요.');
  if (st.type !== 'qa' && !roster.length) throw new Error('명단에서 이름을 찾지 못했습니다. 한 줄에 한 명씩 입력해 주세요.');
  const now = new Date();
  let id = String(data.id || '');
  withLock_(function () {
    const sh = sheet_(SH.CLASSES);
    const cur = id ? findClass_(id) : null;
    if (cur) {
      const pg = reconcile_(roster, cur.progress);
      sh.getRange(cur.row, 2, 1, 6).setValues([[name, rosterText, JSON.stringify(st), JSON.stringify(pg), cur.created || now, now]]);
    } else {
      id = newId_();
      sh.appendRow([id, name, rosterText, JSON.stringify(st), JSON.stringify(freshProgress_(roster)), now, now]);
    }
  });
  bust_(id);
  return getClass(key, id);
}

function saveProgress(key, id, progress) {
  checkAdmin_(key);
  id = String(id || '');
  const out = withLock_(function () {
    const cls = findClass_(id);
    if (!cls) throw new Error('수업을 찾지 못했습니다.');
    const pg = progress || {};
    const order = (Array.isArray(pg.order) ? pg.order : []).map(function (o) {
      return { name: String(o.name || ''), status: o.status === 'done' ? 'done' : 'drawn', used: o.used == null ? null : Number(o.used) || 0, over: Number(o.over) || 0 };
    });
    const clean = reconcile_(cls.roster, { order: order, absent: Array.isArray(pg.absent) ? pg.absent.map(String) : [] });
    sheet_(SH.CLASSES).getRange(cls.row, 5, 1, 3).setValues([[JSON.stringify(clean), cls.created || new Date(), new Date()]]);
    return clean;
  });
  return JSON.stringify(out);
}

function resetProgress(key, id) {
  checkAdmin_(key);
  return saveProgress(key, id, { order: [], absent: [] });
}

/** 시트에서 col번째 열(1부터)이 id인 줄을 모두 지웁니다. 남은 줄을 한 번에 다시 써서 빠르게 처리합니다. */
function removeRows_(sh, col, id) {
  const rows = rows_(sh);
  if (!rows.length) return 0;
  const keep = rows.filter(function (r) { return String(r[col - 1]) !== id; });
  const removed = rows.length - keep.length;
  if (!removed) return 0;
  const width = sh.getLastColumn();
  sh.getRange(2, 1, rows.length, width).clearContent();
  if (keep.length) sh.getRange(2, 1, keep.length, width).setValues(keep);
  return removed;
}

/** 수업 삭제: '수업' 줄과 그 수업의 메모·좋아요를 모두 지웁니다. 되돌릴 수 없습니다. */
function deleteClass(key, id) {
  checkAdmin_(key);
  id = String(id || '');
  if (!id) throw new Error('수업을 찾지 못했습니다.');
  const out = withLock_(function () {
    if (!findClass_(id)) throw new Error('이미 삭제되었거나 없는 수업입니다.');
    return {
      classes: removeRows_(sheet_(SH.CLASSES), 1, id),
      memos: removeRows_(sheet_(SH.MEMOS), 3, id),
      likes: removeRows_(sheet_(SH.LIKES), 2, id),
      questions: removeRows_(sheet_(SH.QUESTIONS), 3, id)
    };
  });
  bust_(id);
  return JSON.stringify(out);
}

/* ---------- 특강 질문 받기 ---------- */
function qaPayload_(cls) {
  const c = cls.id;
  const qs = rows_(sheet_(SH.QUESTIONS))
    .filter(function (r) { return r[0] && String(r[2]) === c && !isTrue_(r[4]); })
    .map(function (r) {
      const q = { id: String(r[0]), ts: new Date(r[1]).getTime(), text: String(r[3]).replace(/^'/, ''), answered: isTrue_(r[5]) };
      if (r[7]) q.orig = String(r[7]).replace(/^'/, '');
      return q;
    });
  const likes = {};
  rows_(sheet_(SH.LIKES)).forEach(function (r) {
    if (String(r[1]) === c && String(r[2]) === 'q') likes[String(r[3])] = (likes[String(r[3])] || 0) + 1;
  });
  qs.forEach(function (q) { q.likes = likes[q.id] || 0; });
  return { type: 'qa', title: cls.settings.title || '특강 질문', questions: qs, updated: Date.now() };
}

function addQuestion(c, text, token) {
  c = String(c || ''); token = String(token || '').slice(0, 64);
  const raw = String(text || '').trim().slice(0, MAX_LEN);
  const tr = translateKo_(raw);
  text = clean_(tr || raw);
  const orig = tr ? clean_(raw) : '';
  const cls = findClass_(c);
  if (!cls || cls.settings.type !== 'qa') throw new Error('질문 보드를 찾지 못했습니다. QR이나 주소를 다시 확인해 주세요.');
  if (!text || text === "'") throw new Error('질문 내용을 입력해 주세요.');
  if (text.length > MAX_LEN) text = text.slice(0, MAX_LEN);
  const cache = CacheService.getScriptCache();
  if (token && cache.get('rate:' + token)) throw new Error('질문은 5초에 하나씩 남길 수 있습니다. 잠시 후 다시 눌러 주세요.');
  withLock_(function () { sheet_(SH.QUESTIONS).appendRow([Utilities.getUuid(), new Date(), c, text, false, false, token, orig]); });
  if (token) cache.put('rate:' + token, '1', 5);
  bust_(c);
  return getBoard(c);
}

function setQuestion_(c, id, key, col, val) {
  checkAdmin_(key);
  withLock_(function () {
    const sh = sheet_(SH.QUESTIONS), rows = rows_(sh);
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(id) && String(rows[i][2]) === String(c)) { sh.getRange(i + 2, col).setValue(val); return; }
    }
  });
  bust_(String(c));
  return getBoard(String(c));
}
/** 교수자: 질문을 답변 완료(true) 또는 답변 전(false)으로 표시 */
function markAnswered(c, id, val, key) { return setQuestion_(c, id, key, 6, val === true || val === 'true'); }
/** 교수자: 질문 숨기기 */
function hideQuestion(c, id, key) { return setQuestion_(c, id, key, 5, true); }


/* ---------- JSON API: GitHub Pages 화면이 이 주소로 데이터를 주고받습니다 ---------- */
// 화면은 GitHub Pages에서 열리므로, 학생 휴대폰에 구글 계정이 여러 개 로그인되어 있어도 오류가 나지 않습니다.
const API_ = {
  getBoard: getBoard, addMemo: addMemo, toggleLike: toggleLike, hideMemo: hideMemo,
  listClasses: listClasses, getClass: getClass, saveClass: saveClass,
  saveProgress: saveProgress, resetProgress: resetProgress, deleteClass: deleteClass,
  addQuestion: addQuestion, markAnswered: markAnswered, hideQuestion: hideQuestion
};
function jsonOut_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = Object.prototype.hasOwnProperty.call(API_, req.fn) ? API_[req.fn] : null;
    if (!fn) throw new Error('알 수 없는 요청입니다.');
    const r = fn.apply(null, Array.isArray(req.args) ? req.args : []);
    return jsonOut_({ ok: true, data: typeof r === 'string' ? JSON.parse(r) : r });
  } catch (err) {
    return jsonOut_({ ok: false, error: String((err && err.message) || err).replace(/^Error:\s*/, '') });
  }
}
