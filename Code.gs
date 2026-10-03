/**
 * 치주물루 유나이티드 FC 팬 사이트 — 후원·제안 신청서 접수 스크립트
 *
 * 하는 일
 *  1) 사이트의 "제안하기" 신청서를 이 구글 시트의 '신청' 탭에 저장합니다.
 *  2) 신청이 들어오면 관리자 이메일 모두에게 알림 메일을 보냅니다.
 *  3) 관리자 모드의 "후원·제안 신청" 탭에서 목록을 보고 상태를 바꿀 수 있게 해 줍니다.
 *
 * 설치 방법은 README.md의 "4. 후원·제안 신청서 이메일 받기"를 보세요.
 * 이 코드는 고칠 필요가 없습니다.
 */

const SHEET_NAME = '신청';
const HEADERS = ['번호', '접수 시각', '유형', '이름', '소속', '이메일', '제목', '내용', '언어', '상태', '메모'];
const MAX = { name: 100, org: 150, email: 200, subject: 200, body: 5000, type: 50 };

function doGet() {
  return out({ ok: true, service: 'cufc-forms', ready: !!props_().getProperty('SECRET') });
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return out({ ok: false, error: 'bad_request' }); }

  try {
    if (req.action === 'submit') return out(submit_(req));
    if (req.action === 'init') return out(init_(req));

    // 아래는 관리자만 (사이트의 관리자 계정 안에 잠겨 있는 비밀 키가 있어야 함)
    const secret = props_().getProperty('SECRET');
    if (!secret || req.secret !== secret) return out({ ok: false, error: 'unauthorized' });

    switch (req.action) {
      case 'list': return out(list_());
      case 'update': return out(update_(req));
      case 'delete': return out(remove_(req));
      case 'setAdmins': return out(setAdmins_(req));
      case 'test': return out(test_());
      default: return out({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  }
}

/* ---------- 처음 연결 ---------- */
function init_(req) {
  const p = props_();
  if (p.getProperty('SECRET')) return { ok: false, error: 'already_initialized' };
  if (!req.secret || String(req.secret).length < 32) return { ok: false, error: 'bad_secret' };
  p.setProperty('SECRET', String(req.secret));
  saveAdmins_(req.admins);
  sheet_();
  return { ok: true, sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl() };
}

/* 사이트를 처음부터 다시 연결해야 할 때만, 편집기에서 이 함수를 직접 실행하세요. */
function resetConnection() {
  props_().deleteProperty('SECRET');
  props_().deleteProperty('ADMINS');
  Logger.log('연결을 초기화했습니다. 사이트의 관리자 모드에서 다시 연결해 주세요.');
}

/* ---------- 신청 접수 ---------- */
function submit_(req) {
  if (req.website) return { ok: true };            // 스팸 로봇이 채우는 숨은 칸
  const f = {};
  for (const k in MAX) f[k] = String(req[k] || '').trim().slice(0, MAX[k]);
  if (!f.name || !f.email || !f.subject || !f.body) return { ok: false, error: 'missing_fields' };
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(f.email)) return { ok: false, error: 'bad_email' };

  // 같은 이메일로 1분 안에 여러 번 보내는 것 막기
  const cache = CacheService.getScriptCache();
  const key = 'rate_' + f.email.toLowerCase();
  if (cache.get(key)) return { ok: false, error: 'too_fast' };
  cache.put(key, '1', 60);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let id;
  try {
    const sh = sheet_();
    id = Utilities.getUuid().slice(0, 8);
    sh.appendRow([id, new Date(), f.type, f.name, f.org, f.email, f.subject, f.body, String(req.lang || 'ko'), '새 신청', '']);
  } finally {
    lock.releaseLock();
  }

  notify_(f, String(req.site || ''));
  return { ok: true, id };
}

function notify_(f, site) {
  const admins = loadAdmins_().map(a => a.email).filter(Boolean);
  if (!admins.length) return;
  const subject = '[치주물루 사이트] 새 ' + (f.type || '후원·제안') + ' 신청 — ' + f.subject;
  const lines = [
    '사이트에 새 신청이 들어왔어요.', '',
    '유형: ' + f.type,
    '이름: ' + f.name,
    f.org ? '소속: ' + f.org : null,
    '이메일: ' + f.email,
    '제목: ' + f.subject, '',
    f.body, '',
    '— 이 메일에 "답장"하면 신청한 분께 바로 회신돼요.',
    site ? '— 관리자 모드에서 보기: ' + site + '#/admin/inbox' : null,
    '— 구글 시트에서 보기: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  ].filter(x => x !== null);
  MailApp.sendEmail({ to: admins.join(','), replyTo: f.email, subject: subject, body: lines.join('\n'), name: '치주물루 사이트' });
}

/* ---------- 관리자 기능 ---------- */
function list_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  const rows = n > 0 ? sh.getRange(2, 1, n, HEADERS.length).getValues() : [];
  const items = rows.map(r => ({
    id: String(r[0]), at: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
    type: r[2], name: r[3], org: r[4], email: r[5], subject: r[6], body: r[7], lang: r[8], status: r[9], memo: r[10]
  })).reverse();
  return { ok: true, items, sheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl() };
}

function findRow_(id) {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(id));
  return i < 0 ? -1 : i + 2;
}

function update_(req) {
  const row = findRow_(req.id);
  if (row < 0) return { ok: false, error: 'not_found' };
  const sh = sheet_();
  if (req.status !== undefined) sh.getRange(row, 10).setValue(String(req.status).slice(0, 20));
  if (req.memo !== undefined) sh.getRange(row, 11).setValue(String(req.memo).slice(0, 1000));
  return { ok: true };
}

function remove_(req) {
  const row = findRow_(req.id);
  if (row < 0) return { ok: false, error: 'not_found' };
  sheet_().deleteRow(row);
  return { ok: true };
}

function setAdmins_(req) {
  saveAdmins_(req.admins);
  return { ok: true, count: loadAdmins_().length };
}

function test_() {
  const admins = loadAdmins_().map(a => a.email).filter(Boolean);
  if (!admins.length) return { ok: false, error: 'no_admin_emails' };
  MailApp.sendEmail({ to: admins.join(','), subject: '[치주물루 사이트] 테스트 메일', body: '신청서 알림이 이 주소로 잘 도착합니다.', name: '치주물루 사이트' });
  return { ok: true, sent: admins.length, remaining: MailApp.getRemainingDailyQuota() };
}

/* ---------- 도우미 ---------- */
function props_() { return PropertiesService.getScriptProperties(); }

function saveAdmins_(list) {
  const clean = (Array.isArray(list) ? list : [])
    .map(a => ({ id: String(a.id || '').slice(0, 40), email: String(a.email || '').trim().slice(0, 200) }))
    .filter(a => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(a.email));
  props_().setProperty('ADMINS', JSON.stringify(clean));
}
function loadAdmins_() {
  try { return JSON.parse(props_().getProperty('ADMINS') || '[]'); } catch (e) { return []; }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  }
  return sh;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
