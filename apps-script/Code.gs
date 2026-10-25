/**
 * 센서 탐험대 스튜디오 — 아이디어 전시관 저장소 (Google Apps Script)
 *
 * 설치 (3분)
 * 1. 새 구글 시트를 만든다 (이름 예: 센서 탐험대 아이디어 전시관 6-3) → 메뉴 [확장 프로그램] → [Apps Script]
 * 2. 원래 코드를 모두 지우고 이 파일 내용을 통째로 붙여 넣는다 → 저장
 * 3. [배포] → [새 배포] → 톱니바퀴에서 [웹 앱]
 *      실행 사용자: 나 / 액세스 권한이 있는 사용자: 모든 사용자 → [배포] → 권한 승인
 * 4. '웹 앱 URL'(https://script.google.com/macros/s/.../exec)을 복사해 사이트의 config.js 안 BOARD_API_URL 에 붙여 넣는다.
 *
 * 코드를 고친 뒤에는 [배포] → [배포 관리] → 연필 → 버전: 새 버전 → [배포] 해야 반영된다.
 */

const TEACHER_KEY = '2022';   // 선생님 모드 비밀번호 (config.js 의 TEACHER_KEY 와 같게)
const LIMIT = { name: 12, group: 4, cls: 6, situation: 40, action: 40, robot: 16 };
const SENSORS = ['ir', 'light', 'sound', 'touch'];
const BLOCKED = ['바보', '멍청', '죽어', '꺼져', '병신', '시발', '씨발', 'ㅅㅂ', 'ㅂㅅ', '개새', '존나', '좆', '지랄', '닥쳐', '찐따', '애미', '느금'];
const HEAD = ['id', 'time', 'name', 'group', 'cls', 'sensor', 'situation', 'action', 'robot', 'likes', 'wows', 'wants', 'hidden'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'ping') return out({ ok: true, time: new Date().toISOString() });
  return out(listAll(p.key === TEACHER_KEY));
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: '보낸 내용을 읽을 수 없어요.' }); }
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (err) { return out({ ok: false, error: '사람이 많아요. 잠시 뒤 다시 눌러 주세요.' }); }
  try {
    switch (body.action) {
      case 'post':  return out(addPost(body));
      case 'react': return out(react(body));
      case 'hide':  return out(setHidden(body));
      default:      return out({ ok: false, error: '알 수 없는 요청이에요.' });
    }
  } finally { lock.releaseLock(); }
}

function listAll(isTeacher) {
  const posts = rows().filter(r => r.id && (isTeacher || !truthy(r.hidden))).map(r => ({
    id: String(r.id), time: iso(r.time), name: String(r.name), group: String(r.group || ''), cls: String(r.cls || ''),
    sensor: String(r.sensor), situation: String(r.situation), action: String(r.action), robot: String(r.robot || ''),
    likes: Number(r.likes) || 0, wows: Number(r.wows) || 0, wants: Number(r.wants) || 0, hidden: truthy(r.hidden)
  }));
  posts.sort((a, b) => (a.time < b.time ? 1 : -1));
  return { ok: true, posts: posts, teacher: !!isTeacher, now: new Date().toISOString() };
}

function addPost(b) {
  const name = clean(b.name, LIMIT.name), situation = clean(b.situation, LIMIT.situation), action = clean(b.act, LIMIT.action);
  const robot = clean(b.robot, LIMIT.robot), group = clean(b.group, LIMIT.group), cls = clean(b.cls, LIMIT.cls);
  const sensor = SENSORS.indexOf(b.sensor) >= 0 ? b.sensor : '';
  if (!name || !situation || !action || !sensor) return { ok: false, error: '이름, 센서, 상황, 로봇의 움직임을 모두 적어 주세요.' };
  if (bad(name + situation + action + robot)) return { ok: false, error: '고운 말로 다시 써 주세요.' };
  const id = 'p' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
  sheet().appendRow([id, new Date().toISOString(), name, group, cls, sensor, situation, action, robot, 0, 0, 0, false]);
  return { ok: true, id: id };
}

function react(b) {
  const kind = ['like', 'wow', 'want'].indexOf(b.kind) >= 0 ? b.kind : null;
  if (!kind || !b.postId) return { ok: false, error: '잘못된 반응이에요.' };
  const sh = sheet(), data = sh.getDataRange().getValues();
  const col = HEAD.indexOf(kind + 's') + 1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(b.postId)) { sh.getRange(i + 1, col).setValue((Number(data[i][col - 1]) || 0) + 1); return { ok: true }; }
  }
  return { ok: false, error: '아이디어를 찾을 수 없어요.' };
}

function setHidden(b) {
  if (b.key !== TEACHER_KEY) return { ok: false, error: '선생님만 할 수 있어요.' };
  const sh = sheet(), data = sh.getDataRange().getValues();
  const col = HEAD.indexOf('hidden') + 1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(b.postId)) { sh.getRange(i + 1, col).setValue(!!b.hidden); return { ok: true }; }
  }
  return { ok: false, error: '아이디어를 찾을 수 없어요.' };
}

/* ---------- 도우미 ---------- */
function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('ideas');
  if (!sh) { sh = ss.insertSheet('ideas'); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}
function rows() {
  const sh = sheet(), data = sh.getDataRange().getValues();
  return data.slice(1).map(r => { const o = {}; HEAD.forEach((h, i) => o[h] = r[i]); return o; });
}
function clean(s, n) { return String(s || '').replace(/[<>]/g, '').replace(/\s+/g, ' ').trim().slice(0, n); }
function bad(s) { const t = String(s || '').replace(/\s/g, ''); return BLOCKED.some(w => t.indexOf(w) >= 0); }
function truthy(v) { return v === true || String(v).toLowerCase() === 'true'; }
function iso(v) { try { return (v instanceof Date ? v : new Date(v)).toISOString(); } catch (e) { return String(v); } }
function out(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
