/**
 * 인융대의 밤 입장 체크 - Google Apps Script (독립 프로젝트, SHEET_ID 시트에 기록)
 *
 * 시트 구성 (setup() 한 번 실행하면 자동 생성)
 *  - 설정   : 인융대 학과코드 목록 (학번 5~7번째 자리 → 학부명)
 *  - 명단   : 학생회비 납부자 명단 (학번 / 이름 / 학부). 여기 있으면 '납부자'
 *  - 입장기록: 시각 / 학번 / 이름 / 학부 / 결과 / 기기 / QR 원문 / 학생회비
 *  - 설정 E:F : 미납부자 입장 시각 (예: 16:45). 이 시각 전에는 납부자만 입장
 *
 * 결과 종류
 *  - 입장   : 인융대 소속, 첫 입장
 *  - 중복   : 이미 '입장' 기록이 있는 학번 (재수령 시도)
 *  - 비대상 : 인융대 소속이 아님
 *  - 대기   : 학생회비 미납부자가 미납부자 입장 시각 전에 온 경우 (입장 처리 안 됨, 나중에 다시 찍으면 됨)
 */

const SECRET = 'inbam-g1lywbsgx4ou';   // index.html 의 KEY 와 같아야 함
const SHEET_ID = '1utCt29aB1XHD0BblLhy_uyVmSjJEPDxLbsrlBQBopzM';   // 연결된 스프레드시트 (김세연 개인 계정)
const TZ = 'Asia/Seoul';
const SHEET_LOG = '입장기록';
const SHEET_CONF = '설정';
const SHEET_ROSTER = '명단';

/** 처음 한 번만 실행: 시트와 머리글을 만든다 */
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const mk = (name, header) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(header);
      sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#eef1f6');
      sh.setFrozenRows(1);
    }
    return sh;
  };
  const conf = mk(SHEET_CONF, ['학과코드(학번 5~7번째 자리)', '학부명', '메모']);
  if (conf.getLastRow() === 1) {
    conf.appendRow(["'403", '소프트웨어학부', '확인됨(예: 2024403xxx)']);
    conf.appendRow(["'", '컴퓨터정보공학부', '코드 입력 필요']);
    conf.appendRow(["'", '정보융합학부', '코드 입력 필요']);
    conf.appendRow(["'", '로봇학부', '코드 입력 필요']);
  }
  mk(SHEET_ROSTER, ['학번', '이름', '학부']);
  const log = mk(SHEET_LOG, ['시각', '학번', '이름', '학부', '결과', '기기', 'QR 원문']);
  log.getRange('B:B').setNumberFormat('@');
}

/** v2: 납부자 우선 입장 설정칸과 '학생회비' 열 추가 (한 번 실행) */
function setupV2() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const conf = ss.getSheetByName(SHEET_CONF);
  conf.getRange('E1:F1').setValues([['항목', '값']]).setFontWeight('bold').setBackground('#eef1f6');
  conf.getRange('E2:F2').setValues([['미납부자 입장 시각', "'16:45"]]);
  conf.getRange('E3:F3').setValues([['(참고) 납부자 입장 시작', "'16:30"]]);
  const log = ss.getSheetByName(SHEET_LOG);
  log.getRange('H1').setValue('학생회비').setFontWeight('bold').setBackground('#eef1f6');
  const roster = ss.getSheetByName(SHEET_ROSTER);
  roster.getRange('D1').setValue('← 학생회비 납부자 명단 (학번은 텍스트로)');
  roster.getRange('A:A').setNumberFormat('@');
}

/** 인융대 학과코드 입력 (출처: noeyeas/hasung-welfare-kiosk js/script.js 학번 검증 규칙) */
function setDeptCodes() {
  const conf = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_CONF);
  conf.getRange('A2:C12').clearContent();
  conf.getRange('A2:A10').setNumberFormat('@');
  conf.getRange('A2:C10').setValues([
    ['202', '컴퓨터정보공학부', '고학번'],
    ['203', '소프트웨어학부', '고학번'],
    ['204', '정보융합학부', '고학번'],
    ['205', '로봇학부', '고학번'],
    ['402', '컴퓨터정보공학부', '신입생'],
    ['403', '소프트웨어학부', '신입생'],
    ['404', '정보융합학부', '신입생'],
    ['405', '로봇학부', '신입생'],
    ['406', '로봇학부', '신입생']
  ]);
}

function doGet(e) { return handle_(e.parameter || {}); }

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { p = e.parameter || {}; }
  return handle_(p);
}

function handle_(p) {
  if (p.key !== SECRET) return json_({ status: 'error', message: '인증 키가 맞지 않아요' });
  if (p.action === 'stats') return json_(stats_());

  const raw = String(p.qr || '').trim();
  const q = parseQr_(raw);
  if (!q) return json_({ status: 'unreadable', message: 'QR에서 학번을 읽지 못했어요', raw: raw });
  const sid = q.sid;
  if (!q.manual && QR_MAX_AGE_MIN > 0 && ageMinutes_(q.issued) > QR_MAX_AGE_MIN) {
    return json_({ status: 'expired', sid: sid, issued: q.issuedText, message: '발급된 지 ' + QR_MAX_AGE_MIN + '분이 지난 QR이에요' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const log = ss.getSheetByName(SHEET_LOG);
    const info = lookup_(sid);
    const prev = info.member ? findFirstEntry_(log, sid) : null;

    const nowHM = Utilities.formatDate(new Date(), TZ, 'HH:mm');
    const openAt = unpaidOpenAt_();

    let status = 'ok';
    if (!info.member) status = 'not_member';
    else if (prev) status = 'duplicate';
    else if (!info.paid && openAt && nowHM < openAt) status = 'wait';

    const label = { ok: '입장', duplicate: '중복', not_member: '비대상', wait: '대기' }[status];
    const now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    const feeText = info.member ? (info.paid ? '납부' : '미납') : '';
    log.appendRow([now, "'" + sid, info.name || '', info.dept || '', label + (q.manual ? '(수기)' : ''), String(p.device || ''), raw, feeText]);

    return json_({
      status: status,
      sid: sid,
      name: info.name || '',
      dept: info.dept || '',
      firstTime: prev ? prev.time : '',
      paid: !!info.paid,
      openAt: openAt || '',
      count: countEntries_(log)
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 도서관 모바일 이용증 QR 형식 (2026-11 확인)
 *   "_KW" + 공백 + "0" + 학번(10자리) + 발급시각(yyyyMMddHHmmss) + 숫자 5자리
 *   예) _KW    0 2024403032 20261118213913 12579  (실제로는 붙어 있음)
 * 학번만 직접 입력한 경우(10자리 숫자)도 받아준다.
 */
const QR_MAX_AGE_MIN = 10;   // 발급 후 이 시간(분)이 지난 QR은 '만료'로 처리 (캡처 돌려쓰기 방지). 0이면 검사 안 함

function parseQr_(raw) {
  if (!raw) return null;
  const s = String(raw).trim();
  let m = s.match(/^_KW\s*0?(\d{10})(\d{14})\d*$/);
  if (m) {
    const t = m[2];
    const issued = new Date(+t.slice(0, 4), +t.slice(4, 6) - 1, +t.slice(6, 8), +t.slice(8, 10), +t.slice(10, 12), +t.slice(12, 14));
    return { sid: m[1], issuedText: t.slice(8, 10) + ':' + t.slice(10, 12), issued: issued, manual: false };
  }
  m = s.match(/^0?(20\d{8})$/);
  if (m) return { sid: m[1], manual: true };
  return null;
}

/** 발급 후 경과 분 (스크립트 시간대를 Asia/Seoul로 두고 계산) */
function ageMinutes_(issued) {
  const nowStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  const n = nowStr.match(/(\d+)-(\d+)-(\d+) (\d+):(\d+):(\d+)/);
  const now = new Date(+n[1], +n[2] - 1, +n[3], +n[4], +n[5], +n[6]);
  return (now - issued) / 60000;
}

/** 소속 판정은 학과코드, 학생회비 납부 여부는 명단(납부자 명단) 기준 */
function lookup_(sid) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let name = '', dept = '', paid = false;
  const roster = ss.getSheetByName(SHEET_ROSTER);
  if (roster && roster.getLastRow() > 1) {
    const rows = roster.getRange(2, 1, roster.getLastRow() - 1, 3).getDisplayValues();
    const hit = rows.find(r => String(r[0]).replace(/\D/g, '') === sid);
    if (hit) { paid = true; name = hit[1]; dept = hit[2]; }
  }
  const conf = ss.getSheetByName(SHEET_CONF);
  const codes = conf.getLastRow() > 1 ? conf.getRange(2, 1, conf.getLastRow() - 1, 2).getDisplayValues() : [];
  const code = sid.substr(4, 3);
  const c = codes.find(r => String(r[0]).replace(/\D/g, '') === code);
  const member = paid || !!c;
  if (!dept && c) dept = c[1];
  return member ? { member: true, paid: paid, name: name, dept: dept } : { member: false };
}

/** 설정 F2의 미납부자 입장 시각 (HH:mm). 비어 있으면 제한 없음 */
function unpaidOpenAt_() {
  const v = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_CONF).getRange('F2').getDisplayValue().trim();
  const m = v.match(/^(\d{1,2}):(\d{2})$/);
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : '';
}

/** '입장' 결과로 이미 기록된 학번이 있는지 */
function findFirstEntry_(log, sid) {
  const n = log.getLastRow() - 1;
  if (n < 1) return null;
  const rows = log.getRange(2, 1, n, 5).getDisplayValues();
  const hit = rows.find(r => String(r[1]).replace(/^'/, '') === sid && String(r[4]).indexOf('입장') === 0);
  return hit ? { time: hit[0].slice(11, 16) } : null;
}

function countEntries_(log) {
  const n = log.getLastRow() - 1;
  if (n < 1) return 0;
  return log.getRange(2, 5, n, 1).getDisplayValues().filter(r => String(r[0]).indexOf('입장') === 0).length;
}

function stats_() {
  const log = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_LOG);
  const n = log.getLastRow() - 1;
  const res = n > 0 ? log.getRange(2, 5, n, 1).getDisplayValues().map(r => r[0]) : [];
  return {
    status: 'stats',
    entered: res.filter(x => String(x).indexOf('입장') === 0).length,
    duplicate: res.filter(x => String(x).indexOf('중복') === 0).length,
    notMember: res.filter(x => String(x).indexOf('비대상') === 0).length,
    waiting: res.filter(x => String(x).indexOf('대기') === 0).length
  };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
