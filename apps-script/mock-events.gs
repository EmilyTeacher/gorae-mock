/**
 * 모의고사 페이지 학습 기록 → GORAE_DRILL_DATA 의 mock_events 탭
 *
 * 독립(standalone) Apps Script 프로젝트로 만들어 웹앱으로 배포한다.
 *  - 실행 계정: 나(gorae20200102@gmail.com), 액세스: 모든 사용자
 *  - 기존 drill_events / vocab_test_events 는 읽지도 쓰지도 않는다.
 *  - 이름이 학생명단에 없으면 저장하지 않는다. 같은 event_id는 한 번만 저장한다.
 */
var DRILL_DATA_ID = '10oRTZ9iIj2f30rR0qwdQ8xX9DVhj1fUXSMpcj0c9iPQ'; // GORAE_DRILL_DATA
var ROSTER_SS_ID = '1LHdBJK8FeFKK6lq5qJ38rZwZg4LSvinGTCmvOPycFds';  // 2026년 학습관리
var ROSTER_SHEET = '학생명단';
var MOCK_SHEET = 'mock_events';
var HEADERS = ['event_id', 'attempt_id', 'student_name', 'passage_no', 'passage_title',
  'section', 'q_index', 'is_correct', 'picked', 'answered_at', 'received_at'];
var SECTIONS = ['solve', 'vocab', 'grammar', 'tf', 'write'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var b = JSON.parse(e.postData.contents);
    var err = validate_(b);
    if (err) return json_({ ok: false, error: err });
    lock.waitLock(20000);
    var sheet = sheet_();
    if (exists_(sheet, b.event_id)) return json_({ ok: true, duplicate: true });
    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm:ssXXX");
    var row = HEADERS.map(function (h) { return h === 'received_at' ? now : (b[h] === undefined ? '' : b[h]); });
    sheet.appendRow(row);
    return json_({ ok: true });
  } catch (x) {
    return json_({ ok: false, error: String(x) });
  } finally {
    try { lock.releaseLock(); } catch (y) {}
  }
}

function doGet() { return json_({ ok: true, service: 'mock_events' }); }

function validate_(b) {
  if (!b || b.event_type !== 'mock_answer') return 'event_type';
  if (!/^[\w:\-]{8,100}$/.test(String(b.event_id || ''))) return 'event_id';
  if (SECTIONS.indexOf(b.section) < 0) return 'section';
  if (!/^M\d{1,3}$/.test(String(b.passage_no || ''))) return 'passage_no';
  var q = Number(b.q_index); if (!(q >= 1 && q <= 100)) return 'q_index';
  if (typeof b.is_correct !== 'boolean') return 'is_correct';
  var name = String(b.student_name || '').trim();
  if (!name || name.length > 20) return 'student_name';
  if (!inRoster_(name)) return 'unknown_student';
  b.student_name = name;
  b.q_index = q;
  return '';
}

// 학생명단 탭 어디에든 이름이 있으면 통과 (10분 캐시)
function inRoster_(name) {
  var cache = CacheService.getScriptCache();
  var raw = cache.get('roster');
  var names;
  if (raw) {
    names = JSON.parse(raw);
  } else {
    var sh = SpreadsheetApp.openById(ROSTER_SS_ID).getSheetByName(ROSTER_SHEET);
    if (!sh) throw new Error(ROSTER_SHEET + ' 시트를 찾을 수 없습니다.');
    names = [];
    sh.getDataRange().getValues().forEach(function (r) {
      r.forEach(function (c) { var s = String(c).trim(); if (s && s.length <= 20) names.push(s); });
    });
    cache.put('roster', JSON.stringify(names), 600);
  }
  return names.indexOf(name) >= 0;
}

function sheet_() {
  var ss = SpreadsheetApp.openById(DRILL_DATA_ID);
  var sh = ss.getSheetByName(MOCK_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MOCK_SHEET);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function exists_(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var ids = sheet.getRange(Math.max(2, last - 2000), 1, Math.min(last - 1, 2001), 1).getValues();
  for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return true;
  return false;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** 배포 전 한 번 실행: mock_events 탭과 헤더만 만든다. */
function setupMockSheet() { sheet_(); }
