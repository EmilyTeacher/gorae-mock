/**
 * 모의고사 페이지 학습 기록 → GORAE_DRILL_DATA
 *   mock_events   : 문제 1개 풀 때마다 1행 (틀린 문제 재풀이 포함, round 열)
 *   mock_complete : 확인 퀴즈의 모든 문제를 맞혀 "완료"된 지문 1건당 1행 (GORAE LIVE·원장·선생님 화면용)
 *
 * 독립(standalone) Apps Script 웹앱. 실행: 나(gorae20200102@gmail.com), 액세스: 모든 사용자.
 * 기존 drill_events / vocab_test_events 는 읽지도 쓰지도 않는다.
 *
 * 저장 조건
 *  - student_name 이 학습관리 '학생명단' 탭의 '학생명' 열에 있거나, 테스트 계정(TEST_NAMES)일 것
 *  - 같은 event_id 는 한 번만 저장
 * 테스트 계정은 기록은 남지만 LIVE·원장·선생님 집계에서는 제외한다(TEST_NAMES 이름으로 거른다).
 */
var DRILL_DATA_ID = '10oRTZ9iIj2f30rR0qwdQ8xX9DVhj1fUXSMpcj0c9iPQ'; // GORAE_DRILL_DATA
var ROSTER_SS_ID = '1LHdBJK8FeFKK6lq5qJ38rZwZg4LSvinGTCmvOPycFds';  // 2026년 학습관리
var ROSTER_SHEET = '학생명단';
var ROSTER_NAME_HEADER = '학생명';
var TEST_NAMES = ['모의고사연결테스트'];

var ANSWER_SHEET = 'mock_events';
var ANSWER_HEADERS = ['event_id', 'attempt_id', 'student_name', 'passage_no', 'passage_title',
  'section', 'q_index', 'is_correct', 'picked', 'answered_at', 'received_at', 'round'];
var COMPLETE_SHEET = 'mock_complete';
var COMPLETE_HEADERS = ['event_id', 'attempt_id', 'student_name', 'passage_no', 'passage_title',
  'rounds', 'first_try_correct', 'total', 'started_at', 'completed_at', 'received_at'];
var SECTIONS = ['solve', 'vocab', 'grammar', 'tf', 'write'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var b = JSON.parse(e.postData.contents);
    var err = validate_(b);
    if (err) return json_({ ok: false, error: err });
    lock.waitLock(20000);
    var isDone = b.event_type === 'mock_complete';
    var sheet = sheet_(isDone ? COMPLETE_SHEET : ANSWER_SHEET, isDone ? COMPLETE_HEADERS : ANSWER_HEADERS);
    if (exists_(sheet, b.event_id)) return json_({ ok: true, duplicate: true });
    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm:ssXXX");
    var headers = isDone ? COMPLETE_HEADERS : ANSWER_HEADERS;
    var row = headers.map(function (h) { return h === 'received_at' ? now : (b[h] === undefined ? '' : b[h]); });
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
  if (!b) return 'body';
  if (b.event_type !== 'mock_answer' && b.event_type !== 'mock_complete') return 'event_type';
  if (!/^[\w:\-]{8,100}$/.test(String(b.event_id || ''))) return 'event_id';
  if (!/^M\d{1,3}$/.test(String(b.passage_no || ''))) return 'passage_no';
  var name = String(b.student_name || '').trim();
  if (!name || name.length > 20) return 'student_name';
  if (!inRoster_(name)) return 'unknown_student';
  b.student_name = name;
  if (b.event_type === 'mock_answer') {
    if (SECTIONS.indexOf(b.section) < 0) return 'section';
    var q = Number(b.q_index); if (!(q >= 1 && q <= 100)) return 'q_index';
    if (typeof b.is_correct !== 'boolean') return 'is_correct';
    b.q_index = q;
    b.round = Number(b.round) >= 1 ? Number(b.round) : 1;
  } else {
    var total = Number(b.total), first = Number(b.first_try_correct), rounds = Number(b.rounds);
    if (!(total >= 1 && total <= 200)) return 'total';
    if (!(first >= 0 && first <= total)) return 'first_try_correct';
    if (!(rounds >= 1 && rounds <= 50)) return 'rounds';
    b.total = total; b.first_try_correct = first; b.rounds = rounds;
  }
  return '';
}

// 학생명단 탭의 '학생명' 열(없으면 거부) + 테스트 계정. 10분 캐시.
function inRoster_(name) {
  if (TEST_NAMES.indexOf(name) >= 0) return true;
  var cache = CacheService.getScriptCache();
  var raw = cache.get('roster_names');
  var names;
  if (raw) {
    names = JSON.parse(raw);
  } else {
    var sh = SpreadsheetApp.openById(ROSTER_SS_ID).getSheetByName(ROSTER_SHEET);
    if (!sh) throw new Error(ROSTER_SHEET + ' 시트를 찾을 수 없습니다.');
    var vals = sh.getDataRange().getValues();
    var col = vals[0].map(function (h) { return String(h).trim(); }).indexOf(ROSTER_NAME_HEADER);
    if (col < 0) throw new Error(ROSTER_SHEET + ' 에서 "' + ROSTER_NAME_HEADER + '" 열을 찾을 수 없습니다.');
    names = [];
    for (var r = 1; r < vals.length; r++) {
      var s = String(vals[r][col]).trim();
      if (s) names.push(s);
    }
    cache.put('roster_names', JSON.stringify(names), 600);
  }
  return names.indexOf(name) >= 0;
}

function sheet_(name, headers) {
  var ss = SpreadsheetApp.openById(DRILL_DATA_ID);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < headers.length) {
    // 열이 추가된 경우 뒤쪽 헤더만 채운다 (기존 열·행은 그대로)
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
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

/** 배포 전 한 번 실행: 두 탭과 헤더만 만든다(기존 행은 건드리지 않음). */
function setupMockSheet() {
  sheet_(ANSWER_SHEET, ANSWER_HEADERS);
  sheet_(COMPLETE_SHEET, COMPLETE_HEADERS);
}
