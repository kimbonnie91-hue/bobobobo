/**
 * ============================================================================
 *  구글시트 한국어 맞춤법 검사 (요청문구 제목/내용 셀 검수용)
 * ============================================================================
 *
 *  동작 방식
 *  ---------
 *  - 시트 상단에 "🔤 맞춤법 검사" 메뉴가 추가됩니다.
 *  - 검사할 셀(제목/내용 열)을 마우스로 선택한 뒤 [선택 영역 검사]를 누르면,
 *    다음(Daum) 한국어 맞춤법 검사기 엔진으로 각 셀을 검사합니다.
 *  - 오탈자가 있는 셀은 연한 빨강으로 강조되고, 셀 메모(마우스 올리면 뜨는 노트)에
 *    "원문 → 수정후보 (설명)" 목록이 달립니다.  ※ 셀 원본 텍스트는 건드리지 않습니다.
 *  - [표시 지우기]로 이 스크립트가 남긴 강조/메모를 한번에 제거할 수 있습니다.
 *
 *  주의
 *  ----
 *  - 다음 검사기는 비공식 엔드포인트라 느리거나 일시적으로 막힐 수 있습니다.
 *    "완벽 자동 교정"이 아니라 "의심 구간 표시 → 사람이 확인" 용도로 쓰세요.
 *  - #{고객명} 같은 치환 변수, URL, 이모지는 오탈자 오탐을 막기 위해 검사 전에 제거합니다.
 *  - 서버 부하/차단을 피하려고 셀마다 약간의 대기(SLEEP_MS)를 둡니다. 셀이 많으면
 *    시간이 걸립니다. 한 번에 200셀 이하로 나눠서 검사하는 것을 권장합니다.
 *
 *  설치 방법은 같은 저장소의 SHEETS_SPELLCHECK_설치가이드.md 를 참고하세요.
 * ============================================================================
 */

// 다음(Daum) 맞춤법 검사기 엔드포인트 (HTTPS).
var SPELLER_URL = 'https://dic.daum.net/grammar_checker.do';

// 서버가 브라우저 요청으로 인식하도록 User-Agent 지정.
var UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
         '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

// 한 요청당 최대 글자 수(다음 검사기 제한). 초과 시 문장 단위로 나눠 검사.
var MAX_CHARS = 1000;

// 셀 사이 대기(ms). 서버 차단 방지용. 너무 짧으면 차단 위험.
var SLEEP_MS = 500;

// 한 번에 검사 허용할 최대 셀 수(안전장치).
var MAX_CELLS = 500;

// 오류 셀 강조 색.
var ERR_BG = '#FCE8E6'; // 연한 빨강
var OK_BG  = null;      // 정상 셀은 배경 초기화(원래대로)

// 이 스크립트가 남긴 메모임을 표시하는 접두어(표시 지우기 때 이 메모만 제거).
var NOTE_TAG = '[맞춤법]';

// 사용자 사전(브랜드 오타 등) 탭 이름. [틀린표현 | 올바른표현 | 메모]
var DICT_SHEET = '맞춤법사전';


/** 스프레드시트 열릴 때 메뉴 생성 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔤 맞춤법 검사')
    .addItem('선택 영역 검사', 'checkSelection')
    .addSeparator()
    .addItem('사용자 사전 만들기/열기', 'openDict')
    .addSeparator()
    .addItem('표시 지우기(선택 영역)', 'clearSelectionMarks')
    .addItem('표시 지우기(현재 시트 전체)', 'clearSheetMarks')
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu('🔁 AF코드 중복')
    .addItem('전체 검사(각 주차 탭 내부)', 'checkAllAfDuplicates')
    .addItem('AF 중복 표시 지우기(현재 시트)', 'clearAfMarksSheet')
    .addSeparator()
    .addItem('🔧 진단(왜 경고가 안 뜨나)', 'afDiag')
    .addItem('캐시 초기화(열/마스터 변경 후)', 'clearAfCache')
    .addToUi();
}


/** 사용자 사전 탭을 만들거나(없으면) 열기 */
function openDict() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DICT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DICT_SHEET);
    sh.getRange('A1:C1')
      .setValues([['틀린표현', '올바른표현', '메모']])
      .setFontWeight('bold').setBackground('#FFF2CC');
    sh.getRange('A2:C4').setValues([
      ['카즈', '카츠', '브랜드명'],
      ['썸머', '서머', '외래어 표기'],
      ['오르뜨', '오르뜨', '(예시 - 삭제하고 실제 규칙 입력)']
    ]);
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 3, 160);
  }
  ss.setActiveSheet(sh);
  ss.toast('사용자 사전 탭입니다. 틀린표현/올바른표현/메모를 채우세요.', '맞춤법 검사', 6);
}


/** 사용자 사전 규칙 로드 → [{find, to, note}] */
function loadCustomRules() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DICT_SHEET);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 3).getValues();
  var rules = [];
  for (var i = 0; i < rows.length; i++) {
    var find = String(rows[i][0] == null ? '' : rows[i][0]).trim();
    if (find === '') continue;
    rules.push({
      find: find,
      to: String(rows[i][1] == null ? '' : rows[i][1]).trim(),
      note: String(rows[i][2] == null ? '' : rows[i][2]).trim()
    });
  }
  return rules;
}


/** 원문에서 사용자 사전 규칙에 걸리는 표현 찾기 → [{orgStr, candWord, help}] */
function applyCustomRules(rawText, rules) {
  var out = [];
  if (!rules || rules.length === 0) return out;
  var text = String(rawText);
  for (var i = 0; i < rules.length; i++) {
    var r = rules[i];
    if (text.indexOf(r.find) === -1) continue;
    // 틀린표현==올바른표현 이면 알림용 예시이므로 건너뜀
    if (r.to && r.to === r.find) continue;
    out.push({
      orgStr: r.find,
      candWord: r.to || '',
      help: (r.note ? r.note + ' ' : '') + '(사용자 사전)'
    });
  }
  return out;
}


/** 선택한 셀들을 검사 */
function checkSelection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var range = ss.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('검사할 셀을 먼저 선택하세요.');
    return;
  }

  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var total = numRows * numCols;
  if (total > MAX_CELLS) {
    SpreadsheetApp.getUi().alert(
      '선택한 셀이 ' + total + '개입니다. 한 번에 ' + MAX_CELLS +
      '개 이하로 나눠서 검사하세요.');
    return;
  }

  var values = range.getValues();
  var bgs   = range.getBackgrounds();
  var notes = range.getNotes();

  var rules = loadCustomRules(); // 사용자 사전(브랜드 오타 등)

  var errCells = 0;
  var checked = 0;

  for (var r = 0; r < numRows; r++) {
    for (var c = 0; c < numCols; c++) {
      var text = values[r][c];
      if (text == null || String(text).trim() === '') {
        continue; // 빈 셀 건너뜀
      }
      checked++;

      var result;
      try {
        result = spellCheck(String(text));
      } catch (e) {
        // 한 셀 실패해도 전체 중단하지 않음. 메모로 남김.
        notes[r][c] = NOTE_TAG + ' 검사 실패: ' + e.message;
        bgs[r][c] = ERR_BG;
        errCells++;
        Utilities.sleep(SLEEP_MS);
        continue;
      }

      // 다음 검사기 결과 + 사용자 사전 결과 병합
      var allErrors = result.errors.concat(applyCustomRules(String(text), rules));

      if (allErrors.length > 0) {
        bgs[r][c] = ERR_BG;
        notes[r][c] = buildNote(allErrors);
        errCells++;
      } else {
        // 오류 없음: 이 스크립트가 남긴 흔적만 정리(사용자 원래 서식/메모는 유지)
        if (isOurNote(notes[r][c])) notes[r][c] = '';
        if (bgs[r][c] === ERR_BG) bgs[r][c] = OK_BG || null;
      }

      Utilities.sleep(SLEEP_MS);
    }
  }

  range.setBackgrounds(bgs);
  range.setNotes(notes);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '검사 완료: ' + checked + '셀 중 ' + errCells + '셀에서 오류 발견',
    '맞춤법 검사', 6);
}


/** 오류 목록을 셀 메모 텍스트로 조립 */
function buildNote(errors) {
  var lines = [NOTE_TAG + ' 오류 ' + errors.length + '건'];
  for (var i = 0; i < errors.length; i++) {
    var e = errors[i];
    var cand = e.candWord ? e.candWord.split('|').join(' / ') : '(대안 없음)';
    var line = '• ' + e.orgStr + ' → ' + cand;
    if (e.help) line += '\n   ' + e.help;
    lines.push(line);
  }
  return lines.join('\n');
}


/** 우리가 남긴 메모인지 판별 */
function isOurNote(note) {
  return note && note.indexOf(NOTE_TAG) === 0;
}


/** 선택 영역에서 우리가 남긴 표시 제거 */
function clearSelectionMarks() {
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  if (!range) return;
  clearMarksInRange(range);
  SpreadsheetApp.getActiveSpreadsheet().toast('선택 영역 표시 제거 완료', '맞춤법 검사', 4);
}


/** 현재 시트 전체에서 우리가 남긴 표시 제거 */
function clearSheetMarks() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  clearMarksInRange(range);
  SpreadsheetApp.getActiveSpreadsheet().toast('시트 전체 표시 제거 완료', '맞춤법 검사', 4);
}


function clearMarksInRange(range) {
  var bgs = range.getBackgrounds();
  var notes = range.getNotes();
  for (var r = 0; r < notes.length; r++) {
    for (var c = 0; c < notes[r].length; c++) {
      if (isOurNote(notes[r][c])) notes[r][c] = '';
      if (bgs[r][c] === ERR_BG) bgs[r][c] = null;
    }
  }
  range.setBackgrounds(bgs);
  range.setNotes(notes);
}


/**
 * 다음(Daum) 맞춤법 검사기 호출.
 * @param {string} rawText 검사할 원문
 * @return {{errors: Array, cleaned: string}}
 */
function spellCheck(rawText) {
  var cleaned = sanitize(rawText);
  if (cleaned.trim() === '') return { errors: [], cleaned: cleaned };

  var chunks = splitByLength(cleaned, MAX_CHARS);
  var errors = [];

  for (var i = 0; i < chunks.length; i++) {
    var part = chunks[i].trim();
    if (part === '') continue;
    if (i > 0) Utilities.sleep(SLEEP_MS); // 청크 간 간격

    var res = UrlFetchApp.fetch(SPELLER_URL, {
      method: 'post',
      payload: { sentence: part },
      headers: { 'User-Agent': UA },
      followRedirects: true,
      muteHttpExceptions: true,
      contentType: 'application/x-www-form-urlencoded; charset=UTF-8'
    });

    var code = res.getResponseCode();
    if (code !== 200) {
      throw new Error('HTTP ' + code + ' (검사 서버 응답 오류)');
    }

    var body = res.getContentText('UTF-8');
    if (body.indexOf('맞춤법 검사기 본문') === -1) {
      throw new Error('검사기 응답 형식 오류(서비스 변경/차단 가능)');
    }

    errors = errors.concat(parseDaum(body));
  }

  return { errors: errors, cleaned: cleaned };
}


/** 다음 응답 HTML에서 오탈자 파싱 → [{orgStr, candWord, help}] */
function parseDaum(response) {
  var typos = [];
  var found = -1;

  for (;;) {
    found = response.indexOf('data-error-type', found + 1);
    if (found === -1) break;

    var end = response.indexOf('>', found + 1);
    var line = response.substring(found, end);

    var orgStr = decodeEntities(getAttr(line, 'data-error-input='));
    var cand   = decodeEntities(getAttr(line, 'data-error-output='));

    var help = '';
    try {
      var infoBegin = response.indexOf('<div>', found);
      if (infoBegin !== -1) {
        var infoEnd = response.indexOf('</div>', infoBegin + 1);
        if (infoEnd !== -1) {
          help = decodeEntities(response.substring(infoBegin, infoEnd + 6))
            .replace(/\t/g, '')
            .replace(/<br[^>]*>/gi, ' ')
            .replace(/<[^>]*>/g, '')
            .replace(/\s+/g, ' ')
            .trim();
          if (help === '도움말이 없습니다.') help = '';
        }
      }
    } catch (e) {
      help = '';
    }

    if (orgStr) {
      typos.push({ orgStr: orgStr, candWord: cand, help: help });
    }
  }
  return typos;
}


/** HTML 속성값 추출: key="값" 에서 값만 반환 */
function getAttr(str, key) {
  var found = str.indexOf(key);
  if (found === -1) return '';
  var firstQuote = str.indexOf('"', found + 1);
  if (firstQuote === -1) return '';
  var secondQuote = str.indexOf('"', firstQuote + 1);
  if (secondQuote === -1) return '';
  return str.substring(firstQuote + 1, secondQuote);
}


/** 긴 문장을 MAX_CHARS 이하로 문장부호 기준 분할 */
function splitByLength(text, maxLen) {
  if (text.length <= maxLen) return [text];
  var out = [];
  var s = text;
  while (s.length > maxLen) {
    var cut = s.lastIndexOf('.', maxLen);
    if (cut < maxLen * 0.5) cut = s.lastIndexOf(' ', maxLen);
    if (cut < maxLen * 0.5) cut = maxLen;
    out.push(s.substring(0, cut));
    s = s.substring(cut);
  }
  if (s.trim() !== '') out.push(s);
  return out;
}


/**
 * 검사 전 정리: 치환 변수/URL/이모지 등 오탐 요소 제거.
 * (인덱스 정합이 필요 없는 '표시' 방식이라 단순 제거로 충분)
 */
function sanitize(text) {
  var t = String(text);
  t = t.replace(/#\{[^}]*\}/g, ' ');          // #{고객명} 등 치환 변수
  t = t.replace(/\{[^}]*\}/g, ' ');           // {변수}
  t = t.replace(/https?:\/\/\S+/g, ' ');      // URL
  t = t.replace(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+/g, ' '); // 이메일
  // 이모지/기호(비한글·비영문·비숫자·비기본문장부호) 제거
  t = t.replace(
    /[\u{1F000}-\u{1FAFF}\u{2600}-\u{27BF}\u{2190}-\u{21FF}\u{2B00}-\u{2BFF}\u{FE00}-\u{FE0F}\u{200D}\u{E000}-\u{F8FF}]/gu,
    ' ');
  t = t.replace(/\s+/g, ' ').trim();
  return t;
}


/** 자주 쓰는 HTML 엔티티 디코드 */
function decodeEntities(s) {
  return String(s)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#0*39;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(parseInt(n, 10)); })
    .replace(/&#x([0-9a-fA-F]+);/g, function (_, n) { return String.fromCharCode(parseInt(n, 16)); })
    .replace(/&amp;/g, '&');
}


/* ============================================================================
 *  AF코드 중복 검사 (같은 주차 탭 '내부'에서만 중복 방지)
 * ============================================================================
 *
 *  - 주차 탭(이름에 '주차' 포함)의 'AF코드' 열에 코드를 입력/붙여넣는 순간,
 *    '그 탭 안에서만' 같은 코드가 이미 쓰였으면 즉시 팝업 + 셀 메모로 경고.
 *    (다른 주차 탭은 검사하지 않음 — 주차가 다르면 같은 코드 사용 가능)
 *  - 마스터 풀(PUSH AF코드(마케팅)) 목록에 없는 코드(오타 의심)도 경고.
 *  - 아래 상수로 동작을 조정할 수 있습니다.
 * ------------------------------------------------------------------------- */

var AFCODE_HEADER = 'AF코드';               // 주차 탭에서 이 헤더가 있는 열을 검사
var WEEKLY_TAB_PATTERN = '주차';            // 이 문자열이 든 탭을 '주차 탭'으로 취급
var MASTER_TAB = 'PUSH AF코드(마케팅)';      // 마스터 풀 탭 이름
var AFCODE_PATTERN = /^[A-Z]{2,4}\d{2,4}$/; // AF코드 형태(AP02, PB15, APZ01, AP300...)

var ONLY_WEEKLY_TABS = false;        // true면 '주차' 든 탭만 검사, false면 AF코드 열 있는 모든 탭 검사(권장)
var CHECK_MASTER_MEMBERSHIP = false; // 마스터 풀 대조(오타 경고). 실시간 속도 위해 기본 OFF. 켜면 캐시로 동작
var AF_USE_BG = true;               // 셀 배경 강조 사용 (AF코드 열에 기존 수동 채우기색이 있으면 false 권장)

var AF_NOTE_TAG = '[AF중복]';
var AF_ERR_BG  = '#FF4D4D'; // 중복(또렷한 빨강)
var AF_WARN_BG = '#FCE5CD'; // 마스터풀에 없음(주황)


/** 편집 시 실시간 AF코드 중복 검사 (설치형 트리거 불필요, 단순 트리거) */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    var name = sh.getName();
    if (ONLY_WEEKLY_TABS && name.indexOf(WEEKLY_TAB_PATTERN) === -1) return;

    var afCol = getAfColCached(sh); // 캐시된 열 번호 (TextFinder 반복 방지)
    if (afCol === -1) return; // AF코드 열 없는 탭이면 검사 안 함

    var c0 = e.range.getColumn();
    var nC = e.range.getNumColumns();
    if (afCol < c0 || afCol > c0 + nC - 1) return; // AF코드 열 편집 아님

    var ss = sh.getParent();
    var last = sh.getLastRow();
    var colVals = sh.getRange(1, afCol, last, 1).getValues(); // AF열 1회만 읽기
    var index = {};
    for (var k = 0; k < colVals.length; k++) {
      var cv = String(colVals[k][0]).trim();
      if (cv && AFCODE_PATTERN.test(cv)) (index[cv] = index[cv] || []).push(k + 1);
    }
    var master = CHECK_MASTER_MEMBERSHIP ? getMasterCodeSetCached(ss) : null;

    var r0 = e.range.getRow();
    var nR = e.range.getNumRows();
    var alerts = [];
    var seen = {}; // 병합 셀이 여러 행에 걸쳐도 앵커 한 번만 처리

    for (var r = r0; r < r0 + nR; r++) {
      var cell = anchorCell(sh, r, afCol); // 병합이면 맨 위 셀
      var rr = cell.getRow();
      if (seen[rr]) continue;
      seen[rr] = true;

      var val = String((colVals[rr - 1] || [''])[0]).trim(); // 위에서 읽은 값 재사용

      if (val === '' || !AFCODE_PATTERN.test(val)) {
        clearAfCell(cell);
        continue;
      }

      var allRows = index[val] || [rr];
      var others = allRows.filter(function (row) { return row !== rr; });

      if (others.length > 0) {
        var note = AF_NOTE_TAG + ' 이 탭에서 중복! 행 ' + allRows.join(', ');
        markAfCell(cell, AF_ERR_BG, note);
        for (var oi = 0; oi < others.length; oi++) {
          markAfCell(anchorCell(sh, others[oi], afCol), AF_ERR_BG, note); // 짝 셀도 빨강
        }
        alerts.push('⚠️ ' + val + ' 중복 → 같은 탭 행 ' + others.join(', '));
      } else if (master && master.size > 0 && !master.has(val)) {
        markAfCell(cell, AF_WARN_BG, AF_NOTE_TAG + ' 마스터 풀에 없는 코드(오타 의심)');
        alerts.push('❓ ' + val + ' : 마스터 풀에 없음');
      } else {
        clearAfCell(cell);
      }
    }

    // 중복을 고쳐서 유일해진 경우, 이전 값(짝)의 빨강도 해제
    if (e.oldValue !== undefined) {
      var ov = String(e.oldValue).trim();
      if (ov && AFCODE_PATTERN.test(ov) && index[ov] && index[ov].length === 1) {
        clearAfCell(anchorCell(sh, index[ov][0], afCol));
      }
    }

    if (alerts.length > 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(alerts.join('\n'), 'AF코드 경고', 8);
    }
  } catch (err) {
    // onEdit는 조용히 실패해야 시트 편집을 방해하지 않음
  }
}


/** 각 주차 탭 '내부'의 기존 중복을 한번에 표시 (초기 점검용) */
function checkAllAfDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = CHECK_MASTER_MEMBERSHIP ? getMasterCodeSet(ss) : null;

  var sheets = ss.getSheets();
  var totalDupCells = 0;
  var dupSummary = []; // '탭명: 코드, 코드'

  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    var nm = sh.getName();
    if (ONLY_WEEKLY_TABS && nm.indexOf(WEEKLY_TAB_PATTERN) === -1) continue;
    var col = findAfCol(sh);
    if (col === -1) continue; // AF코드 열 있는 탭만
    var last = sh.getLastRow();
    if (last < 1) continue;

    var index = buildSheetIndex(sh); // 이 탭 내부만: 코드 -> [행,...]
    var vals = sh.getRange(1, col, last, 1).getValues();
    var dupCodesInTab = {};

    for (var i = 0; i < vals.length; i++) {
      var v = String(vals[i][0]).trim();
      var cell = sh.getRange(i + 1, col);
      if (v === '' || !AFCODE_PATTERN.test(v)) { clearAfCell(cell); continue; }

      if (index[v] && index[v].length > 1) {
        markAfCell(cell, AF_ERR_BG, AF_NOTE_TAG + ' 이 탭에서 중복! 행 ' + index[v].join(', '));
        totalDupCells++;
        dupCodesInTab[v] = true;
      } else if (master && master.size > 0 && !master.has(v)) {
        markAfCell(cell, AF_WARN_BG, AF_NOTE_TAG + ' 마스터 풀에 없는 코드(오타 의심)');
      } else {
        clearAfCell(cell);
      }
    }

    var codes = Object.keys(dupCodesInTab);
    if (codes.length > 0) dupSummary.push('• ' + nm + ': ' + codes.join(', '));
  }

  var msg = dupSummary.length === 0
    ? '탭 내부 중복 없음 ✅'
    : '중복 발견 (셀 ' + totalDupCells + '개):\n' + dupSummary.join('\n');
  SpreadsheetApp.getUi().alert('AF코드 중복 검사 결과(각 주차 탭 내부)', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}


/** 현재 시트에서 AF 중복 표시 제거 */
function clearAfMarksSheet() {
  var sh = SpreadsheetApp.getActiveSheet();
  var col = findAfCol(sh);
  if (col === -1) {
    SpreadsheetApp.getActiveSpreadsheet().toast('이 시트에서 AF코드 열을 못 찾음', 'AF코드', 4);
    return;
  }
  var last = sh.getLastRow();
  if (last < 1) return;
  for (var i = 1; i <= last; i++) clearAfCell(sh.getRange(i, col));
  SpreadsheetApp.getActiveSpreadsheet().toast('AF 중복 표시 제거 완료', 'AF코드', 4);
}


/** 한 시트(주차 탭) 내부의 AF코드 인덱스: 코드 -> [행번호, ...] */
function buildSheetIndex(sh) {
  var index = {};
  var col = findAfCol(sh);
  if (col === -1) return index;
  var last = sh.getLastRow();
  if (last < 1) return index;
  var vals = sh.getRange(1, col, last, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var v = String(vals[i][0]).trim();
    if (v === '' || !AFCODE_PATTERN.test(v)) continue;
    (index[v] = index[v] || []).push(i + 1);
  }
  return index;
}


/** 마스터 풀 탭에서 AF코드 집합 추출 (셀 안에 섞여 있어도 코드만 뽑음) */
function getMasterCodeSet(ss) {
  var set = new Set();
  var sh = ss.getSheetByName(MASTER_TAB);
  if (!sh) return set;
  var last = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (last < 1 || lastCol < 1) return set;
  var vals = sh.getRange(1, 1, last, lastCol).getValues();
  var re = /[A-Z]{2,4}\d{2,4}/g;
  for (var r = 0; r < vals.length; r++) {
    for (var c = 0; c < vals[r].length; c++) {
      var cellStr = String(vals[r][c] == null ? '' : vals[r][c]);
      var m;
      while ((m = re.exec(cellStr)) !== null) set.add(m[0]);
    }
  }
  return set;
}


/** AF코드 열 번호를 캐시 (시트별 6시간). 캐시 실패해도 정상 동작(직접 탐색) */
function getAfColCached(sh) {
  try {
    var cache = CacheService.getScriptCache();
    var key = 'afcol_' + sh.getSheetId();
    var v = cache.get(key);
    if (v !== null) return parseInt(v, 10);
    var col = findAfCol(sh);
    cache.put(key, String(col), 21600);
    return col;
  } catch (e) {
    return findAfCol(sh); // 실시간 트리거 제한 권한 등에서 캐시 실패 시 폴백
  }
}


/** 마스터 풀 코드 집합을 캐시 (1시간). 캐시 실패해도 정상 동작 */
function getMasterCodeSetCached(ss) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('masterset');
    if (cached) return new Set(JSON.parse(cached));
    var set = getMasterCodeSet(ss);
    cache.put('masterset', JSON.stringify(Array.from(set)), 3600);
    return set;
  } catch (e) {
    return getMasterCodeSet(ss);
  }
}


/** 캐시 초기화 (AF열/마스터 풀 변경 후 사용) */
function clearAfCache() {
  var cache = CacheService.getScriptCache();
  cache.remove('masterset');
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var keys = [];
  for (var i = 0; i < sheets.length; i++) keys.push('afcol_' + sheets[i].getSheetId());
  cache.removeAll(keys);
  SpreadsheetApp.getActiveSpreadsheet().toast('캐시 초기화 완료', 'AF코드', 4);
}


/** 시트에서 'AF코드' 헤더가 있는 열 번호 찾기 (없으면 -1). 앞뒤 공백 허용 */
function findAfCol(sh) {
  var matches = sh.createTextFinder(AFCODE_HEADER).findAll();
  for (var i = 0; i < matches.length; i++) {
    if (String(matches[i].getValue()).trim() === AFCODE_HEADER) {
      return matches[i].getColumn();
    }
  }
  return matches.length ? matches[0].getColumn() : -1;
}


/** 진단: 현재 탭이 검사 대상인지, AF코드 열을 찾는지, 탭 내부 중복이 있는지 보고 */
function afDiag() {
  var sh = SpreadsheetApp.getActiveSheet();
  var name = sh.getName();
  var col = findAfCol(sh);
  var nameOk = !ONLY_WEEKLY_TABS || name.indexOf(WEEKLY_TAB_PATTERN) !== -1;
  var willRun = nameOk && col !== -1;

  var msg = '탭 이름: "' + name + '"\n'
    + '· 실시간 검사 동작?: ' + (willRun ? '예 ✅' : '아니오 ❌') + '\n'
    + (ONLY_WEEKLY_TABS ? ("· '" + WEEKLY_TAB_PATTERN + "' 포함: " + (name.indexOf(WEEKLY_TAB_PATTERN) !== -1 ? '예' : '아니오 ← ONLY_WEEKLY_TABS=false로 두면 해결') + '\n') : '')
    + '· AF코드 열: '
    + (col === -1 ? '못 찾음 ❌  ← 헤더가 정확히 "AF코드"인지 확인' : col + '열 (' + colLetter(col) + ') ✅') + '\n';

  if (col !== -1) {
    var idx = buildSheetIndex(sh);
    var codeCount = Object.keys(idx).length;
    var dups = [];
    for (var c in idx) {
      if (idx[c].length > 1) dups.push(c + ' (행 ' + idx[c].join(',') + ')');
    }
    // 실제로 읽은 코드 앞부분 샘플
    var sample = Object.keys(idx).slice(0, 8).join(', ');
    // 현재 선택 셀의 병합 여부
    var act = sh.getActiveCell();
    var merged = act.isPartOfMerge() ? ('예 (앵커 ' + colLetter(anchorCell(sh, act.getRow(), act.getColumn()).getColumn()) + anchorCell(sh, act.getRow(), act.getColumn()).getRow() + ')') : '아니오';

    msg += '· 인식된 코드 수: ' + codeCount + '\n'
        + '· 코드 샘플: ' + (sample || '(없음 ← 읽기 실패)') + '\n'
        + '· 현재 셀 병합됨? ' + merged + '\n'
        + '· 이 탭 중복: ' + (dups.length ? dups.join(' / ') : '없음');
  }

  SpreadsheetApp.getUi().alert('AF코드 진단', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}


/** 병합 셀이면 맨 위(앵커) 셀을 반환, 아니면 그 셀 그대로 */
function anchorCell(sh, row, col) {
  var cell = sh.getRange(row, col);
  if (cell.isPartOfMerge()) {
    var mr = cell.getMergedRanges();
    if (mr && mr.length) return mr[0].getCell(1, 1);
  }
  return cell;
}


/** 열 번호 → 문자(A,B,...) */
function colLetter(col) {
  var s = '';
  while (col > 0) {
    var m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}


/** 셀에 AF 경고 표시 (기존 값/서식은 최대한 보존) */
function markAfCell(cell, bg, note) {
  cell.setNote(note);
  if (AF_USE_BG) cell.setBackground(bg);
}


/** 우리가 남긴 AF 경고 표시만 제거 (사용자 원래 서식/메모는 유지) */
function clearAfCell(cell) {
  var note = cell.getNote();
  if (note && note.indexOf(AF_NOTE_TAG) === 0) cell.setNote('');
  if (AF_USE_BG) {
    var bg = (cell.getBackground() || '').toUpperCase();
    if (bg === AF_ERR_BG.toUpperCase() || bg === AF_WARN_BG.toUpperCase()) {
      cell.setBackground(null);
    }
  }
}
