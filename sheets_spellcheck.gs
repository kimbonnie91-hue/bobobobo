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


/** 스프레드시트 열릴 때 메뉴 생성 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔤 맞춤법 검사')
    .addItem('선택 영역 검사', 'checkSelection')
    .addSeparator()
    .addItem('표시 지우기(선택 영역)', 'clearSelectionMarks')
    .addItem('표시 지우기(현재 시트 전체)', 'clearSheetMarks')
    .addToUi();
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

      if (result.errors.length > 0) {
        bgs[r][c] = ERR_BG;
        notes[r][c] = buildNote(result.errors);
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
