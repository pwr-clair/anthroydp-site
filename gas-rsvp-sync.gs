/**
 * 연계프로그램 접수 동기화 — 구글 폼 응답 → (1) '접수' 탭  (2) 발송용 비공개 시트
 *
 * 붙이는 곳: "연계프로그램 신청자 현황" 스프레드시트 > 확장 프로그램 > Apps Script
 * 시트1(폼 목록)의 '링크' 열에 있는 폼을 하나씩 열어 응답을 읽는다.
 *
 * (1) 이 시트의 '접수' 탭 — 운영인력 페이지(RSVP 탭)가 읽는 명단.
 *     [프로그램, 성함, 전화 뒷자리, 인원, 접수 시각, 응답 id]  ※ 전화는 뒷자리 4자리만 (시트가 링크 공유라서)
 * (2) 발송용 시트 — 소유자 드라이브에 자동 생성되는 비공개 스프레드시트 ('연계프로그램 발송용 명단').
 *     '요약' 탭: 프로그램별 접수 팀·인원·정원 잔여·발송일 + 전화번호/이메일을 쉼표로 이어 붙인 셀(복사해서 문자·메일 수신자에 붙여넣기)
 *     프로그램별 탭: 성함·전화번호·이메일·인원·접수 시각 (열을 통째로 긁어갈 수 있게)
 *     링크는 '접수' 탭 H3 셀에 적힌다. 소유자만 열 수 있다.
 *
 * 처음 1회: installTrigger 실행(권한 승인) → 5분마다 자동 동기화 + 즉시 1회 실행.
 * 수동 갱신: syncRsvp 실행.
 */
var SRC_SHEET = '시트1';   // 폼 목록 탭 (A=폼 이름, D=링크)
var OUT_SHEET = '접수';    // 앱용 결과 탭 (없으면 만든다)
var OUT_BOOK_NAME = '연계프로그램 발송용 명단 (비공개)';
var CAPACITY = 15;         // 프로그램 정원

/* 폼 이름 앞 번호 → 일시 (홈페이지 프로그램 카드 기준) */
var WHEN = {
  '01': '9/16(수) 18:00', '02': '9/16(수) 19:00', '03': '9/17(목) 12:00 ZOOM',
  '04': '9/18(금) 17:00', '05': '9/18(금) 18:00', '06': '9/19(토) 15:00',
  '07': '9/19(토) 16:00', '08': '9/20(일) 18:00', '폐막식': '9/20(일) 19:00'
};

function syncRsvp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName(SRC_SHEET) || ss.getSheets()[0];
  var list = src.getDataRange().getValues().slice(1);
  var programs = [];   // {title, no, when, rows:[...]}
  var errors = [];

  list.forEach(function (r) {
    var title = String(r[0] || '').trim(), url = String(r[3] || '').trim();
    if (!url) return;
    if (!/신청|폐막식/.test(title) || /만족도|개막식/.test(title)) return;   // 크로스토크 8개 + 폐막식만
    var form;
    try { form = FormApp.openByUrl(url); } catch (e) { errors.push(title + ': ' + e); return; }
    var m = title.match(/^\[(\d\d)\]/);
    var no = m ? m[1] : (/폐막식/.test(title) ? '폐막식' : '');
    var prog = { title: title, no: no, when: WHEN[no] || '', rows: [] };
    form.getResponses().forEach(function (resp) {
      var name = '', phone = '', email = '', n = '', all = [];
      resp.getItemResponses().forEach(function (ir) {
        var q = ir.getItem().getTitle(), v = ir.getResponse();
        v = (v === null || v === undefined) ? '' : (Array.isArray(v) ? v.join(' ') : String(v));
        all.push(v);
        if (!name && /성함|성명|이름|name/i.test(q)) name = v;
        else if (!phone && /전화|연락|휴대|핸드폰|phone/i.test(q)) phone = v;
        else if (!email && /이메일|메일|e-?mail/i.test(q)) email = v;
        else if (!n && /인원|참석.*수|몇\s*명|people/i.test(q)) n = v;
      });
      /* 질문 제목으로 못 찾으면 답변 모양으로 짐작 */
      if (!phone) phone = all.filter(function (v) { return /0?1[016789]\D?\d{3,4}\D?\d{4}/.test(v); })[0] || '';
      if (!email) email = all.filter(function (v) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v.trim()); })[0] || '';
      if (!email) { try { email = resp.getRespondentEmail() || ''; } catch (e) {} }   // 폼의 '이메일 수집' 설정
      if (!name) name = all.filter(function (v) { return /^[가-힣]{2,6}$/.test(v.trim()); })[0] || '';
      var digits = phone.replace(/\D/g, '');
      if (/^1[016789]\d{7,8}$/.test(digits)) digits = '0' + digits;             // 시트에서 앞 0이 빠진 경우
      var pretty = /^01\d{8,9}$/.test(digits) ? digits.replace(/^(\d{3})(\d{3,4})(\d{4})$/, '$1-$2-$3') : phone.trim();
      prog.rows.push({
        name: name.trim(), phone: pretty, digits: digits, p4: digits.slice(-4), email: email.trim().toLowerCase(),
        n: parseInt(String(n).replace(/\D/g, ''), 10) || 1, ts: resp.getTimestamp(), id: resp.getId()
      });
    });
    /* 같은 프로그램에 같은 사람(이름+전화)이 두 번 접수했으면 나중 접수 기준 — 앱과 같은 규칙 */
    var seen = {};
    prog.rows.sort(function (a, b) { return a.ts - b.ts; }).forEach(function (x) { seen[x.name.replace(/\s/g, '') + '|' + x.p4] = x; });
    prog.rows = Object.keys(seen).map(function (k) { return seen[k]; }).sort(function (a, b) { return a.ts - b.ts; });
    programs.push(prog);
  });

  /* (1) 앱용 '접수' 탭 — 뒷자리 4자리만 */
  var out = [['프로그램', '성함', '전화 뒷자리', '인원', '접수 시각', '응답 id']];
  programs.forEach(function (p) { p.rows.forEach(function (x) { out.push([p.title, x.name, x.p4, x.n, x.ts, x.id]); }); });
  var sh = ss.getSheetByName(OUT_SHEET) || ss.insertSheet(OUT_SHEET);
  sh.clearContents();
  sh.getRange(1, 3, out.length, 1).setNumberFormat('@');      // 뒷자리가 0으로 시작해도 문자로 유지
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.getRange(1, 5, out.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'MM-dd HH:mm');
  sh.getRange('H1').setValue('동기화 ' + stamp + (errors.length ? ' · 오류 ' + errors.length + '건' : ''));
  sh.getRange('H2').setValue(errors.length ? errors.join('\n') : '');

  /* (2) 발송용 비공개 시트 */
  var book = outBook_();
  writeSummary_(book, programs, stamp);
  programs.forEach(function (p) { writeProgram_(book, p); });
  sh.getRange('H3').setValue('발송용 명단(비공개, 소유자만): ' + book.getUrl());
}

/* 발송용 시트 — 처음 실행 때 소유자 드라이브에 만들고 id를 기억한다 */
function outBook_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('OUT_BOOK'), book = null;
  if (id) { try { book = SpreadsheetApp.openById(id); } catch (e) { book = null; } }
  if (!book) { book = SpreadsheetApp.create(OUT_BOOK_NAME); props.setProperty('OUT_BOOK', book.getId()); }
  return book;
}

function sheetOf_(book, name) {
  var s = book.getSheetByName(name);
  if (!s) s = book.insertSheet(name);
  s.clearContents();
  return s;
}

/* 발송 예정일 = 프로그램 전날 */
function sendDay_(when) {
  var m = String(when || '').match(/^(\d+)\/(\d+)/); if (!m) return '';
  var d = new Date(2026, +m[1] - 1, +m[2] - 1);
  return (d.getMonth() + 1) + '/' + d.getDate() + '(' + '일월화수목금토'[d.getDay()] + ')';
}

function writeSummary_(book, programs, stamp) {
  var s = sheetOf_(book, '요약');
  var rows = [['프로그램', '일시', '발송일(전날)', '접수 팀', '인원 합계', '정원 잔여(' + CAPACITY + ')', '전화번호 목록 (복사용)', '이메일 목록 (복사용)']];
  programs.forEach(function (p) {
    var ppl = p.rows.reduce(function (a, x) { return a + x.n; }, 0);
    var phones = uniq_(p.rows.map(function (x) { return x.phone; }));
    var mails = uniq_(p.rows.map(function (x) { return x.email; }));
    rows.push([p.title, p.when, sendDay_(p.when), p.rows.length, ppl, CAPACITY - ppl, phones.join(', '), mails.join(', ')]);
  });
  rows.push(['']);
  rows.push(['동기화 ' + stamp + ' · 5분마다 자동 갱신 · 프로그램별 탭에 전체 명단']);
  s.getRange(1, 1, rows.length, 8).setValues(rows.map(function (r) { while (r.length < 8) r.push(''); return r; }));
  s.getRange(1, 1, 1, 8).setFontWeight('bold');
  s.setFrozenRows(1);
  s.setColumnWidth(1, 360); s.setColumnWidth(7, 420); s.setColumnWidth(8, 420);
  s.getRange(2, 7, Math.max(programs.length, 1), 2).setWrap(true);
  book.setActiveSheet(s); book.moveActiveSheet(1);
  var def = book.getSheetByName('Sheet1') || book.getSheetByName('시트1');
  if (def && def.getName() !== '요약' && book.getSheets().length > 1 && def.getLastRow() === 0) book.deleteSheet(def);
}

function writeProgram_(book, p) {
  var name = (p.no === '폐막식' ? '폐막식' : p.no + ' ' + p.title.replace(/^\[\d\d\]\s*/, '').replace(/\s*[—-]\s*신청$/, '')).slice(0, 40);
  var s = sheetOf_(book, name);
  var rows = [['성함', '전화번호', '이메일', '인원', '접수 시각', p.title + ' · ' + p.when]];
  p.rows.forEach(function (x) { rows.push([x.name, x.phone, x.email, x.n, x.ts, '']); });
  var ppl = p.rows.reduce(function (a, x) { return a + x.n; }, 0);
  rows.push(['합계', p.rows.length + '팀', '', ppl + '명', '', '']);
  s.getRange(1, 2, rows.length, 1).setNumberFormat('@');
  s.getRange(1, 1, rows.length, 6).setValues(rows);
  s.getRange(1, 5, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  s.getRange(1, 1, 1, 6).setFontWeight('bold');
  s.getRange(rows.length, 1, 1, 6).setFontWeight('bold');
  s.setFrozenRows(1);
  s.setColumnWidth(2, 130); s.setColumnWidth(3, 220); s.setColumnWidth(5, 130); s.setColumnWidth(6, 360);
}

function uniq_(arr) {
  var seen = {}; return arr.filter(function (v) { v = String(v || '').trim(); if (!v || seen[v]) return false; seen[v] = 1; return true; });
}

function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncRsvp') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncRsvp').timeBased().everyMinutes(5).create();
  syncRsvp();
}
