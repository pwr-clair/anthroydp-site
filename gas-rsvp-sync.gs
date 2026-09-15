/**
 * 연계프로그램 접수 동기화 — 구글 폼 응답 → '접수' 탭
 *
 * 붙이는 곳: "연계프로그램 신청자 현황" 스프레드시트 > 확장 프로그램 > Apps Script
 * 시트1(폼 목록)의 '링크' 열에 있는 폼을 하나씩 열어 응답을 읽고,
 * '접수' 탭에 [프로그램, 성함, 전화 뒷자리, 인원, 접수 시각, 응답 id] 로 다시 쓴다.
 * 전화번호는 뒷자리 4자리만 남긴다 — 운영인력 페이지(RSVP 탭)가 이 탭만 읽는다.
 *
 * 처음 1회: installTrigger 실행(권한 승인) → 5분마다 자동 동기화 + 즉시 1회 실행.
 * 수동 갱신: syncRsvp 실행.
 */
var SRC_SHEET = '시트1';   // 폼 목록 탭 (A=폼 이름, D=링크)
var OUT_SHEET = '접수';    // 결과 탭 (없으면 만든다)

function syncRsvp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = ss.getSheetByName(SRC_SHEET) || ss.getSheets()[0];
  var list = src.getDataRange().getValues().slice(1);
  var out = [['프로그램', '성함', '전화 뒷자리', '인원', '접수 시각', '응답 id']];
  var errors = [];

  list.forEach(function (r) {
    var title = String(r[0] || '').trim(), url = String(r[3] || '').trim();
    if (!url) return;
    if (!/신청|폐막식/.test(title) || /만족도|개막식/.test(title)) return;   // 크로스토크 8개 + 폐막식만
    var form;
    try { form = FormApp.openByUrl(url); } catch (e) { errors.push(title + ': ' + e); return; }
    form.getResponses().forEach(function (resp) {
      var name = '', phone = '', n = '', all = [];
      resp.getItemResponses().forEach(function (ir) {
        var q = ir.getItem().getTitle(), v = ir.getResponse();
        v = (v === null || v === undefined) ? '' : (Array.isArray(v) ? v.join(' ') : String(v));
        all.push(v);
        if (!name && /성함|성명|이름|name/i.test(q)) name = v;
        else if (!phone && /전화|연락|휴대|핸드폰|phone/i.test(q)) phone = v;
        else if (!n && /인원|참석.*수|몇\s*명|people/i.test(q)) n = v;
      });
      /* 질문 제목으로 못 찾으면 답변 모양으로 짐작 */
      if (!phone) phone = all.filter(function (v) { return /0?1[016789]\D?\d{3,4}\D?\d{4}/.test(v); })[0] || '';
      if (!name) name = all.filter(function (v) { return /^[가-힣]{2,6}$/.test(v.trim()); })[0] || '';
      var p4 = phone.replace(/\D/g, '').slice(-4);
      var cnt = parseInt(String(n).replace(/\D/g, ''), 10) || 1;
      out.push([title, name.trim(), p4, cnt, resp.getTimestamp(), resp.getId()]);
    });
  });

  var sh = ss.getSheetByName(OUT_SHEET) || ss.insertSheet(OUT_SHEET);
  sh.clearContents();
  sh.getRange(1, 3, out.length, 1).setNumberFormat('@');      // 뒷자리가 0으로 시작해도 문자로 유지
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);
  sh.getRange(1, 5, out.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('H1').setValue('동기화 ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MM-dd HH:mm') + (errors.length ? ' · 오류 ' + errors.length + '건' : ''));
  if (errors.length) sh.getRange('H2').setValue(errors.join('\n'));
}

function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncRsvp') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncRsvp').timeBased().everyMinutes(5).create();
  syncRsvp();
}
