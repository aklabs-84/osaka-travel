/**
 * 오사카 플래너 — 구글 시트 중계 프로그램
 *
 * 설치: 구글 시트 → 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기
 *       → 아래 SECRET을 길고 아무도 모르는 글자로 바꾸기 → 배포(웹 앱, 액세스: 모든 사용자)
 *
 * 하는 일: 폰이 보낸 내용을 시트에 합쳐 저장하고, 합친 결과를 돌려준다.
 * 합치는 규칙: 항목마다 'u'(마지막으로 고친 시각)가 더 큰 쪽이 이긴다.
 */
const SECRET = '여기를-길고-어려운-글자로-바꾸세요';  // 폰의 "비밀 코드"와 똑같이
const SHEET = 'data';
const CHUNK = 30000; // 시트 한 칸에 넣는 글자 수(한 칸 최대치보다 훨씬 작게)

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
}

function read_() {
  const sh = sheet_(), n = sh.getLastRow();
  if (n < 1) return {};
  const txt = sh.getRange(1, 1, n, 1).getValues().map(r => r[0]).join('');
  try { return JSON.parse(txt) || {}; } catch (e) { return {}; }
}

function write_(doc) {
  const sh = sheet_(), txt = JSON.stringify(doc), rows = [];
  for (let i = 0; i < txt.length; i += CHUNK) rows.push([txt.slice(i, i + CHUNK)]);
  sh.clear();
  if (rows.length) sh.getRange(1, 1, rows.length, 1).setValues(rows);
}

function merge_(a, b) {
  Object.keys(b || {}).forEach(cat => {
    a[cat] = a[cat] || {};
    Object.keys(b[cat] || {}).forEach(key => {
      const x = a[cat][key], y = b[cat][key];
      if (y && typeof y.u === 'number' && (!x || y.u > x.u)) a[cat][key] = y;
    });
  });
  return a;
}

// 브라우저로 주소를 열어 "연결 OK"가 보이는지 확인하는 용도
function doGet() {
  return out_({ ok: 1, msg: '오사카 플래너 중계 프로그램이 켜져 있어요' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: 0, error: 'bad' }); }
  if (!body || body.k !== SECRET || SECRET.indexOf('바꾸세요') >= 0) return out_({ ok: 0, error: 'code' });
  if (body.act === 'receipt') return out_(receipt_(body.img));   // 영수증 읽기(시트는 건드리지 않음)
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const merged = merge_(read_(), body.doc || {});
    write_(merged);
    let rerr = '';
    if (body.rep) { try { drawReport_(body.rep); } catch (err) { rerr = String(err); } }
    return out_({ ok: 1, doc: merged, rerr: rerr });
  } finally {
    lock.releaseLock();
  }
}

/* ================= 영수증 읽기 =================
 * 폰이 보낸 영수증 사진(base64 JPEG)을 Gemini에 보여 주고 합계·가게·품목·종류를 받아 돌려준다.
 * 열쇠는 코드에 쓰지 않는다: 프로젝트 설정 → 스크립트 속성에 GEMINI_KEY 로 넣는다.
 * 모델을 바꾸고 싶으면 같은 곳에 GEMINI_MODEL 을 넣는다(없으면 아래 기본값).
 */
const DEFAULT_MODEL = 'gemini-3.1-flash-lite';
const DAILY_MAX = 150;   // 하루 최대 읽기 횟수(실수로 많이 눌러도 요금·한도 보호)

function receipt_(img) {
  const P = PropertiesService.getScriptProperties();
  const key = P.getProperty('GEMINI_KEY');
  if (!key) return { ok: 0, error: 'nokey' };
  if (!img || String(img).length > 4000000) return { ok: 0, error: 'img' };
  const day = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
  const used = P.getProperty('rc_day') === day ? +P.getProperty('rc_n') || 0 : 0;
  if (used >= DAILY_MAX) return { ok: 0, error: 'limit' };
  P.setProperty('rc_day', day); P.setProperty('rc_n', String(used + 1));

  const model = P.getProperty('GEMINI_MODEL') || DEFAULT_MODEL;
  const prompt = '이것은 일본 여행 중 받은 영수증 사진이다. 아래 항목을 읽어 JSON으로만 답해라.\n' +
    '- total: 세금 포함 최종 결제 합계(엔, 정수). 합계(合計)·税込·お会計 중 최종 지불액. 못 찾으면 0\n' +
    '- store: 가게 이름(한국어로 읽기 쉽게, 15자 이내. 모르면 빈 문자열)\n' +
    '- items: 산 품목을 한국어로 짧게 요약(예: "라멘, 교자, 맥주"). 20자 이내, 많으면 대표 2~3개 + "외"\n' +
    '- cat: food(식당·카페·편의점 음식), move(교통), shop(쇼핑), do(관광·입장), etc(그 외) 중 하나\n' +
    '영수증이 아니거나 읽을 수 없으면 total을 0으로 한다.';
  const payload = {
    contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: img } }] }],
    generationConfig: {
      temperature: 0, responseMimeType: 'application/json',
      responseSchema: { type: 'OBJECT', properties: {
        total: { type: 'INTEGER' }, store: { type: 'STRING' }, items: { type: 'STRING' },
        cat: { type: 'STRING', enum: ['food', 'move', 'shop', 'do', 'etc'] } },
        required: ['total', 'store', 'items', 'cat'] }
    }
  };
  let res;
  try {
    res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': key }, payload: JSON.stringify(payload)
    });
  } catch (err) { return { ok: 0, error: 'net' }; }
  if (res.getResponseCode() !== 200) return { ok: 0, error: 'ai' + res.getResponseCode() };
  try {
    const t = JSON.parse(res.getContentText()).candidates[0].content.parts[0].text;
    const r = JSON.parse(t);
    return { ok: 1, total: Math.max(0, Math.round(+r.total || 0)), store: String(r.store || '').slice(0, 20),
             items: String(r.items || '').slice(0, 30), cat: ['food', 'move', 'shop', 'do', 'etc'].indexOf(r.cat) >= 0 ? r.cat : 'etc' };
  } catch (err) { return { ok: 0, error: 'parse' }; }
}

/* ================= 여행 보고서 탭 그리기 =================
 * 앱이 보낸 표(rep)를 보기 좋게 그린다. 내용이 바뀐 때만 다시 그린다.
 * 여기서 고친 글자는 다음에 다시 그릴 때 사라진다 → 고치는 곳은 앱.
 */
const NAMES = ['📅 일정', '📍 장소 목록', '💴 여행 경비', '✅ 체크리스트', '🏨 숙소·항공'];
const COL = { head: '#0E9AA7', title: '#0B5F68', zebra: '#F1F8F9', sum: '#FFF4CC', sec: '#D7EEF1', note: '#6B7280' };
const DAYBG = { '11/30': '#E3F2FD', '12/1': '#FFF3E0', '12/2': '#E8F5E9', '날짜 미정': '#F3F4F6' };
const KINDBG = { '쇼핑': '#FCE4EC', '맛집': '#FFF3E0', '구경': '#E8F5E9', '이동': '#E3F2FD', '숙소': '#EDE7F6' };

function drawReport_(rep) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const P = PropertiesService.getScriptProperties();
  const h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(rep)));
  const exist = NAMES.every(n => ss.getSheetByName(n));
  if (P.getProperty('rep') === h && exist) return;

  drawPlan_(fresh_(ss, NAMES[0]), rep);
  drawPlaces_(fresh_(ss, NAMES[1]), rep);
  drawExp_(fresh_(ss, NAMES[2]), rep);
  drawTodo_(fresh_(ss, NAMES[3]), rep);
  drawInfo_(fresh_(ss, NAMES[4]), rep);

  // 탭 순서: 보고서 탭이 앞, 데이터 창고(data)는 맨 뒤에 숨김
  NAMES.forEach((n, i) => { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1); });
  const data = ss.getSheetByName(SHEET);
  if (data) { ss.setActiveSheet(data); ss.moveActiveSheet(ss.getNumSheets()); }
  ss.getSheets().forEach(s => {
    if ((s.getName() === '시트1' || s.getName() === 'Sheet1') && s.getLastRow() === 0 && ss.getNumSheets() > 1) ss.deleteSheet(s);
  });
  ss.setActiveSheet(ss.getSheetByName(NAMES[0]));
  if (data && !data.isSheetHidden()) data.hideSheet();
  SpreadsheetApp.flush();
  P.setProperty('rep', h);
}

function fresh_(ss, name) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.setFrozenRows(0);
  sh.setHiddenGridlines(true);
  return sh;
}

function title_(sh, text, n) {
  sh.getRange(1, 1, 1, n).merge().setValue(text).setBackground(COL.title).setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 40);
  sh.getRange(2, 1, 1, n).merge().setValue('앱에서 고친 내용이 이 표에 자동으로 나타나요. 고치는 곳은 앱이에요 (여기서 고치면 다음 맞춤 때 사라져요).')
    .setFontColor(COL.note).setFontSize(10).setVerticalAlignment('middle');
}

// 표 그리기: 머리줄 + 줄 사이 색 + 테두리. numCols = 숫자로 넣을 칸 번호(0부터)
function table_(sh, row, headers, rows, numCols) {
  numCols = numCols || [];
  const nc = headers.length;
  const hr = sh.getRange(row, 1, 1, nc);
  hr.setValues([headers]).setBackground(COL.head).setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(row, 30);
  if (rows.length) {
    const r = sh.getRange(row + 1, 1, rows.length, nc);
    // 글자 칸은 '글자'로 고정 (= + 로 시작해도 식으로 읽히지 않게)
    for (let c = 0; c < nc; c++) if (numCols.indexOf(c) < 0) sh.getRange(row + 1, c + 1, rows.length, 1).setNumberFormat('@');
    r.setValues(rows).setVerticalAlignment('middle').setWrap(true);
    for (let i = 0; i < rows.length; i++) if (i % 2 === 1) sh.getRange(row + 1 + i, 1, 1, nc).setBackground(COL.zebra);
    r.setBorder(true, true, true, true, true, true, '#D5DDE0', SpreadsheetApp.BorderStyle.SOLID);
  }
  numCols.forEach(c => { if (rows.length) sh.getRange(row + 1, c + 1, rows.length, 1).setHorizontalAlignment('right'); });
  return row + 1 + rows.length;
}

function widths_(sh, ws) { ws.forEach((w, i) => sh.setColumnWidth(i + 1, w)); }

function drawPlan_(sh, rep) {
  title_(sh, '📅 ' + rep.title + ' — 일정', 6);
  const end = table_(sh, 4, ['날짜', '시간', '장소', '종류', '메모', '링크'], rep.plan);
  rep.plan.forEach((r, i) => { const bg = DAYBG[r[0]]; if (bg) sh.getRange(5 + i, 1, 1, 2).setBackground(bg).setFontWeight('bold'); });
  widths_(sh, [90, 70, 230, 70, 260, 200]);
  sh.setFrozenRows(4);
  return end;
}

function drawPlaces_(sh, rep) {
  title_(sh, '📍 가고 싶은 곳 · 장소 목록', 6);
  table_(sh, 4, ['장소', '종류', '날짜', '시간', '메모', '링크'], rep.places);
  rep.places.forEach((r, i) => { const bg = KINDBG[r[1]]; if (bg) sh.getRange(5 + i, 2).setBackground(bg).setFontWeight('bold').setHorizontalAlignment('center'); });
  widths_(sh, [240, 70, 90, 70, 260, 200]);
  sh.setFrozenRows(4);
}

function drawExp_(sh, rep) {
  const ex = rep.ex, cats = rep.cats;
  title_(sh, '💴 여행 경비', 5);
  // 날짜 목록(처음 나온 순서)
  const dates = []; ex.forEach(r => { if (dates.indexOf(r[0]) < 0) dates.push(r[0]); });
  // 줄 위치 정하기
  const rSum = 4;                       // 4~8 요약
  const rCat = 10;                      // 종류별 제목
  const rCatH = rCat + 1, rCatB = rCatH + 1, rCatT = rCatB + cats.length;
  const rDay = rCatT + 2, rDayH = rDay + 1, rDayB = rDayH + 1, rDayT = rDayB + dates.length;
  const rDet = rDayT + 2, rDetH = rDet + 1, rDetB = rDetH + 1;
  const END = 2000;
  const D = '$D$' + rDetB + ':$D$' + END, A = '$A$' + rDetB + ':$A$' + END, B = '$B$' + rDetB + ':$B$' + END;

  // 요약
  const sumRows = [
    ['환율 (¥100 = ?원)', rep.rate],
    ['총 지출 (엔)', '=SUM(' + D + ')'],
    ['총 지출 (원)', '=B5*B4/100'],
    ['기록 건수', '=COUNT(' + D + ')'],
    ['하루 평균 (원, 3일 기준)', '=ROUND(B6/3)']
  ];
  const sr = sh.getRange(rSum, 1, 5, 2);
  sr.setValues(sumRows).setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, '#D5DDE0', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(rSum, 1, 5, 1).setBackground(COL.sec).setFontWeight('bold');
  sh.getRange(rSum, 2, 5, 1).setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('B4').setNumberFormat('#,##0"원"');
  sh.getRange('B5').setNumberFormat('"¥"#,##0');
  sh.getRange('B6').setNumberFormat('#,##0"원"');
  sh.getRange('B7').setNumberFormat('0"건"');
  sh.getRange('B8').setNumberFormat('#,##0"원"');
  sh.getRange('B6').setBackground(COL.sum).setFontSize(13);
  sh.getRange('B5').setBackground(COL.sum).setFontSize(13);

  // 종류별 합계
  sec_(sh, rCat, '종류별 합계', 4);
  const catRows = cats.map((c, i) => {
    const r = rCatB + i;
    return [c, '=SUMIFS(' + D + ',' + B + ',A' + r + ')', '=B' + r + '*$B$4/100', '=IFERROR(B' + r + '/$B$5,0)'];
  });
  catRows.push(['합계', '=SUM(B' + rCatB + ':B' + (rCatT - 1) + ')', '=SUM(C' + rCatB + ':C' + (rCatT - 1) + ')', '=IFERROR(B' + rCatT + '/$B$5,0)']);
  fmtBlock_(sh, rCatH, ['종류', '엔(¥)', '원(₩)', '비율'], catRows);

  // 날짜별 합계
  sec_(sh, rDay, '날짜별 합계', 4);
  const dayRows = dates.map((d, i) => {
    const r = rDayB + i;
    return [d, '=SUMIFS(' + D + ',' + A + ',A' + r + ')', '=B' + r + '*$B$4/100', '=COUNTIF(' + A + ',A' + r + ')'];
  });
  dayRows.push(['합계', '=SUM(B' + rDayB + ':B' + (rDayT - 1) + ')', '=SUM(C' + rDayB + ':C' + (rDayT - 1) + ')', '=SUM(D' + rDayB + ':D' + (rDayT - 1) + ')']);
  fmtBlock_(sh, rDayH, ['날짜', '엔(¥)', '원(₩)', '건수'], dayRows);
  sh.getRange(rDayB, 4, dayRows.length, 1).setNumberFormat('0"건"');

  // 자세한 기록
  sec_(sh, rDet, '자세한 기록', 5);
  const det = ex.map((r, i) => [r[0], r[1], r[2], r[3], '=D' + (rDetB + i) + '*$B$4/100']);
  table_(sh, rDetH, ['날짜', '종류', '쓴 곳', '엔(¥)', '원(₩)'], det.map(r => [r[0], r[1], r[2], r[3], '']), [3, 4]);
  if (det.length) {
    sh.getRange(rDetB, 5, det.length, 1).setFormulas(det.map(r => [r[4]]));
    sh.getRange(rDetB, 4, det.length, 1).setNumberFormat('"¥"#,##0');
    sh.getRange(rDetB, 5, det.length, 1).setNumberFormat('#,##0"원"');
  }
  widths_(sh, [150, 110, 190, 110, 110]);
  sh.setFrozenRows(2);
}

function sec_(sh, row, text, n) {
  sh.getRange(row, 1, 1, n).merge().setValue(text).setBackground(COL.sec).setFontWeight('bold').setFontSize(12).setVerticalAlignment('middle');
  sh.setRowHeight(row, 28);
}

// 합계 칸이 있는 작은 표 (식 포함, 마지막 줄은 합계 모양)
function fmtBlock_(sh, headRow, headers, rows) {
  const nc = headers.length;
  sh.getRange(headRow, 1, 1, nc).setValues([headers]).setBackground(COL.head).setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  const body = sh.getRange(headRow + 1, 1, rows.length, nc);
  body.setValues(rows).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#D5DDE0', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(headRow + 1, 2, rows.length, 1).setNumberFormat('"¥"#,##0').setHorizontalAlignment('right');
  sh.getRange(headRow + 1, 3, rows.length, 1).setNumberFormat('#,##0"원"').setHorizontalAlignment('right');
  sh.getRange(headRow + 1, 4, rows.length, 1).setNumberFormat('0%').setHorizontalAlignment('right');
  sh.getRange(headRow + rows.length, 1, 1, nc).setBackground(COL.sum).setFontWeight('bold');
}

function drawTodo_(sh, rep) {
  title_(sh, '✅ 해야 할 일 체크리스트', 3);
  sh.getRange('A4').setValue('진행률').setFontWeight('bold').setBackground(COL.sec);
  sh.getRange('B4').setFormula('=COUNTIF(C7:C300,"✓ 완료")&" / "&COUNTA(B7:B300)&" 완료"').setFontWeight('bold').setBackground(COL.sum);
  const rows = rep.todo.map(r => [r[0], r[1], r[2] ? '✓ 완료' : '']);
  table_(sh, 6, ['구분', '할 일', '상태'], rows);
  rows.forEach((r, i) => { if (r[2]) sh.getRange(7 + i, 1, 1, 3).setBackground('#E6F4EA'); });
  sh.getRange(7, 3, Math.max(rows.length, 1), 1).setHorizontalAlignment('center').setFontWeight('bold');
  widths_(sh, [120, 420, 90]);
  sh.setFrozenRows(6);
}

function drawInfo_(sh, rep) {
  title_(sh, '🏨 숙소 · ✈️ 항공 정보', 2);
  const rows = rep.info.map(r => [r[0], r[1] || '(아직 안 적었어요)']);
  table_(sh, 4, ['항목', '내용'], rows);
  sh.getRange(5, 1, Math.max(rows.length, 1), 1).setBackground(COL.sec).setFontWeight('bold');
  widths_(sh, [170, 420]);
}
