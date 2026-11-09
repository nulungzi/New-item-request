/**
 * 소요제기 판정 웹앱 — 계약(LTA/3PP) 품목은 구매내역서를 만들고, Spot 품목은 견적 대상으로 분류한다.
 *
 * 스크립트 속성(선택):
 *   SHEET_ID         자재정보가 있는 스프레드시트 ID (없으면 이 스크립트가 붙은 시트 사용)
 *   OUTPUT_FOLDER_ID 생성한 구매내역서를 넣을 드라이브 폴더 ID (없으면 내 드라이브 루트)
 */
var SHEET_MASTER = '자재정보';
var SHEET_RATE = '환율';
var DETAIL_TAB = '구매세부내역';
var SUMMARY_TAB = '구매요약';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('소요제기 판정')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function db_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function loadContracts_() {
  var sh = db_().getSheetByName(SHEET_MASTER);
  if (!sh) throw new Error('"' + SHEET_MASTER + '" 시트가 없습니다.');
  return parseContracts_(sh.getDataRange().getValues());
}

/** 환율 시트: A열 통화, B열 환율(KRW). 없으면 만든다. */
function rateSheet_() {
  var ss = db_();
  var sh = ss.getSheetByName(SHEET_RATE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_RATE);
    sh.getRange(1, 1, 1, 3).setValues([['CUR', 'KRW 환율', '수정일']]);
  }
  return sh;
}

function loadRates_() {
  var rates = {};
  rateSheet_().getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[0]) rates[String(r[0]).trim().toUpperCase()] = r[1];
  });
  return rates;
}

function saveRates_(rates) {
  var sh = rateSheet_();
  var rows = sh.getDataRange().getValues();
  var now = new Date();
  Object.keys(rates).forEach(function (cur) {
    var i = rows.findIndex(function (r, n) { return n > 0 && String(r[0]).toUpperCase() === cur; });
    if (i < 0) { rows.push([cur, rates[cur], now]); sh.appendRow([cur, rates[cur], now]); }
    else sh.getRange(i + 1, 2, 1, 2).setValues([[rates[cur], now]]);
  });
}

/** 웹앱이 처음 열릴 때 */
function apiInit() {
  return { today: fmtDate_(new Date()), rates: loadRates_() };
}

/** 판정. Date는 문자열로 바꿔서 돌려준다 (google.script.run 직렬화 안전). */
/** baseDateStr = 작성일 (파일명·환율 기준일·리드타임 계산). 계약 판정은 품목별 Need-by-date 기준. */
function apiClassify(text, baseDateStr) {
  var parsed = parseRequest_(text);
  var base = toDate_(baseDateStr) || toDate_(new Date());
  var result = classify_(parsed.requests, loadContracts_(), base).map(function (r) {
    r.needBy = fmtDate_(r.needBy);
    r.candidates = r.candidates.map(function (c) {
      var o = {};
      Object.keys(c).forEach(function (k) { o[k] = c[k] instanceof Date ? fmtDate_(c[k]) : c[k]; });
      return o;
    });
    return r;
  });
  return { results: result, skipped: parsed.skipped, rates: loadRates_() };
}

/**
 * 구매내역서 생성.
 * payload = { baseDate, rates: {USD: 1339.3, ...}, remarks: {업체명: '비고'}, lines: [{item, qty, needBy, program, contractRow}] }
 * contractRow 로 자재정보 행을 다시 읽어 단가를 서버에서 확정한다(화면 값 신뢰 안 함).
 */
function apiCreatePurchase(payload) {
  var byRow = {};
  loadContracts_().forEach(function (c) { byRow[c.row] = c; });
  var lines = payload.lines.map(function (l) {
    var c = byRow[l.contractRow];
    if (!c || c.item !== l.item) throw new Error(l.item + ': 자재정보가 바뀌었습니다. 다시 판정해 주세요.');
    var o = {};
    Object.keys(c).forEach(function (k) { o[k] = c[k]; });
    o.qty = Number(l.qty);
    o.program = l.program || c.program || '';   // 소요제기에 적으면 우선, 없으면 자재정보 값
    o.needBy = toDate_(l.needBy);
    o.remark = (payload.remarks || {})[c.supplier] || '';
    return o;
  });
  if (!lines.length) throw new Error('구매내역서에 넣을 계약 품목이 없습니다.');

  var rates = {};
  Object.keys(payload.rates || {}).forEach(function (k) {
    var n = Number(payload.rates[k]);
    if (!(n > 0)) throw new Error(k + ' 환율을 입력해 주세요.');
    rates[k] = n;
  });
  var curs = [];
  lines.forEach(function (l) { if (l.cur !== 'KRW' && curs.indexOf(l.cur) < 0) curs.push(l.cur); });
  curs.forEach(function (c) { if (!rates[c]) throw new Error(c + ' 환율을 입력해 주세요.'); });
  saveRates_(rates);

  var base = toDate_(payload.baseDate) || toDate_(new Date());
  var suppliers = [];
  lines.forEach(function (l) { if (suppliers.indexOf(l.supplier) < 0) suppliers.push(l.supplier); });
  var ymd = fmtDate_(base).replace(/-/g, '').slice(2);
  var ss = SpreadsheetApp.create(ymd + ' 구매내역서_' + suppliers.join('_'));

  var detail = ss.getSheets()[0].setName(DETAIL_TAB);
  var info = writeDetail_(detail, lines, curs, rates, base);
  writeSummary_(ss.insertSheet(SUMMARY_TAB, 0), lines, curs, info);

  var folderId = PropertiesService.getScriptProperties().getProperty('OUTPUT_FOLDER_ID');
  if (folderId) DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(folderId));
  return { url: ss.getUrl(), name: ss.getName() };
}

var DETAIL_HEADERS = ['L/I', 'Program', 'Item No.', 'Material', 'Description', 'Product\nName', 'Qty', 'CUR',
  'Unit Price', 'Amount', 'Ship From', 'Leadtime', 'UoM', 'Shelf Life', 'Supplier', 'Tax\nCode', 'Price\nMode', '비고', 'Need-by\nDate'];

/** 구매세부내역 탭. 열 배치는 예시 양식과 같다 (B=L/I … S=비고). */
function writeDetail_(sh, lines, curs, rates, base) {
  // 기준일과 통화별 환율: O1 기준일, P1.. 통화, P2.. 환율
  sh.getRange('N1').setValue('작성일');
  sh.getRange('O1').setValue(base).setNumberFormat('yyyy-mm-dd');
  sh.getRange('N2').setValue('환율(KRW)');
  curs.forEach(function (c, i) {
    sh.getRange(1, 16 + i).setValue(c);
    sh.getRange(2, 16 + i).setValue(rates[c]).setNumberFormat('#,##0.00');
  });
  var rateCell = {};
  curs.forEach(function (c, i) { rateCell[c] = sh.getRange(2, 16 + i).getA1Notation().replace(/(\D+)(\d+)/, '$$$1$$$2'); });

  var header = 6 + curs.length;            // 통화 1개면 7행 (예시와 동일)
  var first = header + 1, last = header + lines.length;
  var I = '$I$' + first + ':$I$' + last, K = '$K$' + first + ':$K$' + last;

  // 합계: TTL(통화별), TTL(KRW), Total(KRW 환산)
  var r = 3, totalParts = [];
  curs.forEach(function (c) {
    sh.getRange(r, 10, 1, 2).setValues([['TTL(' + c + ')', '=SUMIF(' + I + ',"' + c + '",' + K + ')']]);
    totalParts.push('K' + r + '*' + rateCell[c]);
    r++;
  });
  sh.getRange(r, 10, 1, 2).setValues([['TTL(KRW)', '=SUMIF(' + I + ',"KRW",' + K + ')']]);
  totalParts.push('K' + r);
  r++;
  sh.getRange(r, 10, 1, 2).setValues([['Total', '=' + totalParts.join('+')]]);
  var note = '※ 적용환율: ' + curs.map(function (c) {
    return 'KRW ' + Number(rates[c]).toLocaleString('en-US', { minimumFractionDigits: 2 }) + '/' + c;
  }).join(', ') + ' (' + fmtDate_(base).replace(/-/g, '/') + ' 기준)';
  sh.getRange(r, 12).setValue(note);
  sh.getRange(3, 11, r - 2, 1).setNumberFormat('#,##0.00');
  var totalRow = r;

  sh.getRange(header, 2, 1, DETAIL_HEADERS.length).setValues([DETAIL_HEADERS])
    .setFontWeight('bold').setBackground('#d9e1f2').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');

  var rows = lines.map(function (l, i) {
    var n = first + i;
    return [i + 1, l.program, l.item, l.material, l.desc, l.product, l.qty, l.cur, l.price, '=J' + n + '*H' + n,
      l.shipFrom, l.leadtime, l.uom, l.shelfLife, l.supplier, l.taxCode, l.priceMode, l.remark, l.needBy || ''];
  });
  sh.getRange(first, 2, rows.length, rows[0].length).setValues(rows);
  sh.getRange(first, 3, rows.length, 2).setNumberFormat('@');      // Program, Item No.는 텍스트
  sh.getRange(first, 8, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(first, 20, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(first, 10, rows.length, 2).setNumberFormat('#,##0.00');
  sh.getRange(header, 2, rows.length + 1, DETAIL_HEADERS.length).setBorder(true, true, true, true, true, true);
  sh.setFrozenRows(header);
  sh.autoResizeColumns(2, DETAIL_HEADERS.length);
  sh.setColumnWidth(6, 320);
  sh.setColumnWidth(19, 280);
  return { first: first, last: last, rateCell: rateCell, totalRow: totalRow, note: note };
}

/** 구매요약 탭: 업체·통화별 한 줄 + 통화별 합계 + KRW 총합 */
function writeSummary_(sh, lines, curs, info) {
  var D = "'" + DETAIL_TAB + "'!";
  var P = D + '$P$' + info.first + ':$P$' + info.last;
  var I = D + '$I$' + info.first + ':$I$' + info.last;
  var H = D + '$H$' + info.first + ':$H$' + info.last;
  var K = D + '$K$' + info.first + ':$K$' + info.last;

  sh.getRange('B1').setValue('▣ 구매 요약').setFontWeight('bold').setFontSize(12);
  var head = ['L/I', '구매업체', '대표품목', '품목수', '구매량', 'CUR', '구매금액', '업체선정사유'];
  sh.getRange(3, 2, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#d9e1f2').setHorizontalAlignment('center');

  var groups = summarize_(lines);
  var rows = groups.map(function (g, i) {
    var crit = P + ',"' + g.supplier + '",' + I + ',"' + g.cur + '"';
    return [i + 1, g.supplier, g.mainMaterial, '=COUNTIFS(' + crit + ')', '=SUMIFS(' + H + ',' + crit + ')', g.cur,
      '=SUMIFS(' + K + ',' + crit + ')', g.priceMode];
  });
  sh.getRange(4, 2, rows.length, head.length).setValues(rows);
  var gEnd = 3 + rows.length;

  // CUR 별 합계
  var r = gEnd + 1, krwParts = [];
  var allCurs = curs.slice();
  if (lines.some(function (l) { return l.cur === 'KRW'; })) allCurs.push('KRW');
  allCurs.forEach(function (c) {
    var rng = function (col) { return col + '4:' + col + gEnd; };
    sh.getRange(r, 2).setValue('CUR 별 합계');
    sh.getRange(r, 5, 1, 4).setValues([['=SUMIF(' + rng('G') + ',"' + c + '",' + rng('E') + ')',
      '=SUMIF(' + rng('G') + ',"' + c + '",' + rng('F') + ')', c,
      '=SUMIF(' + rng('G') + ',"' + c + '",' + rng('H') + ')']]);
    krwParts.push(c === 'KRW' ? 'H' + r : 'H' + r + '*' + D + info.rateCell[c]);
    r++;
  });
  sh.getRange(r, 2).setValue('총 합계 (KRW 표기)').setFontWeight('bold');
  sh.getRange(r, 8).setFormula('=' + krwParts.join('+')).setFontWeight('bold');
  sh.getRange(r + 1, 6).setValue(info.note);

  sh.getRange(4, 6, r - 3, 1).setNumberFormat('#,##0');
  sh.getRange(4, 8, r - 3, 1).setNumberFormat('#,##0.00');
  sh.getRange(r, 8).setNumberFormat('#,##0');
  sh.getRange(3, 2, r - 2, head.length).setBorder(true, true, true, true, true, true);
  sh.autoResizeColumns(2, head.length);
}
