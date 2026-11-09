/**
 * 소요제기 판정 웹앱 — 계약(LTA/3PP) 품목은 구매내역서를 만들고, Spot 품목은 견적 대상으로 분류한다.
 *
 * 스크립트 속성(선택):
 *   SHEET_ID         자재정보가 있는 스프레드시트 ID (없으면 이 스크립트가 붙은 시트 사용)
 *   TEMPLATE_ID      (필수) 구매내역서 템플릿 구글 시트 ID — templates/구매내역서_템플릿.xlsx 를 드라이브에 올려 구글 시트로 변환
 *   OUTPUT_FOLDER_ID 생성한 구매내역서를 넣을 드라이브 폴더 ID (권장)
 */
var SHEET_MASTER = '자재정보';
var SHEET_RATE = '환율';

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
    o.program = '';   // Program은 생성된 시트에서 담당자가 입력
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
  var name = ymd + ' 구매내역서_' + suppliers.join('_');

  var props = PropertiesService.getScriptProperties();
  var templateId = props.getProperty('TEMPLATE_ID');
  if (!templateId) throw new Error('스크립트 속성 TEMPLATE_ID(구매내역서 템플릿 시트 ID)를 설정해 주세요. README 참고.');
  var folderId = props.getProperty('OUTPUT_FOLDER_ID');
  var file = folderId
    ? DriveApp.getFileById(templateId).makeCopy(name, DriveApp.getFolderById(folderId))
    : DriveApp.getFileById(templateId).makeCopy(name);
  var ss = SpreadsheetApp.openById(file.getId());

  applyPlan_(ss, buildPurchasePlan_(lines, curs, rates, base));
  var proto = ss.getSheetByName(TPL.PROTO);
  if (proto) ss.deleteSheet(proto);
  ss.setActiveSheet(ss.getSheetByName(TPL.SUMMARY));
  return { url: ss.getUrl(), name: name };
}

/** buildPurchasePlan_ 의 작업 목록을 시트에 적용 */
function applyPlan_(ss, ops) {
  var sheets = {};
  var sh = function (n) {
    if (!sheets[n]) {
      sheets[n] = ss.getSheetByName(n);
      if (!sheets[n]) throw new Error('템플릿에 "' + n + '" 탭이 없습니다.');
    }
    return sheets[n];
  };
  var FORMAT = SpreadsheetApp.CopyPasteType.PASTE_FORMAT;
  ops.forEach(function (o) {
    switch (o.op) {
      case 'fmt':
        sh(TPL.PROTO).getRange(o.from, o.c1, 1, o.c2 - o.c1 + 1).copyTo(sh(o.s).getRange(o.r, o.c1), FORMAT, false);
        break;
      case 'cellFmt':
        sh(o.fs).getRange(o.fr, o.fc).copyTo(sh(o.s).getRange(o.r, o.c), FORMAT, false);
        break;
      case 'set': {
        var cell = sh(o.s).getRange(o.r, o.c);
        if (o.v && typeof o.v === 'object' && o.v.date) cell.setValue(toDate_(o.v.date));
        else if (typeof o.v === 'string' && o.v.charAt(0) === '=') cell.setFormula(o.v);
        else cell.setValue(o.v);
        break;
      }
      case 'nf': sh(o.s).getRange(o.r, o.c).setNumberFormat(o.f); break;
      case 'bg': sh(o.s).getRange(o.r, o.c).setBackground(o.color); break;
      case 'merge': sh(o.s).getRange(o.r1, o.c1, o.r2 - o.r1 + 1, o.c2 - o.c1 + 1).merge(); break;
      case 'unmerge': sh(o.s).getRange(o.r1, o.c1, o.r2 - o.r1 + 1, o.c2 - o.c1 + 1).breakApart(); break;
      case 'height': sh(o.s).setRowHeight(o.r, Math.round(o.pt * 4 / 3)); break;
      case 'bottom':
        sh(o.s).getRange(o.r, o.c1, 1, o.c2 - o.c1 + 1)
          .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
        break;
      case 'list':
        sh(o.s).getRange(o.r1, o.c, o.r2 - o.r1 + 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(o.items, true).setAllowInvalid(true).build());
        break;
      case 'freeze': sh(o.s).setFrozenRows(o.rows); sh(o.s).setFrozenColumns(o.cols); break;
      default: throw new Error('알 수 없는 작업: ' + o.op);
    }
  });
}
