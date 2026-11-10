// ===== 소요제기 판정 웹앱 · Code.gs =====
// 시트·드라이브 입출력, 웹앱 함수(api*), 설정 점검(checkSetup). Apps Script 편집기의 Code.gs 에 이 파일 전체를 붙여넣습니다.
// Logic.gs 와 함께 있어야 동작합니다. Logic.gs 의 값(TPL 등)은 파일 맨 위가 아니라 함수 안에서만 씁니다 (파일 읽는 순서 때문).

/**
 * 소요제기 판정 웹앱 — 계약(LTA/3PP) 품목은 구매내역서를 만들고, Spot 품목은 견적 대상으로 분류한다.
 * 판정·계산은 1부(순수 함수), 여기 2부는 시트·드라이브 입출력만 맡는다.
 *
 * 스크립트 속성
 *   TEMPLATE_ID      (필수) 구매내역서 템플릿 구글 시트 ID — templates/구매내역서_템플릿.xlsx 를 드라이브에 올려 구글 시트로 저장
 *   OUTPUT_FOLDER_ID (권장) 생성한 구매내역서를 넣을 드라이브 폴더 ID
 *   SHEET_ID         (선택) 자재정보가 있는 스프레드시트 ID (없으면 이 스크립트가 붙은 시트)
 */
var SHEET_MASTER = '자재정보';
var SHEET_RATE = '환율';

/* ================================================================
 * 설정·시트 접근
 * ================================================================ */
function prop_(key) { return PropertiesService.getScriptProperties().getProperty(key); }

function db_() {
  var id = prop_('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function masterSheet_() {
  var sh = db_().getSheetByName(SHEET_MASTER);
  if (!sh) throw new Error('"' + SHEET_MASTER + '" 시트가 없습니다.');
  return sh;
}

function masterValues_() { return masterSheet_().getDataRange().getValues(); }

function loadContracts_() { return parseContracts_(masterValues_()); }

/** 자재정보 머리글 위치 {row(1부터), headers} */
function masterHeader_(values) {
  for (var r = 0; r < Math.min(values.length, 10); r++) {
    if (values[r].some(function (h) { return keyForHeader_(h) === 'item'; })) return { row: r + 1, headers: values[r] };
  }
  throw new Error('자재정보 시트에서 "Item" 머리글을 찾지 못했습니다.');
}

/** 신규 등록 양식에 쓸 머리글 목록 (시트 열 순서) */
function masterFields_(values) {
  return masterHeader_(values).headers
    .map(function (h) { return { header: String(h).replace(/\s+/g, ' ').trim(), key: keyForHeader_(h) }; })
    .filter(function (f) { return f.key; });
}

/** 환율 시트: A열 통화, B열 환율(KRW), C열 수정일. 없으면 만든다. */
function rateSheet_() {
  var ss = db_(), sh = ss.getSheetByName(SHEET_RATE);
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
  var sh = rateSheet_(), rows = sh.getDataRange().getValues(), now = new Date();
  Object.keys(rates).forEach(function (cur) {
    var i = -1;
    rows.forEach(function (r, n) { if (i < 0 && n > 0 && String(r[0]).toUpperCase() === cur) i = n; });
    if (i < 0) { rows.push([cur, rates[cur], now]); sh.appendRow([cur, rates[cur], now]); }
    else sh.getRange(i + 1, 2, 1, 2).setValues([[rates[cur], now]]);
  });
}

/** google.script.run 으로 보낼 수 있게 Date → 'yyyy-mm-dd' */
function plainDates_(o) {
  var out = {};
  Object.keys(o).forEach(function (k) { out[k] = Object.prototype.toString.call(o[k]) === '[object Date]' ? fmtDate_(o[k]) : o[k]; });
  return out;
}

/* ================================================================
 * 웹앱
 * ================================================================ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('소요제기 판정')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 화면이 처음 열릴 때: 작성일, 저장된 환율, 구매일정 설정, Item 안내, 신규 등록 양식 */
function apiInit() {
  var values = masterValues_();
  return {
    today: fmtDate_(new Date()),
    rates: loadRates_(),
    schedule: { tasks: SCHEDULE_TASKS, example: SCHEDULE_EXAMPLE, slots: SCHEDULE_SLOTS },
    items: itemIndex_(parseContracts_(values)),   // [Item, Description, Ship From, 계약 구분들, Shelf Life]
    fields: masterFields_(values)
  };
}

/** 소요제기 판정. input: 화면 표(행 배열) 또는 붙여넣기 글. baseDateStr: 작성일 (계약 판정은 품목별 Need-by 기준) */
function apiClassify(input, baseDateStr) {
  var parsed = Array.isArray(input) ? parseRows_(input) : parseRequest_(input);
  var base = toDate_(baseDateStr) || toDate_(new Date());
  var results = classify_(parsed.requests, loadContracts_(), base).map(function (r) {
    r.needBy = fmtDate_(r.needBy);
    r.candidates = r.candidates.map(plainDates_);
    return r;
  });
  return { results: results, skipped: parsed.skipped, rates: loadRates_() };
}

/** 신규 자재를 자재정보 시트 맨 아래에 추가. rec 키는 masterFields_ 의 key. 머리글 없는 맨 앞 번호 열은 윗줄 + 1 */
function apiRegisterItem(rec) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = masterSheet_(), values = sh.getDataRange().getValues(), head = masterHeader_(values);
    var v = validateNewItem_(rec, parseContracts_(values));
    if (v.errors.length) throw new Error(v.errors.join('\n'));

    var row = newItemRow_(head.headers, v.rec), n = row.length, last = sh.getLastRow(), at = last + 1;
    var prev = values[last - 1] || [];
    if (!String(head.headers[0]).trim() && last > head.row && typeof prev[0] === 'number') row[0] = prev[0] + 1;

    var target = sh.getRange(at, 1, 1, n);
    if (last > head.row) sh.getRange(last, 1, 1, n).copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    head.headers.forEach(function (h, i) {
      var k = keyForHeader_(h);
      if (k === 'item') sh.getRange(at, i + 1).setNumberFormat('@');
      if (k === 'start' || k === 'end') sh.getRange(at, i + 1).setNumberFormat('yyyy-mm-dd');
    });
    target.setValues([row]);
    return { row: at, item: v.rec.item, desc: v.rec.desc || '', ship: v.rec.shipFrom || '', shelf: v.rec.shelfLife || '', mode: v.rec.priceMode };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 구매내역서 생성.
 * payload = { baseDate, rates: {USD: 1339.3}, remarks: {업체명: '비고'},
 *             lines: [{item, qty, needBy, program, shelfLife, shipFrom, taxCode, contractRow}],
 *             schedule: [{start, len}] × 5 (구매일정, 반 달 단위),
 *             name: 구매내역서 제목 (비우면 자동), confirmer: 사업부 담당자 '홍길동 과장',
 *             folder: 저장할 드라이브 폴더 주소/ID (비우면 OUTPUT_FOLDER_ID 또는 기본 위치) }
 * 단가는 contractRow 로 자재정보를 다시 읽어 서버에서 확정한다(화면 값은 믿지 않음).
 */
function apiCreatePurchase(payload) {
  var lines = purchaseLines_(payload);
  var curs = uniq_(lines.map(function (l) { return l.cur; })).filter(function (c) { return c !== 'KRW'; });
  var rates = checkedRates_(payload.rates, curs);
  saveRates_(rates);

  var base = toDate_(payload.baseDate) || toDate_(new Date());
  var name = purchaseFileName_(payload.name, base, uniq_(lines.map(function (l) { return l.supplier; })));
  var ss = copyTemplate_(name, payload.folder);
  applyPlan_(ss, buildPurchasePlan_(lines, curs, rates, base, payload.schedule, { title: name, confirmer: payload.confirmer }));
  var proto = ss.getSheetByName(TPL.PROTO);
  if (proto) ss.deleteSheet(proto);
  ss.setActiveSheet(ss.getSheetByName(TPL.SUMMARY));
  return { url: ss.getUrl(), name: name, id: ss.getId() };
}

/** 화면에서 받은 줄 + 자재정보 계약 행 → 구매내역서 한 줄 */
function purchaseLines_(payload) {
  var byRow = {};
  loadContracts_().forEach(function (c) { byRow[c.row] = c; });
  var lines = (payload.lines || []).map(function (l) {
    var c = byRow[l.contractRow];
    if (!c || c.item !== l.item) throw new Error(l.item + ': 자재정보가 바뀌었습니다. 다시 판정해 주세요.');
    var o = {};
    Object.keys(c).forEach(function (k) { o[k] = c[k]; });
    o.qty = Number(l.qty);
    o.program = l.program || '';                       // 비우면 시트에서 노란 칸
    if (l.shipFrom) o.shipFrom = l.shipFrom;           // 화면에 적으면 자재정보 값 대신
    if (l.shelfLife) o.shelfLife = l.shelfLife;
    if (l.taxCode !== undefined && l.taxCode !== '') o.taxCode = isNaN(+l.taxCode) ? l.taxCode : +l.taxCode;
    o.needBy = toDate_(l.needBy);
    o.remark = (payload.remarks || {})[c.supplier] || '';
    return o;
  });
  if (!lines.length) throw new Error('구매내역서에 넣을 계약 품목이 없습니다.');
  return lines;
}

/** 필요한 통화의 환율이 모두 양수인지 확인 */
function checkedRates_(input, curs) {
  var rates = {};
  Object.keys(input || {}).forEach(function (k) {
    var n = Number(input[k]);
    if (!(n > 0)) throw new Error(k + ' 환율을 입력해 주세요.');
    rates[k] = n;
  });
  curs.forEach(function (c) { if (!rates[c]) throw new Error(c + ' 환율을 입력해 주세요.'); });
  return rates;
}

/** 템플릿 시트를 복사해 새 스프레드시트를 연다. folderInput: 담당자가 고른 폴더 주소/ID (없으면 OUTPUT_FOLDER_ID) */
function copyTemplate_(name, folderInput) {
  var templateId = prop_('TEMPLATE_ID');
  if (!templateId) throw new Error('스크립트 속성 TEMPLATE_ID(구매내역서 템플릿 시트 ID)를 설정해 주세요. README 참고.');
  var folder = targetFolder_(folderInput), tpl = DriveApp.getFileById(templateId);
  var file = folder ? tpl.makeCopy(name, folder) : tpl.makeCopy(name);
  return SpreadsheetApp.openById(file.getId());
}

/** 저장할 드라이브 폴더: 화면에서 고른 폴더 → 스크립트 속성 OUTPUT_FOLDER_ID → null(기본 위치) */
function targetFolder_(folderInput) {
  var picked = parseFolderId_(folderInput);
  if (String(folderInput || '').trim() && !picked) throw new Error('드라이브 폴더 주소를 확인해 주세요. 폴더를 열었을 때 주소창의 주소를 그대로 붙여넣으면 됩니다.');
  var id = picked || prop_('OUTPUT_FOLDER_ID');
  if (!id) return null;
  try { var f = DriveApp.getFolderById(id); f.getName(); return f; }
  catch (e) { throw new Error('드라이브 폴더를 열 수 없습니다. 주소가 맞는지, 이 폴더에 파일을 만들 권한(편집자)이 있는지 확인해 주세요.'); }
}

/** 화면 "확인" 버튼: 폴더 이름을 돌려준다 */
function apiCheckFolder(folderInput) {
  var f = targetFolder_(folderInput);
  return f ? { name: f.getName(), url: f.getUrl(), picked: !!parseFolderId_(folderInput) } : { name: '내 드라이브 (기본 위치)', url: '', picked: false };
}

/* ================================================================
 * 입력용 시트: 엑셀처럼 적어 두고 웹앱 표에 복사·붙여넣기
 * ================================================================ */
function apiCreateInputSheet(folderInput, title) {
  var folder = targetFolder_(folderInput);
  var name = String(title || '').trim() || '소요제기 입력 시트';
  var ss = SpreadsheetApp.create(name);
  var cols = INPUT_SHEET.COLUMNS, n = cols.length, ROWS = 300;

  // 1) 입력 탭
  var sh = ss.getSheets()[0].setName(INPUT_SHEET.TAB);
  sh.getRange(1, 1, 1, n).setValues([cols.map(function (c) { return c[0]; })])
    .setFontWeight('bold').setBackground('#c9daf8').setHorizontalAlignment('center').setVerticalAlignment('middle');
  cols.forEach(function (c, i) {
    sh.setColumnWidth(i + 1, c[5]);
    sh.getRange(2, i + 1, ROWS, 1).setNumberFormat(c[4]);
    sh.getRange(1, i + 1).setNote(c[1] + ' · ' + c[2] + '\n예: ' + c[3]);
    if (c[1] === '필수') sh.getRange(1, i + 1).setFontColor('#1c3d8f');
  });
  sh.getRange(2, 6, ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Air', 'Ocean'], true).setAllowInvalid(true).build());
  sh.getRange(2, 3, ROWS, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).setHelpText('날짜를 넣어 주세요. 예: 2027-06-10').build());
  sh.getRange(1, 1, ROWS + 1, n).setBorder(true, true, true, true, true, true, '#d0d7e2', SpreadsheetApp.BorderStyle.SOLID);
  sh.setRowHeight(1, 28);
  sh.setFrozenRows(1);
  if (sh.getMaxColumns() > n) sh.deleteColumns(n + 1, sh.getMaxColumns() - n);

  // 2) 사용 방법 탭
  var g = ss.insertSheet(INPUT_SHEET.GUIDE);
  var guide = [
    ['소요제기 입력 시트 사용 방법', '', '', ''],
    ['', '', '', ''],
    ['1. 적기', '"' + INPUT_SHEET.TAB + '" 탭 2행부터 한 줄에 한 품목씩 적습니다. 열 순서를 바꾸지 마세요.', '', ''],
    ['2. 복사', '2행부터 마지막 품목까지 A~G열을 드래그해서 복사(Ctrl+C)합니다. 머리글(1행)을 같이 복사해도 괜찮습니다.', '', ''],
    ['3. 붙여넣기', '웹앱 "소요제기 입력" 표의 1번 Item 칸을 누르고 붙여넣기(Ctrl+V) 합니다. 줄이 모자라면 자동으로 늘어납니다.', '', ''],
    ['4. 다음에', '2행 아래 내용을 지우고 다시 씁니다. 이 시트는 계속 써도 됩니다.', '', ''],
    ['', '', '', ''],
    ['열', '필수', '설명', '예시']
  ].concat(cols.map(function (c) { return [c[0], c[1], c[2], c[3]]; }));
  g.getRange(1, 1, guide.length, 4).setValues(guide);
  g.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  g.getRange(3, 1, 4, 1).setFontWeight('bold').setFontColor('#1c3d8f');
  var head = 8;
  g.getRange(head, 1, 1, 4).setFontWeight('bold').setBackground('#c9daf8');
  g.getRange(head, 1, cols.length + 1, 4).setBorder(true, true, true, true, true, true, '#d0d7e2', SpreadsheetApp.BorderStyle.SOLID);
  g.getRange(head + 1, 4, cols.length, 1).setNumberFormat('@').setFontColor('#5a6680');
  g.setColumnWidth(1, 110); g.setColumnWidth(2, 60); g.setColumnWidth(3, 420); g.setColumnWidth(4, 140);
  g.getRange(3, 2, 4, 3).mergeAcross();
  g.getRange(1, 1, guide.length, 4).setVerticalAlignment('middle').setWrap(true);
  ss.setActiveSheet(sh);

  if (folder) DriveApp.getFileById(ss.getId()).moveTo(folder);
  return { url: ss.getUrl(), name: ss.getName(), folder: folder ? folder.getName() : '내 드라이브' };
}

/* ================================================================
 * 시트 쓰기: buildPurchasePlan_ 의 작업 목록 적용
 * ================================================================ */
function applyPlan_(ss, ops) {
  var cache = {};
  var sh = function (n) {
    if (!cache[n]) {
      cache[n] = ss.getSheetByName(n);
      if (!cache[n]) throw new Error('템플릿에 "' + n + '" 탭이 없습니다.');
    }
    return cache[n];
  };
  var FORMAT = SpreadsheetApp.CopyPasteType.PASTE_FORMAT;
  var area = function (o) { return sh(o.s).getRange(o.r1, o.c1, o.r2 - o.r1 + 1, o.c2 - o.c1 + 1); };
  ops.forEach(function (o) {
    switch (o.op) {
      case 'fmt': sh(TPL.PROTO).getRange(o.from, o.c1, 1, o.c2 - o.c1 + 1).copyTo(sh(o.s).getRange(o.r, o.c1), FORMAT, false); break;
      case 'cellFmt': sh(o.fs).getRange(o.fr, o.fc).copyTo(sh(o.s).getRange(o.r, o.c), FORMAT, false); break;
      case 'set': setCell_(sh(o.s).getRange(o.r, o.c), o.v); break;
      case 'nf': sh(o.s).getRange(o.r, o.c).setNumberFormat(o.f); break;
      case 'bg': sh(o.s).getRange(o.r, o.c).setBackground(o.color); break;
      case 'merge': area(o).merge(); break;
      case 'unmerge': area(o).breakApart(); break;
      case 'height': sh(o.s).setRowHeight(o.r, Math.round(o.pt * 4 / 3)); break;   // pt → px
      case 'bottom':
        sh(o.s).getRange(o.r, o.c1, 1, o.c2 - o.c1 + 1)
          .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
        break;
      case 'list':
        sh(o.s).getRange(o.r1, o.c, o.r2 - o.r1 + 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(o.items, true).setAllowInvalid(true).build());
        break;
      case 'freeze': sh(o.s).setFrozenRows(o.rows); sh(o.s).setFrozenColumns(o.cols); break;
      case 'width': sh(o.s).setColumnWidth(o.c, o.px); break;
      default: throw new Error('알 수 없는 작업: ' + o.op);
    }
  });
}

/** 값 넣기: {date:'yyyy-mm-dd'} 는 날짜, '='로 시작하면 수식 */
function setCell_(cell, v) {
  if (v && typeof v === 'object' && v.date) cell.setValue(toDate_(v.date));
  else if (typeof v === 'string' && v.charAt(0) === '=') cell.setFormula(v);
  else cell.setValue(v);
}

/* ================================================================
 * 미리보기: 만든 구매내역서 탭을 HTML로
 * 1순위 구글 시트 "웹페이지(zip)" 내보내기, 실패하면 값·배경·병합만 그린 간단한 표
 * ================================================================ */
function apiPreview(id) {
  SpreadsheetApp.flush();
  var exported = {};
  try { exported = exportHtml_(id); } catch (e) { console.warn('zip 내보내기 실패: ' + e); }
  var ss = SpreadsheetApp.openById(id);
  return [TPL.SUMMARY, TPL.DETAIL, TPL.SCHEDULE, TPL.CONFIRM].map(function (name) {
    if (exported[name]) return { name: name, html: exported[name], mode: 'export' };
    var sh = ss.getSheetByName(name);
    return { name: name, html: sh ? simpleHtml_(sh) : '<p>탭이 없습니다.</p>', mode: 'simple' };
  });
}

function exportHtml_(id) {
  var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + id + '/export?format=zip', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());
  var css = '', pages = {};
  Utilities.unzip(resp.getBlob().setContentType('application/zip')).forEach(function (f) {
    var n = f.getName();
    if (/\.css$/i.test(n)) css += f.getDataAsString('UTF-8');
    else if (/\.html$/i.test(n)) pages[n.replace(/^.*\//, '').replace(/\.html$/i, '')] = f.getDataAsString('UTF-8');
  });
  var extra = 'body{margin:8px;background:#fff}.row-headers-background,.column-headers-background,' +
    '.row-header-wrapper,.freezebar-cell,.freezebar-origin-ltr{display:none!important}';
  Object.keys(pages).forEach(function (k) {
    pages[k] = pages[k].replace(/<link[^>]+sheet\.css[^>]*>/i, '').replace(/<\/head>/i, '<style>' + css + extra + '</style></head>');
  });
  return pages;
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}

function simpleHtml_(sh) {
  var nr = Math.min(sh.getLastRow(), 80), nc = Math.min(sh.getLastColumn(), 30);
  if (!nr || !nc) return '<p>내용 없음</p>';
  var rg = sh.getRange(1, 1, nr, nc);
  var v = rg.getDisplayValues(), bg = rg.getBackgrounds(), fw = rg.getFontWeights(), fc = rg.getFontColors(),
    al = rg.getHorizontalAlignments(), fs = rg.getFontSizes(), wrap = rg.getWraps();
  var skip = {}, span = {};
  rg.getMergedRanges().forEach(function (m) {
    var r0 = m.getRow(), c0 = m.getColumn();
    if (r0 > nr || c0 > nc) return;
    span[r0 + ',' + c0] = [Math.min(m.getNumRows(), nr - r0 + 1), Math.min(m.getNumColumns(), nc - c0 + 1)];
    for (var i = 0; i < m.getNumRows(); i++) for (var j = 0; j < m.getNumColumns(); j++) if (i || j) skip[(r0 + i) + ',' + (c0 + j)] = 1;
  });
  var html = '<table style="border-collapse:collapse;font-family:Malgun Gothic,Arial,sans-serif;font-size:12px;table-layout:fixed"><colgroup>';
  for (var c = 1; c <= nc; c++) html += '<col style="width:' + sh.getColumnWidth(c) + 'px">';
  html += '</colgroup>';
  for (var r = 0; r < nr; r++) {
    html += '<tr style="height:' + sh.getRowHeight(r + 1) + 'px">';
    for (var k = 0; k < nc; k++) {
      var key = (r + 1) + ',' + (k + 1);
      if (skip[key]) continue;
      var sp = span[key], filled = v[r][k] !== '' || (bg[r][k] && bg[r][k] !== '#ffffff');
      var st = 'padding:2px 4px;vertical-align:middle;overflow:hidden;' + (wrap[r][k] ? 'white-space:pre-wrap;' : 'white-space:nowrap;') +
        'background:' + bg[r][k] + ';color:' + fc[r][k] + ';font-weight:' + fw[r][k] + ';font-size:' + fs[r][k] + 'pt;' +
        'text-align:' + (al[r][k] || 'general').replace('general-', '').replace('general', 'left') + ';' +
        (filled ? 'border:1px solid #bbb;' : '');
      html += '<td' + (sp ? ' rowspan="' + sp[0] + '" colspan="' + sp[1] + '"' : '') + ' style="' + st + '">' + esc_(v[r][k]) + '</td>';
    }
    html += '</tr>';
  }
  return html + '</table>';
}

/* ================================================================
 * 설정 점검 — 편집기에서 checkSetup 을 선택하고 ▶ 실행, 아래 "실행 로그"에서 결과 확인
 * ================================================================ */
function checkSetup() {
  var lines = [], ok = function (m) { lines.push('✅ ' + m); }, bad = function (m) { lines.push('❌ ' + m); };
  try {
    var values = masterValues_(), head = masterHeader_(values), contracts = parseContracts_(values);
    ok('자재정보: 머리글 ' + head.row + '행, 품목 ' + contracts.length + '줄');
    var keys = masterFields_(values).map(function (f) { return f.key; });
    ['item', 'priceMode', 'supplier', 'price', 'cur', 'start', 'end'].forEach(function (k) {
      if (keys.indexOf(k) < 0) bad('자재정보에 필수 열이 없습니다: ' + MASTER_COLUMNS[k][0]);
    });
  } catch (e) { bad('자재정보: ' + e.message); }

  var tid = prop_('TEMPLATE_ID');
  if (!tid) bad('스크립트 속성 TEMPLATE_ID 가 없습니다');
  else {
    try {
      var tpl = SpreadsheetApp.openById(tid);
      var missing = [TPL.SUMMARY, TPL.DETAIL, TPL.SCHEDULE, TPL.CONFIRM, TPL.PROTO].filter(function (n) { return !tpl.getSheetByName(n); });
      if (missing.length) bad('템플릿에 없는 탭: ' + missing.join(', ') + ' — templates/구매내역서_템플릿.xlsx 를 다시 올려 주세요');
      else ok('템플릿: ' + tpl.getName());
    } catch (e) { bad('템플릿을 열 수 없습니다 (ID·권한 확인): ' + e.message); }
  }
  var fid = prop_('OUTPUT_FOLDER_ID');
  if (fid) { try { ok('저장 폴더: ' + DriveApp.getFolderById(fid).getName()); } catch (e) { bad('OUTPUT_FOLDER_ID 폴더를 열 수 없습니다: ' + e.message); } }
  else lines.push('ℹ️ OUTPUT_FOLDER_ID 없음 — 구매내역서가 템플릿과 같은 위치에 저장됩니다');

  console.log(lines.join('\n'));
  return lines.join('\n');
}
