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
 *   RFQ_TEMPLATE_ID  (견적요청서) 견적요청서 템플릿 구글 시트 ID — templates/견적요청서_템플릿.xlsx
 *   ARCHIVE_SHEET_ID (완료 보관) 완료된 기록을 옮겨 둘 다른 스프레드시트 ID
 *   SPOT_TEMPLATE_ID (Spot 견적) Spot 견적 구매내역서 템플릿 구글 시트 ID — templates/Spot_구매내역서_템플릿.xlsx
 */
var SHEET_MASTER = '자재정보';
var SHEET_RATE = '환율';
var SHEET_VENDORS = ['업체정보', '업체목록'];   // 업체명 · 이메일 열 (한 업체 여러 줄, 한 칸에 ; , 로 여러 주소 가능)
var SHEET_RECORD = '진행기록';     // 내가 만든 구매내역서·견적요청서 (완료 전)
var RECORD_HEADERS = ['ID', '종류', '담당자', '작성일', '제목', '파일 URL', '파일 ID', '상태', '메일 초안', '숨은 참조', '품목(JSON)', '수정일'];

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
  var rates = checkedRates_(payload.rates, uniq_(lines.map(function (l) { return l.cur; })).filter(function (c) { return c !== 'KRW'; }));
  saveRates_(rates);
  var base = toDate_(payload.baseDate) || toDate_(new Date());
  // Spot 단가(DB) 품목은 따로 (계약과 같은 양식, 업체선정사유 Sole source)
  var groups = payload.split === false ? [{ spot: false, lines: lines }] : [
    { spot: false, lines: lines.filter(function (l) { return !isSpotMode_(l.priceMode); }) },
    { spot: true, lines: lines.filter(function (l) { return isSpotMode_(l.priceMode); }) }
  ].filter(function (g) { return g.lines.length; });
  var files = groups.map(function (g) {
    var sups = uniq_(g.lines.map(function (l) { return l.supplier; }));
    var name = String(payload.name || '').trim()
      ? purchaseFileName_(payload.name, base, sups) + (groups.length > 1 && g.spot ? ' (Spot)' : '')
      : purchaseFileName_('', base, sups).replace(' 구매내역서_', g.spot ? ' Spot 구매내역서_' : ' 구매내역서_');
    var curs = uniq_(g.lines.map(function (l) { return l.cur; })).filter(function (c) { return c !== 'KRW'; });
    var ss = copyTemplate_(name, payload.folder);
    applyPlan_(ss, buildPurchasePlan_(g.lines, curs, rates, base, payload.schedule, { title: name, confirmer: payload.confirmer }));
    var proto = ss.getSheetByName(TPL.PROTO);
    if (proto) ss.deleteSheet(proto);
    ss.setActiveSheet(ss.getSheetByName(TPL.SUMMARY));
    var rec = addRecord_({ kind: g.spot ? 'Spot 구매내역서' : '계약 구매내역서', title: name, url: ss.getUrl(), fileId: ss.getId(), status: '작성',
      items: g.lines.map(function (l) { return { item: l.item, qty: l.qty, needBy: fmtDate_(l.needBy), supplier: l.supplier, price: l.price, cur: l.cur }; }) });
    return { url: ss.getUrl(), name: name, id: ss.getId(), recordId: rec.id, spot: g.spot };
  });
  return { files: files, url: files[0].url, name: files[0].name, id: files[0].id, recordId: files[0].recordId };
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

/* ================================================================
 * 폴더 선택 창: 각자 내 드라이브의 폴더를 따라 들어가며 고른다
 * (웹앱 실행 사용자를 "웹 앱에 액세스하는 사용자"로 배포하면 각자 자기 드라이브가 보인다)
 * ================================================================ */
/** parentId 안의 폴더 목록. parentId 가 없으면 내 드라이브 맨 위 */
function apiListFolders(parentId) {
  var root = DriveApp.getRootFolder(), rootId = root.getId();
  var here;
  try { here = parentId ? DriveApp.getFolderById(parentId) : root; here.getName(); }
  catch (e) { throw new Error('폴더를 열 수 없습니다. 주소가 맞는지, 볼 수 있는 권한이 있는지 확인해 주세요.'); }
  var folders = [], it = here.getFolders();
  while (it.hasNext() && folders.length < 300) {
    var f = it.next();
    if (!f.isTrashed()) folders.push({ id: f.getId(), name: f.getName() });
  }
  folders.sort(function (a, b) { return a.name.localeCompare(b.name, 'ko'); });
  // 위치 표시: 내 드라이브 > … > 현재 폴더
  var path = [], cur = here, guard = 0;
  while (cur && guard++ < 15) {
    var isRoot = cur.getId() === rootId;
    path.unshift({ id: isRoot ? '' : cur.getId(), name: isRoot ? '내 드라이브' : cur.getName() });
    if (isRoot) break;
    var ps = cur.getParents();
    cur = ps.hasNext() ? ps.next() : null;
  }
  if (!path.length || path[0].name !== '내 드라이브') path.unshift({ id: '', name: '내 드라이브' });
  return { id: here.getId() === rootId ? '' : here.getId(), name: here.getId() === rootId ? '내 드라이브' : here.getName(), path: path, folders: folders };
}

/** 폴더 선택 창의 "새 폴더" */
function apiCreateFolder(parentId, name) {
  name = String(name || '').replace(/\s+/g, ' ').trim().slice(0, 100);
  if (!name) throw new Error('폴더 이름을 입력해 주세요.');
  var parent = parentId ? DriveApp.getFolderById(parentId) : DriveApp.getRootFolder();
  var f = parent.createFolder(name);
  return { id: f.getId(), name: f.getName() };
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
      case 'rowFmt': sh(o.fs).getRange(o.fr, o.c1, 1, o.c2 - o.c1 + 1).copyTo(sh(o.s).getRange(o.r, o.c1), FORMAT, false); break;
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
      case 'checkbox': sh(o.s).getRange(o.r1, o.c, o.r2 - o.r1 + 1, 1).insertCheckboxes(); break;
      case 'deleteCols': sh(o.s).deleteColumns(o.c, o.n); break;
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
  var rid = prop_('RFQ_TEMPLATE_ID');
  if (!rid) bad('스크립트 속성 RFQ_TEMPLATE_ID 가 없습니다 (견적요청서 템플릿)');
  else { try { var rt = SpreadsheetApp.openById(rid); if (rt.getSheetByName(RFQ.SHEET)) ok('견적요청서 템플릿: ' + rt.getName()); else bad('견적요청서 템플릿에 "' + RFQ.SHEET + '" 탭이 없습니다'); } catch (e) { bad('견적요청서 템플릿을 열 수 없습니다: ' + e.message); } }
  var sid = prop_('SPOT_TEMPLATE_ID');
  if (!sid) bad('스크립트 속성 SPOT_TEMPLATE_ID 가 없습니다 (Spot 견적 구매내역서 템플릿)');
  else { try { var st = SpreadsheetApp.openById(sid), miss = [SPOT.SUMMARY, SPOT.DETAIL, SPOT.QUOTE, SPOT.PROTO].filter(function (n) { return !st.getSheetByName(n); });
    if (miss.length) bad('Spot 템플릿에 없는 탭: ' + miss.join(', ')); else ok('Spot 견적 구매내역서 템플릿: ' + st.getName()); } catch (e) { bad('Spot 템플릿을 열 수 없습니다: ' + e.message); } }
  var aid = prop_('ARCHIVE_SHEET_ID');
  if (!aid) bad('스크립트 속성 ARCHIVE_SHEET_ID 가 없습니다 (완료 보관 파일)');
  else { try { ok('완료 보관 파일: ' + SpreadsheetApp.openById(aid).getName()); } catch (e) { bad('완료 보관 파일을 열 수 없습니다: ' + e.message); } }
  var vs = vendorSheet_();
  if (vs) ok(vs.getName() + ': ' + apiVendors().vendors.length + '개 업체'); else bad('자재정보 스프레드시트에 "업체정보" 탭이 없습니다 (업체명 · 이메일 열)');
  lines.push(typeof Gmail !== 'undefined' ? '✅ Gmail API 서비스: 사용 중 (Gmail 서명 자동 사용)' : 'ℹ️ Gmail API 서비스 없음 — 사이트에서 저장한 서명을 씁니다');
  var fid = prop_('OUTPUT_FOLDER_ID');
  if (fid) { try { ok('저장 폴더: ' + DriveApp.getFolderById(fid).getName()); } catch (e) { bad('OUTPUT_FOLDER_ID 폴더를 열 수 없습니다: ' + e.message); } }
  else lines.push('ℹ️ OUTPUT_FOLDER_ID 없음 — 구매내역서가 템플릿과 같은 위치에 저장됩니다');

  console.log(lines.join('\n'));
  return lines.join('\n');
}

/* ================================================================
 * 견적요청서 (RFQ) · Gmail 임시보관함 초안
 * ================================================================ */
/** 견적요청서 만들기. payload = {baseDate, title, due('yyyy-mm-dd'), folder, items:[{item, registered, name, desc, qpl, uom, remark, qty, needBy, program, shelfLife, shipFrom, taxCode}]} */
function apiCreateRfq(payload) {
  var items = (payload.items || []).map(function (it) {
    var o = {};
    Object.keys(it).forEach(function (k) { o[k] = it[k]; });
    o.needBy = toDate_(it.needBy);
    return o;
  });
  if (!items.length) throw new Error('견적요청서에 넣을 품목이 없습니다.');
  var base = toDate_(payload.baseDate) || toDate_(new Date());
  var title = String(payload.title || '').trim() || rfqTitle_(base, '', '', undefined, payload.due);
  var tid = prop_('RFQ_TEMPLATE_ID');
  if (!tid) throw new Error('스크립트 속성 RFQ_TEMPLATE_ID(견적요청서 템플릿 시트 ID)를 설정해 주세요. README 참고.');
  var folder = targetFolder_(payload.folder), tpl = DriveApp.getFileById(tid);
  var file = folder ? tpl.makeCopy(title, folder) : tpl.makeCopy(title);
  var ss = SpreadsheetApp.openById(file.getId());
  applyPlan_(ss, buildRfqPlan_(items, base));
  var rec = addRecord_({ kind: '견적요청서', title: title, url: ss.getUrl(), fileId: ss.getId(), status: '견적요청서 작성',
    items: (payload.items || []), due: payload.due || '' });
  return { url: ss.getUrl(), name: title, id: ss.getId(), recordId: rec.id };
}

/** 업체정보 탭: [{name, emails:[...]}] */
function vendorSheet_() {
  var ss = db_();
  for (var i = 0; i < SHEET_VENDORS.length; i++) { var sh = ss.getSheetByName(SHEET_VENDORS[i]); if (sh) return sh; }
  return null;
}
function apiVendors() {
  var sh = vendorSheet_();
  if (!sh) return { vendors: [], missing: true };
  return { vendors: parseVendors_(sh.getDataRange().getValues()), missing: false };
}

/** 메일 서명: Gmail 설정의 서명(고급 서비스 Gmail API 추가 시) → 없으면 사이트에 저장한 서명 */
function signature_() {
  try {
    if (typeof Gmail !== 'undefined') {
      var list = Gmail.Users.Settings.SendAs.list('me').sendAs || [];
      var me = list.filter(function (x) { return x.isPrimary; })[0] || list[0];
      if (me && me.signature) return { html: me.signature, source: 'gmail' };
    }
  } catch (e) { console.warn('Gmail 서명 읽기 실패: ' + e); }
  var saved = PropertiesService.getUserProperties().getProperty('SIGNATURE_HTML') || '';
  return { html: saved, source: saved ? 'saved' : 'none' };
}
function apiSignature() { return signature_(); }
function apiSaveSignature(text) {
  var html = String(text || '').trim();
  if (html && !/<[a-z][\s\S]*>/i.test(html)) html = esc_(html).replace(/\n/g, '<br>');   // 글자로 적으면 줄바꿈만 살림
  PropertiesService.getUserProperties().setProperty('SIGNATURE_HTML', html);
  return signature_();
}

/** Gmail 임시보관함에 초안 만들기. 받는 사람은 비우고 고른 업체만 숨은 참조. payload = {recordId, bcc:[...]} */
function apiCreateRfqDraft(payload) {
  var bcc = uniq_((payload.bcc || []).map(function (m) { return String(m).trim(); }).filter(function (m) { return /@/.test(m); }));
  if (!bcc.length) throw new Error('메일을 받을 업체를 한 곳 이상 골라 주세요.');
  var found = findRecord_(payload.recordId, true);
  var rec = found.rec;
  if (rec.kind !== '견적요청서') throw new Error('견적요청서 기록이 아닙니다.');
  var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + rec.fileId + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('견적요청서 파일을 엑셀로 내보내지 못했습니다 (HTTP ' + resp.getResponseCode() + ').');
  var blob = resp.getBlob().setName(rec.title + '.xlsx');
  var mail = rfqEmail_(rec.items, dueText_(rec.due), signature_().html);
  var draft = GmailApp.createDraft('', rec.title, mail.text, { htmlBody: mail.html, bcc: bcc.join(','), attachments: [blob] });
  var names = uniq_((payload.vendors || []).map(function (v) { return String(v).trim(); }));
  updateRecord_(found, { status: '메일 초안', draft: draft.getId(), bcc: bcc.join(', '), data: names.length ? { vendors: names } : null });
  return { draftId: draft.getId(), link: 'https://mail.google.com/mail/#drafts', bccCount: bcc.length };
}

/* ================================================================
 * 진행 기록 — 각자 자기 기록만. 완료하면 보관 파일(ARCHIVE_SHEET_ID)로 옮긴다
 * ================================================================ */
function me_() { return (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').toLowerCase(); }

function recordSheet_() {
  var ss = db_(), sh = ss.getSheetByName(SHEET_RECORD);
  if (!sh) {
    sh = ss.insertSheet(SHEET_RECORD);
    sh.getRange(1, 1, 1, RECORD_HEADERS.length).setValues([RECORD_HEADERS]).setFontWeight('bold').setBackground('#c9daf8');
    sh.setFrozenRows(1);
  }
  return sh;
}

function rowToRecord_(r) {
  var items = [];
  try { items = JSON.parse(r[10] || '[]'); } catch (e) {}
  var extra = items && !Array.isArray(items) ? items : null;
  return {
    id: String(r[0]), kind: r[1], owner: String(r[2]).toLowerCase(), created: Object.prototype.toString.call(r[3]) === '[object Date]' ? fmtDate_(r[3]) : String(r[3]),
    title: r[4], url: r[5], fileId: r[6], status: r[7], draft: r[8], bcc: r[9],
    items: extra ? extra.items : items, due: extra ? extra.due : '', vendors: extra && extra.vendors || [], quote: extra && extra.quote || null,
    updated: r[11] instanceof Date ? r[11].toISOString() : String(r[11] || '')
  };
}

function addRecord_(o) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = recordSheet_(), now = new Date();
    var id = Utilities.formatDate(now, 'Asia/Seoul', 'yyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
    sh.appendRow([id, o.kind, me_(), now, o.title, o.url, o.fileId, o.status, '', '', JSON.stringify({ items: o.items || [], due: o.due || '' }), now]);
    return { id: id };
  } finally { lock.releaseLock(); }
}

/** 기록 찾기. mine 이면 다른 담당자 기록은 거부 */
function findRecord_(id, mine) {
  var sh = recordSheet_(), values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      var rec = rowToRecord_(values[i]);
      if (mine && rec.owner !== me_()) throw new Error('다른 담당자의 기록입니다.');
      return { sheet: sh, row: i + 1, values: values[i], rec: rec };
    }
  }
  throw new Error('기록을 찾지 못했습니다. 새로고침해 주세요.');
}

function updateRecord_(found, patch) {
  var sh = found.sheet, r = found.row;
  if (patch.data) {   // 품목(JSON) 칸에 값 합치기
    var cur = {};
    try { cur = JSON.parse(found.values[10] || '{}'); } catch (e) {}
    if (Array.isArray(cur)) cur = { items: cur };
    Object.keys(patch.data).forEach(function (k) { cur[k] = patch.data[k]; });
    sh.getRange(r, 11).setValue(JSON.stringify(cur));
  }
  if (patch.status !== undefined) sh.getRange(r, 8).setValue(patch.status);
  if (patch.draft !== undefined) sh.getRange(r, 9).setValue(patch.draft);
  if (patch.bcc !== undefined) sh.getRange(r, 10).setValue(patch.bcc);
  sh.getRange(r, 12).setValue(new Date());
}

/** 내 진행 기록 (최근 순) */
function apiMyRecords() {
  var who = me_();
  var list = recordSheet_().getDataRange().getValues().slice(1).map(rowToRecord_)
    .filter(function (r) { return r.id && r.owner === who; });
  list.sort(function (a, b) { return String(b.id).localeCompare(String(a.id)); });
  return { me: who, records: list, archiveReady: !!prop_('ARCHIVE_SHEET_ID') };
}

/** 완료: 진행기록에서 보관 파일로 옮긴다 (본인 기록만) */
function apiCompleteRecord(id) {
  var aid = prop_('ARCHIVE_SHEET_ID');
  if (!aid) throw new Error('스크립트 속성 ARCHIVE_SHEET_ID(완료 보관 스프레드시트 ID)를 설정해 주세요.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var found = findRecord_(id, true);
    var arc = SpreadsheetApp.openById(aid), sh = arc.getSheetByName('완료기록');
    if (!sh) {
      sh = arc.insertSheet('완료기록');
      sh.getRange(1, 1, 1, RECORD_HEADERS.length + 1).setValues([RECORD_HEADERS.concat(['완료일'])]).setFontWeight('bold').setBackground('#c9daf8');
      sh.setFrozenRows(1);
    }
    var row = found.values.slice(0, RECORD_HEADERS.length);
    row[7] = '완료';
    sh.appendRow(row.concat([new Date()]));
    found.sheet.deleteRow(found.row);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ================================================================
 * Spot 견적 구매내역서 — 견적요청서 기록에서 견적값을 입력해 만든다
 * ================================================================ */
/** 견적 입력 임시 저장. quote = {vendors, items:[{prev, program, quotes}], rates} */
function apiSaveQuotes(recordId, quote) {
  var found = findRecord_(recordId, true);
  if (found.rec.kind !== '견적요청서') throw new Error('견적요청서 기록이 아닙니다.');
  updateRecord_(found, { data: { quote: quote } });
  return { ok: true, saved: new Date().toISOString() };
}

/**
 * payload = {recordId, baseDate, rates, folder, name, confirmer, schedule, vendors:[업체명],
 *            items:[{item, name, desc, material, qpl, uom, remark, qty, needBy, program, prev, quotes}]}
 */
function apiCreateSpotPurchase(payload) {
  var found = findRecord_(payload.recordId, true);
  if (found.rec.kind !== '견적요청서') throw new Error('견적요청서 기록이 아닙니다.');
  var tid = prop_('SPOT_TEMPLATE_ID');
  if (!tid) throw new Error('스크립트 속성 SPOT_TEMPLATE_ID(Spot 견적 구매내역서 템플릿 시트 ID)를 설정해 주세요. README 참고.');
  var items = (payload.items || []).map(function (it) {
    var o = {};
    Object.keys(it).forEach(function (k) { o[k] = it[k]; });
    o.needBy = toDate_(it.needBy);
    return o;
  });
  if (!items.length) throw new Error('구매내역서에 넣을 품목이 없습니다.');
  var vendors = uniq_((payload.vendors || []).map(function (v) { return String(v).trim(); }));
  var rates = {};
  Object.keys(payload.rates || {}).forEach(function (k) { var n = Number(payload.rates[k]); if (n > 0) rates[k] = n; });
  if (Object.keys(rates).length) saveRates_(rates);
  var base = toDate_(payload.baseDate) || toDate_(new Date());
  var name = String(payload.name || '').trim() || found.rec.title.replace(/^\[KE\] RFQ \/\/ /, '') + ' 견적 구매내역서';
  var folder = targetFolder_(payload.folder), tpl = DriveApp.getFileById(tid);
  var file = folder ? tpl.makeCopy(name, folder) : tpl.makeCopy(name);
  var ss = SpreadsheetApp.openById(file.getId());
  applyPlan_(ss, buildSpotPlan_(items, vendors, rates, base, payload.schedule, { title: name, confirmer: payload.confirmer }));
  var proto = ss.getSheetByName(SPOT.PROTO);
  if (proto) ss.deleteSheet(proto);
  ss.setActiveSheet(ss.getSheetByName(SPOT.SUMMARY));
  updateRecord_(found, { status: '견적 구매내역서 작성', data: { quote: { vendors: vendors, items: payload.items, rates: payload.rates }, spotFile: ss.getUrl() } });
  var rec = addRecord_({ kind: '견적 구매내역서', title: name, url: ss.getUrl(), fileId: ss.getId(), status: '작성',
    items: items.map(function (it) { return { item: it.item, qty: it.qty, needBy: fmtDate_(it.needBy) }; }) });
  return { url: ss.getUrl(), name: name, id: ss.getId(), recordId: rec.id };
}
