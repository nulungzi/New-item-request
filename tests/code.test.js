// Code.js 가 Apps Script 없이도 끝까지 실행되는지 가짜 객체로 확인
const test = require('node:test');
const assert = require('node:assert');
const { inContext, source } = require('./load.js');
const vm = require('vm');

test('apiCreatePurchase: 템플릿 복사 후 계획 적용', () => {
  const calls = [];
  const range = new Proxy({}, { get: (t, k) => (...a) => { calls.push(k); return range; } });
  const sheet = name => ({ getRange: () => range, setRowHeight: () => {}, setColumnWidth: () => calls.push('setColumnWidth'), setFrozenRows: () => {}, setFrozenColumns: () => {},
    getDataRange: () => ({ getValues: () => name === '자재정보'
      ? [['Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료'], ['A1', 'LTA', 'AKZO', 2, 'USD', new Date(2026, 0, 1), new Date(2028, 11, 31)]]
      : [['CUR', 'KRW 환율', '수정일']] }), appendRow: () => {} });
  const ss = { getSheetByName: n => sheet(n), deleteSheet: () => calls.push('deleteSheet'), setActiveSheet: () => {}, getUrl: () => 'URL', getId: () => 'new' };
  const ctx = {
    SpreadsheetApp: { getActive: () => ss, openById: () => ss, CopyPasteType: { PASTE_FORMAT: 'F' }, BorderStyle: { SOLID: 'S', SOLID_MEDIUM: 'M' },
      newDataValidation: () => range },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'TEMPLATE_ID' ? 'tpl' : null) }) },
    DriveApp: { getFileById: () => ({ makeCopy: () => ({ getId: () => 'new' }) }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
    Utilities: { formatDate: () => '261110-120000' },
  };
  inContext(ctx);
  const res = ctx.apiCreatePurchase({ baseDate: '2026-11-10', rates: { USD: 1350 }, remarks: {}, lines: [{ item: 'A1', qty: 3, needBy: '2027-06-10', contractRow: 2 }] });
  assert.strictEqual(res.name, '261110 구매내역서_AKZO');
  assert.strictEqual(res.id, 'new');
  assert.ok(calls.includes('copyTo') && calls.includes('merge') && calls.includes('deleteSheet'));
});

for (const f of ['Logic.gs', 'Code.gs']) {
  test(f + ': 첫 줄, 보이지 않는 문자, 문법', () => {
    const src = source(f);
    assert.ok(src.startsWith('// ===== 소요제기 판정 웹앱 · ' + f), '첫 줄이 바뀌었습니다');
    const bad = [...src].filter(ch => { const c = ch.codePointAt(0); return (c < 32 && c !== 9 && c !== 10 && c !== 13) || [0xa0, 0x2028, 0x2029, 0xfeff].includes(c); });
    assert.deepStrictEqual(bad, [], '보이지 않는 문자가 있습니다');
    assert.ok(!/module\.exports|require\(/.test(src), 'Node 전용 코드가 들어 있습니다');
    assert.doesNotThrow(() => new vm.Script(src, { filename: f }));
  });
}

test('파일 읽는 순서와 상관없이 올라가고, 판정 함수는 Logic.gs 에 있다', () => {
  for (const order of [['Logic.gs', 'Code.gs'], ['Code.gs', 'Logic.gs']]) assert.doesNotThrow(() => inContext({}, order), order.join(' → '));
  const onlyCode = inContext({}, ['Code.gs']);
  assert.strictEqual(typeof onlyCode.parseRows_, 'undefined');
  const both = inContext({});
  for (const fn of ['parseRows_', 'classify_', 'buildPurchasePlan_', 'apiClassify', 'apiCreatePurchase', 'checkSetup']) assert.strictEqual(typeof both[fn], 'function', fn);
});

test('checkSetup: 자재정보·템플릿 점검', () => {
  const values = [
    ['', 'Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료', 'Shelf Life', 'Description'],
    [1, 'A1', 'LTA', 'PPG', 1, 'USD', new Date(2026, 0, 1), new Date(2028, 11, 31), '12MM', 'PRIMER'],
  ];
  const sheet = { getName: () => '업체정보', getDataRange: () => ({ getValues: () => values }) };
  const tabs = ['구매요약', '구매세부내역', '구매일정', '사업부 담당 확인'];
  const ctx = inContext({
    console: { log: () => {}, warn: () => {} },
    SpreadsheetApp: { getActive: () => ({ getSheetByName: () => sheet }), openById: () => ({ getName: () => 'T', getSheetByName: n => (tabs.includes(n) ? {} : null) }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'TEMPLATE_ID' ? 'tpl' : null) }) },
  });
  const report = ctx.checkSetup();
  assert.match(report, /✅ 자재정보: 머리글 1행, 품목 1줄/);
  assert.match(report, /❌ 템플릿에 없는 탭: _proto/);
  assert.match(report, /RFQ_TEMPLATE_ID 가 없습니다/);
  assert.match(report, /ARCHIVE_SHEET_ID 가 없습니다/);
});

test('apiRegisterItem: 자재정보 맨 아래에 머리글 순서대로 추가, 번호 열 +1', () => {
  const values = [
    ['', '', '', '', '', '', '', '', '', '', ''],
    ['', 'Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료', 'MOQ', 'Ship Form', 'Description'],
    [1, 'A1', 'LTA', 'PPG', 1, 'USD', new Date(2026, 0, 1), new Date(2026, 11, 31), '없음', 'Air', 'X'],
  ];
  const written = [];
  const rng = (r, c, nr, nc) => ({ setValues: v => written.push({ r, c, v }), copyTo: () => {}, setNumberFormat: () => {} });
  const sheet = { getDataRange: () => ({ getValues: () => values }), getLastRow: () => values.length, getRange: rng };
  const ctx = {
    SpreadsheetApp: { getActive: () => ({ getSheetByName: () => sheet }), CopyPasteType: { PASTE_FORMAT: 'F' } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
  };
  inContext(ctx);
  const res = ctx.apiRegisterItem({ item: 'kern99999-001', priceMode: 'SPOT', supplier: 'Hexcel', shipFrom: 'Air', desc: 'PREPREG' });
  assert.deepStrictEqual([res.row, res.item, res.mode], [4, 'KERN99999-001', 'SPOT']);
  assert.strictEqual(written.length, 1);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(written[0].v[0])), [2, 'KERN99999-001', 'SPOT', 'Hexcel', '', '', '', '', '', 'Air', 'PREPREG']);
  assert.throws(() => ctx.apiRegisterItem({ item: 'B', priceMode: 'LTA' }), /업체명/);
  // 초기화: 등록 양식 필드는 시트 머리글 순서
  const init = ctx.apiInit();
  assert.deepStrictEqual(JSON.parse(JSON.stringify(init.fields.map(f => f.key))), ['item', 'priceMode', 'supplier', 'price', 'cur', 'start', 'end', 'moq', 'shipFrom', 'desc']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(init.items[0])), ['A1', 'X', 'Air', 'LTA', '']);
});

test('apiCreateInputSheet · apiCheckFolder: 고른 폴더로 저장', () => {
  const calls = [];
  const any = new Proxy(function () {}, { get: (t, k) => (k === 'getName' ? () => '구매팀 폴더' : k === 'getUrl' ? () => 'FURL' : k === 'getId' ? () => 'NEW' : (...a) => { calls.push(k); return any; }), apply: () => any });
  const sheet = new Proxy({}, { get: (t, k) => (k === 'getMaxColumns' ? () => 26 : (...a) => { calls.push(k); return k === 'setName' ? sheet : any; }) });
  const ctx = inContext({
    SpreadsheetApp: { create: n => ({ getSheets: () => [sheet], insertSheet: () => sheet, setActiveSheet: () => {}, getId: () => 'NEW', getUrl: () => 'URL', getName: () => n }),
      newDataValidation: () => any, BorderStyle: { SOLID: 'S', SOLID_MEDIUM: 'M' } },
    DriveApp: { getFolderById: id => { if (id !== '1AbCdEfGhIjKlMnOp') throw new Error('x'); return any; }, getFileById: () => ({ moveTo: f => calls.push('moveTo') }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
  });
  const r = ctx.apiCreateInputSheet('https://drive.google.com/drive/folders/1AbCdEfGhIjKlMnOp', '');
  assert.deepStrictEqual([r.name, r.folder], ['소요제기 입력 시트', '구매팀 폴더']);
  assert.ok(calls.includes('moveTo') && calls.includes('setFrozenRows') && calls.includes('deleteColumns'));
  assert.strictEqual(ctx.apiCheckFolder('').name, '내 드라이브 (기본 위치)');
  assert.throws(() => ctx.apiCheckFolder('https://drive.google.com/drive/folders/1ZZZZZZZZZZZZZZZZZ'), /열 수 없습니다/);
  assert.throws(() => ctx.apiCheckFolder('내 폴더'), /주소를 확인/);
});

test('apiListFolders · apiCreateFolder: 내 드라이브 폴더 탐색', () => {
  const mkFolder = (id, name, parent, kids = []) => ({ getId: () => id, getName: () => name, isTrashed: () => false,
    getParents: () => { let d = !parent; return { hasNext: () => !d, next: () => { d = true; return parent; } }; },
    getFolders: () => { let i = 0; return { hasNext: () => i < kids.length, next: () => kids[i++] }; },
    createFolder: n => mkFolder('new1', n, null) });
  const root = mkFolder('ROOT', 'My Drive', null);
  const team = mkFolder('T', '구매팀', root);
  const kids = [mkFolder('B', '소요제기', team), mkFolder('A', '견적', team)];
  Object.assign(team, { getFolders: () => { let i = 0; return { hasNext: () => i < kids.length, next: () => kids[i++] }; } });
  Object.assign(root, { getFolders: () => { let i = 0; const k = [team]; return { hasNext: () => i < k.length, next: () => k[i++] }; } });
  const byId = { T: team, A: kids[1], B: kids[0] };
  const ctx = inContext({ DriveApp: { getRootFolder: () => root, getFolderById: id => { if (!byId[id]) throw new Error('x'); return byId[id]; } } });
  const top = JSON.parse(JSON.stringify(ctx.apiListFolders('')));
  assert.deepStrictEqual(top, { id: '', name: '내 드라이브', path: [{ id: '', name: '내 드라이브' }], folders: [{ id: 'T', name: '구매팀' }] });
  const t = JSON.parse(JSON.stringify(ctx.apiListFolders('T')));
  assert.deepStrictEqual(t.path.map(p => p.name), ['내 드라이브', '구매팀']);
  assert.deepStrictEqual(t.folders.map(f => f.name), ['견적', '소요제기']);
  assert.strictEqual(ctx.apiCreateFolder('T', '  2027 소요 ').name, '2027 소요');
  assert.throws(() => ctx.apiCreateFolder('T', ' '), /폴더 이름/);
  assert.throws(() => ctx.apiListFolders('NOPE'), /폴더를 열 수 없습니다/);
});


/* ---------- 견적요청서 · 메일 초안 · 진행 기록 (가짜 시트에 실제로 값이 쌓이게) ---------- */
function fakeBook(tabs) {
  const mk = (name, rows) => {
    const sh = {
      name, rows,
      getDataRange: () => ({ getValues: () => sh.rows.map(r => r.slice()) }),
      appendRow: r => sh.rows.push(r.slice()),
      deleteRow: n => sh.rows.splice(n - 1, 1),
      getRange: (r, c, nr, nc) => ({
        setValues: v => { v.forEach((row, i) => row.forEach((x, j) => { (sh.rows[r - 1 + i] = sh.rows[r - 1 + i] || [])[c - 1 + j] = x; })); return chain; },
        setValue: x => { (sh.rows[r - 1] = sh.rows[r - 1] || [])[c - 1] = x; return chain; },
        setFontWeight: () => chain, setBackground: () => chain,
      }),
      setFrozenRows: () => {},
    };
    const chain = { setFontWeight: () => chain, setBackground: () => chain };
    return sh;
  };
  const book = { tabs: {} };
  Object.entries(tabs).forEach(([n, rows]) => (book.tabs[n] = mk(n, rows)));
  book.getSheetByName = n => book.tabs[n] || null;
  book.insertSheet = n => (book.tabs[n] = mk(n, []));
  return book;
}

test('견적요청서 → 메일 초안 → 내 기록 → 완료 보관', () => {
  const db = fakeBook({ 업체목록: [['업체명', '이메일'], ['Hexcel', 'a@hexcel.com; b@hexcel.com'], ['3M', 'q@3m.com'], ['', ''], ['Hexcel', 'a@hexcel.com']] });
  const archive = fakeBook({});
  let user = 'me@ke.com', drafts = [], applied = 0;
  const rfqBook = { getUrl: () => 'RFQURL', getId: () => 'RFQID', getSheetByName: () => new Proxy({}, { get: () => () => new Proxy({}, { get: (t, k) => () => (applied++, t) }) }) };
  const props = { RFQ_TEMPLATE_ID: 'tpl', ARCHIVE_SHEET_ID: 'arc' };
  const ctx = inContext({
    console: { warn: () => {}, log: () => {} },
    SpreadsheetApp: { getActive: () => db, openById: id => (id === 'arc' ? archive : rfqBook), CopyPasteType: { PASTE_FORMAT: 'F' }, BorderStyle: { SOLID: 'S', SOLID_MEDIUM: 'M' } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null }), getUserProperties: () => ({ getProperty: () => '<b>김민형 사원</b>', setProperty: () => {} }) },
    DriveApp: { getFileById: () => ({ makeCopy: () => ({ getId: () => 'RFQID' }) }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => user }), getEffectiveUser: () => ({ getEmail: () => user }) },
    Utilities: { formatDate: () => '261110-120000' },
    UrlFetchApp: { fetch: () => ({ getResponseCode: () => 200, getBlob: () => ({ setName: n => ({ name: n }) }) }) },
    ScriptApp: { getOAuthToken: () => 't' },
    GmailApp: { createDraft: (to, subj, text, opt) => { drafts.push({ to, subj, text, opt }); return { getId: () => 'D1' }; } },
  });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(ctx.apiVendors())), { vendors: [{ name: 'Hexcel', emails: ['a@hexcel.com', 'b@hexcel.com'] }, { name: '3M', emails: ['q@3m.com'] }], missing: false });

  const title = '[KE] RFQ // 261110 Boeing Chemical Material Spot (due 11/20)';
  const r = ctx.apiCreateRfq({ baseDate: '2026-11-10', title, due: '2026-11-20', items: [
    { item: 'MS20995C20', registered: false, name: 'MS20995C20', desc: 'Safety Wire', qpl: 'MALIN Co', uom: 'RL', remark: 'MFG COC/TR', qty: 3, needBy: '2027-01-15' }] });
  assert.ok(applied > 0);
  assert.strictEqual(r.name, title);
  assert.throws(() => ctx.apiCreateRfqDraft({ recordId: r.recordId, bcc: [] }), /한 곳 이상/);
  const d = ctx.apiCreateRfqDraft({ recordId: r.recordId, bcc: ['a@hexcel.com', 'q@3m.com', 'a@hexcel.com', 'x'] });
  assert.strictEqual(d.bccCount, 2);
  assert.strictEqual(drafts[0].to, '');                        // 받는 사람 비움
  assert.strictEqual(drafts[0].opt.bcc, 'a@hexcel.com,q@3m.com');
  assert.strictEqual(drafts[0].subj, title);
  assert.match(drafts[0].opt.htmlBody, /due 11\/20\./);
  assert.match(drafts[0].opt.htmlBody, /김민형 사원/);
  assert.strictEqual(drafts[0].opt.attachments[0].name, title + '.xlsx');

  let mine = ctx.apiMyRecords();
  assert.deepStrictEqual([mine.records.length, mine.records[0].status, mine.records[0].bcc], [1, '메일 초안', 'a@hexcel.com, q@3m.com']);
  user = 'other@ke.com';                                          // 다른 담당자: 안 보이고 완료도 못 함
  assert.strictEqual(ctx.apiMyRecords().records.length, 0);
  assert.throws(() => ctx.apiCompleteRecord(r.recordId), /다른 담당자/);
  user = 'me@ke.com';
  ctx.apiCompleteRecord(r.recordId);
  assert.strictEqual(ctx.apiMyRecords().records.length, 0);
  assert.strictEqual(archive.tabs['완료기록'].rows.length, 2);   // 머리글 + 1건
  assert.strictEqual(archive.tabs['완료기록'].rows[1][7], '완료');

  // 템플릿 2 (목록형): RFQ2_TEMPLATE_ID, 같은 자재도 Need-by 마다 한 줄로 기록
  assert.throws(() => ctx.apiCreateRfq({ template: 'list', baseDate: '2026-11-10', items: [{ item: 'A', name: 'A', qty: 1 }] }), /RFQ2_TEMPLATE_ID/);
  props.RFQ2_TEMPLATE_ID = 'tpl2';
  const r2 = ctx.apiCreateRfq({ template: 'list', baseDate: '2026-11-10', title: 'RFQ_261110_KE_HW_YWK_AH-64E', due: '2026-11-20', items: [
    { item: 'KERN00534-001', name: 'PC938', uom: 'EA', qty: 488, needBy: '2027-03-01', parts: [{ qty: 244, needBy: '2027-03-01', program: 'AH-64E' }, { qty: 244, needBy: '2027-09-01', program: 'AH-64E' }] }] });
  assert.strictEqual(r2.name, 'RFQ_261110_KE_HW_YWK_AH-64E');
  const rec2 = ctx.apiMyRecords().records.find(x => x.id === r2.recordId);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(rec2.items.map(x => [x.qty, x.needBy]))), [[244, '2027-03-01'], [244, '2027-09-01']]);
});

test('견적 저장 → Spot 견적 구매내역서, 계약+Spot 단가는 두 파일로', () => {
  const db = fakeBook({});
  const ops = [];
  const book = (name) => ({ getUrl: () => 'URL-' + name, getId: () => 'ID-' + name, getName: () => name, deleteSheet: () => {}, setActiveSheet: () => {},
    getSheetByName: () => { const rec = new Proxy({}, { get: (t, k) => (...a) => { ops.push(k); return k === 'getValue' ? 'Need-by' : rec; } }); return rec; } });
  const made = [];
  const ctx = inContext({
    console: { warn: () => {}, log: () => {} },
    SpreadsheetApp: { getActive: () => db, openById: id => book(id), CopyPasteType: { PASTE_FORMAT: 'F' }, BorderStyle: { SOLID: 'S', SOLID_MEDIUM: 'M' },
      newDataValidation: () => { const b = new Proxy({}, { get: (t, k) => () => (k === 'build' ? {} : b) }); return b; },
      newConditionalFormatRule: () => { const b = new Proxy({}, { get: (t, k) => () => (k === 'build' ? {} : b) }); return b; } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => ({ TEMPLATE_ID: 'tpl', SPOT_TEMPLATE_ID: 'spot', RFQ_TEMPLATE_ID: 'rfq' })[k] || null }) },
    DriveApp: { getFileById: () => ({ makeCopy: n => (made.push(n), { getId: () => n }) }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
    Utilities: { formatDate: () => '261110-120000' },
  });
  // 자재정보: LTA 1개 + Spot 단가 1개
  db.insertSheet('자재정보').rows.push(['Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료'],
    ['A1', 'LTA', 'AKZO', 2, 'USD', new Date(2026, 0, 1), new Date(2028, 11, 31)],
    ['S1', 'Spot', 'Airtech', 462, 'USD', new Date(2025, 0, 1), new Date(2027, 11, 31)]);
  db.insertSheet('환율');
  const res = ctx.apiCreatePurchase({ baseDate: '2026-11-10', rates: { USD: 1350 }, remarks: {}, lines: [
    { item: 'A1', qty: 3, needBy: '2027-06-10', contractRow: 2 }, { item: 'S1', qty: 30, needBy: '2027-06-10', contractRow: 3 }] });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(res.files.map(f => [f.name, f.spot]))), [['261110 구매내역서_AKZO', false], ['261110 Spot 구매내역서_Airtech', true]]);

  // 견적요청서 기록 → 업체 이름 저장 → 견적 임시 저장 → Spot 견적 구매내역서
  const rfq = ctx.apiCreateRfq({ baseDate: '2026-11-10', title: '[KE] RFQ // 261110 Boeing Chemical Material Spot (due 11/20)', due: '2026-11-20',
    items: [{ item: 'MS20995C20', registered: false, name: 'MS20995C20', desc: 'Safety Wire', uom: 'RL', qty: 3, needBy: '2027-01-15' }] });
  const found = ctx.findRecord_(rfq.recordId, true);
  ctx.updateRecord_(found, { data: { vendors: ['Hexcel', '3M'] } });
  const quote = { vendors: ['Hexcel', '3M'], items: [{ quotes: { Hexcel: { sel: true, review: '최저가', price: 10, cur: 'USD' } } }] };
  ctx.apiSaveQuotes(rfq.recordId, quote);
  let mine = ctx.apiMyRecords().records.find(r => r.id === rfq.recordId);
  assert.deepStrictEqual(JSON.parse(JSON.stringify([mine.vendors, mine.quote])), [['Hexcel', '3M'], quote]);
  const sp = ctx.apiCreateSpotPurchase({ recordId: rfq.recordId, baseDate: '2026-11-10', rates: { USD: 1350 }, vendors: ['Hexcel', '3M'],
    items: [{ item: 'MS20995C20', desc: 'Safety Wire', uom: 'RL', qty: 3, needBy: '2027-01-15', program: '', quotes: quote.items[0].quotes }] });
  assert.strictEqual(sp.name, '261110 Boeing Chemical Material Spot (due 11/20) 견적 구매내역서');
  assert.ok(ops.includes('insertCheckboxes') && ops.includes('deleteColumns'));
  mine = ctx.apiMyRecords().records;
  assert.deepStrictEqual(mine.map(r => [r.kind, r.status]).sort(), [['Spot 구매내역서', '작성'], ['견적 구매내역서', '작성'], ['견적요청서', '견적 구매내역서 작성'], ['계약 구매내역서', '작성']].sort());
});

test('로그인: 입력한 이메일 = 접속한 Google 계정이어야 하고, 처음이면 사용자 탭에 등록, 4시간 뒤 만료', () => {
  const mkCtx = (active, effective) => {
    const props = {}, rows = [];
    const users = { appendRow: r => rows.push(r), setFrozenRows: () => {}, getDataRange: () => ({ getValues: () => rows }), getRange: (r, c) => ({ setValue: v => { rows[r - 1][c - 1] = v; } }) };
    let hasUsers = false;
    const db = { getSheetByName: n => (n === '사용자' && hasUsers ? users : null), insertSheet: () => { hasUsers = true; return users; } };
    const ctx = {
      SpreadsheetApp: { getActive: () => db },
      PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }),
        getUserProperties: () => ({ getProperty: k => props[k] || null, setProperties: o => Object.assign(props, o), deleteProperty: k => { delete props[k]; } }) },
      LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
      Session: { getActiveUser: () => ({ getEmail: () => active }), getEffectiveUser: () => ({ getEmail: () => effective }) },
      ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/X/exec' }) },
    };
    inContext(ctx, null, { realAuth: true });
    return { ctx, props, rows };
  };
  // "나" 권한으로 배포 → 다른 사람이 열면 막고 배포 방법 안내
  const bad = mkCtx('', 'owner@ke.com');
  assert.strictEqual(bad.ctx.apiAuthState().deployOk, false);
  assert.throws(() => bad.ctx.apiLogin('kim@ke.com'), /웹 앱에 액세스하는 사용자/);
  assert.throws(() => bad.ctx.apiMyRecords(), /LOGIN_REQUIRED/);

  const { ctx, props, rows } = mkCtx('Kim@KE.com', 'Kim@KE.com');
  assert.strictEqual(ctx.apiAuthState().loggedIn, false);
  assert.throws(() => ctx.apiMyRecords(), /LOGIN_REQUIRED/);
  assert.throws(() => ctx.apiLogin('lee@ke.com'), /지금 접속한 Google 계정\(kim@ke.com\)이 다릅니다/);
  assert.throws(() => ctx.apiLogin('kim'), /정확히/);
  const st = ctx.apiLogin('  KIM@ke.com ');
  assert.strictEqual(st.loggedIn, true);
  assert.strictEqual(st.isNew, true);
  assert.strictEqual(rows.length, 2);                     // 머리글 + 등록
  assert.strictEqual(rows[1][0], 'kim@ke.com');
  assert.strictEqual(st.expiresAt - Number(props.LOGIN_AT), 4 * 3600000);
  assert.strictEqual(ctx.auth_(), 'kim@ke.com');
  assert.strictEqual(ctx.apiLogin('kim@ke.com').isNew, false);   // 두 번째는 등록하지 않음
  assert.strictEqual(rows.length, 2);
  props.LOGIN_AT = String(Date.now() - 4 * 3600000 - 1000);     // 4시간 지남
  assert.strictEqual(ctx.apiAuthState().loggedIn, false);
  assert.throws(() => ctx.auth_(), /LOGIN_REQUIRED/);
  ctx.apiLogin('kim@ke.com'); ctx.apiLogout();
  assert.strictEqual(ctx.apiAuthState().loggedIn, false);
});
