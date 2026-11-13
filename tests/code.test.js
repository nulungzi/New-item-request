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
  // 여러 줄: 하나라도 틀리면 아무것도 안 씀, 같은 묶음 안의 겹치는 SPOT 도 막음
  const bad = ctx.apiRegisterItems([{ item: 'C1', priceMode: 'SPOT', supplier: 'X' }, { item: 'C1', priceMode: 'spot', supplier: 'x' }, {}, { item: 'C2', priceMode: 'LTA' }]);
  assert.strictEqual(bad.ok, false);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(bad.errors.map(e => [e.i, e.item]))), [[1, 'C1'], [3, 'C2']]);
  assert.match(bad.errors[0].messages[0], /이번 1번째 줄/);
  assert.strictEqual(written.length, 1);
  const good = ctx.apiRegisterItems([{ item: 'D1', priceMode: 'SPOT', supplier: 'X' }, { item: 'D2', priceMode: 'LTA', supplier: 'PPG', price: '2.5', cur: 'usd', start: '2027-01-01', end: '2027-12-31' }, {}]);
  assert.deepStrictEqual(JSON.parse(JSON.stringify([good.ok, good.added.map(a => [a.row, a.item])])), [true, [[4, 'D1'], [5, 'D2']]]);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(written[1].v.map(r => [r[0], r[1], r[4], r[5]]))), [[2, 'D1', '', ''], [3, 'D2', 2.5, 'USD']]);
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
      getLastRow: () => sh.rows.length,
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
  assert.deepStrictEqual(JSON.parse(JSON.stringify(ctx.apiVendors(), (k, x) => ['people', 'supplier', 'site'].includes(k) ? undefined : x)), { vendors: [{ name: 'Hexcel', emails: ['a@hexcel.com', 'b@hexcel.com'] }, { name: '3M', emails: ['q@3m.com'] }], missing: false });

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

  // 재요청: 같은 견적요청서에 이어서, 제목 [Reminder] + 새 회신 기한, 업체는 기존에 더함, 메일 기록 2건
  assert.throws(() => ctx.apiCreateRfqDraft({ recordId: r.recordId, bcc: [] , remind: true }), /한 곳 이상/);
  ctx.updateRecord_(ctx.findRecord_(r.recordId, true), { data: { vendors: ['Hexcel', '3M'] } });
  const d2 = ctx.apiCreateRfqDraft({ recordId: r.recordId, bcc: ['q@3m.com', 'new@aero.com'], vendors: ['3M', 'Aerovac'], remind: true, due: '2026-11-25' });
  assert.strictEqual(d2.subject, '[Reminder] [KE] RFQ // 261110 Boeing Chemical Material Spot (due 11/25)');
  assert.match(drafts[1].opt.htmlBody, /kind reminder of our RFQ/);
  assert.match(drafts[1].opt.htmlBody, /by 11\/25\./);
  assert.strictEqual(drafts[1].opt.bcc, 'q@3m.com,new@aero.com');
  const rr = ctx.apiMyRecords().records.find(x => x.id === r.recordId);
  assert.deepStrictEqual(JSON.parse(JSON.stringify([rr.status, rr.vendors, rr.mails.map(m => [m.no, m.kind, m.count]), rr.due])),
    ['재요청 초안', ['Hexcel', '3M', 'Aerovac'], [[1, '견적 요청', 2], [2, '재요청', 2]], '2026-11-25']);

  let mine = ctx.apiMyRecords();
  assert.deepStrictEqual([mine.records.length, mine.records[0].status, mine.records[0].bcc], [1, '재요청 초안', 'q@3m.com, new@aero.com']);
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

  // 하나로 만들기: 계약 줄(A1)을 견적요청서에 붙여 두고, 견적 구매내역서를 만들 때 같이 넣고 묶음·부모를 기록
  const rfq2 = ctx.apiCreateRfq({ baseDate: '2026-11-13', title: 'R2', due: '2026-11-20', group: 'G1',
    items: [{ item: 'MS1', name: 'MS1', uom: 'EA', qty: 1, needBy: '2027-01-15' }],
    fixed: [{ item: 'A1', qty: 2, needBy: '2027-06-10', contractRow: 2 }, { item: 'A1', qty: 3, needBy: '2027-07-10', contractRow: 2 }] });
  let r2 = ctx.apiMyRecords().records.find(r => r.id === rfq2.recordId);
  assert.deepStrictEqual(JSON.parse(JSON.stringify([r2.group, r2.fixed.length, r2.fixed[0].supplier, r2.fixed[0].qty, r2.fixed[0].needs])),
    ['G1', 1, 'AKZO', 5, [{ needBy: '2027-06-10', qty: 2 }, { needBy: '2027-07-10', qty: 3 }]]);
  ctx.updateRecord_(ctx.findRecord_(rfq2.recordId, true), { data: { vendors: ['Hexcel'] } });
  const sp2 = ctx.apiCreateSpotPurchase({ recordId: rfq2.recordId, baseDate: '2026-11-13', rates: { USD: 1350 }, vendors: ['Hexcel'],
    items: [{ item: 'MS1', uom: 'EA', qty: 1, needBy: '2027-01-15', quotes: { Hexcel: { sel: true, price: 4, cur: 'USD' } } }] });
  const child = ctx.apiMyRecords().records.find(r => r.id === sp2.recordId);
  assert.deepStrictEqual([child.parent, child.group, child.items.map(x => x.item).join(',')], [rfq2.recordId, 'G1', 'A1,MS1']);
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

test('apiCreateQuoteSheet: 업체마다 탭, 사이트 표와 같은 값, 기록에 시트 주소', () => {
  const calls = [];
  const chain = new Proxy(function () {}, { get: (t, k) => (k === 'getId' ? () => 'QS' : k === 'getUrl' ? () => 'QSURL' : k === 'getSheets' ? () => [chain, chain] : (...a) => { calls.push([k, a]); return chain; }), apply: () => chain });
  const ctx = inContext({
    SpreadsheetApp: { create: () => chain, newDataValidation: () => chain, BorderStyle: { SOLID: 'S' } },
    DriveApp: { getFileById: () => chain, getFolderById: () => chain },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
  });
  let saved = null;
  const rec = { kind: '견적요청서', title: '[KE] RFQ // 261113 Material Spot (due 11/14)', vendors: ['OMI', 'CBOL'],
    items: [{ item: 'KERN00736-001', name: '88TB0110', qty: 1, uom: 'SH', needBy: '2027-03-01' }] };
  ctx.findRecord_ = () => ({ rec });
  ctx.updateRecord_ = (f, p) => { saved = p; };
  const d = ctx.apiCreateQuoteSheet('R1', '');
  assert.deepStrictEqual([d.url, d.name], ['QSURL', '261113 Material Spot (due 11/14) 견적 입력']);
  assert.deepStrictEqual(calls.filter(c => c[0] === 'insertSheet').map(c => c[1][0]), ['OMI', 'CBOL']);
  const vals = calls.filter(c => c[0] === 'setValues').map(c => c[1][0]);
  assert.strictEqual(JSON.stringify(vals[1][1].slice(0, 5)), JSON.stringify([1, 'KERN00736-001', '88TB0110', 1, '2027-03-01']));
  assert.strictEqual(JSON.stringify(saved), JSON.stringify({ data: { quoteSheet: 'QSURL' } }));
  rec.quoteSheet = 'QSURL';
  assert.strictEqual(ctx.apiCreateQuoteSheet('R1', '').url, 'QSURL');   // 이미 있으면 그 시트
});

test('동시 사용: 다른 사람이 완료해서 줄이 밀려도 기록 고치기는 ID 로 다시 찾아 그 줄에만 씀', () => {
  const db = fakeBook({ 진행기록: [['ID', '종류', '담당자', '작성일', '제목', 'URL', '파일ID', '상태', '초안', 'BCC', 'JSON', '수정일'],
    ['A', '견적요청서', 'me@ke.com', '', 'A', '', '', '작성', '', '', '{"items":[]}', ''],
    ['B', '견적요청서', 'me@ke.com', '', 'B', '', '', '작성', '', '', '{"items":[1]}', '']] });
  const ctx = inContext({
    SpreadsheetApp: { getActive: () => db }, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
  });
  const found = ctx.findRecord_('B', true);          // B 는 3행
  db.tabs['진행기록'].rows.splice(1, 1);             // 그 사이 다른 사람이 A 를 완료 → B 가 2행으로
  ctx.updateRecord_(found, { status: '메일 초안', data: { vendors: ['OMI'] } });
  const rows = db.tabs['진행기록'].rows;
  assert.strictEqual(rows.length, 2);
  assert.deepStrictEqual([rows[1][0], rows[1][7], JSON.parse(rows[1][10]).vendors[0], JSON.parse(rows[1][10]).items[0]], ['B', '메일 초안', 'OMI', 1]);
  db.tabs['진행기록'].rows.splice(1, 1);             // B 도 완료됨
  assert.throws(() => ctx.updateRecord_(found, { status: 'x' }), /이미 완료/);
});

test('자재정보 캐시: 큰 값도 나눠 저장·복원, 날짜는 글자로, 등록·새로고침이면 새로 읽음', () => {
  const store = {};
  let reads = 0;
  const rows = [['Item', '계약 구분', '계약시작', 'Description']].concat(Array.from({ length: 900 }, (_, i) => ['K' + i, 'LTA', new Date(2026, 0, 1), '설명 '.repeat(20)]));
  const sheet = { getDataRange: () => ({ getValues: () => { reads++; return rows; } }) };
  const ctx = inContext({
    SpreadsheetApp: { getActive: () => ({ getSheetByName: () => sheet }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    CacheService: { getScriptCache: () => ({ putAll: m => Object.assign(store, m), get: k => store[k] || null,
      getAll: ks => Object.fromEntries(ks.map(k => [k, store[k]])), remove: k => { delete store[k]; } }) },
  });
  const a = ctx.masterValues_();
  assert.ok(Number(store.master_n) > 1);             // 여러 조각
  assert.strictEqual(a[1][2], '2026-01-01');
  const b = ctx.masterValues_();
  assert.strictEqual(reads, 1);                      // 두 번째는 캐시
  assert.strictEqual(JSON.stringify(a), JSON.stringify(b));
  ctx.masterValues_(true); assert.strictEqual(reads, 2);   // DB 새로고침
  ctx.cacheDrop_('master'); ctx.masterValues_(); assert.strictEqual(reads, 3);   // 자재 등록 뒤
  assert.strictEqual(ctx.parseContracts_(b)[0].start.getFullYear(), 2026);
});

test('PO 상태 기록 · 완료 기록 보기', () => {
  const db = fakeBook({ 진행기록: [['ID', '종류', '담당자', '작성일', '제목', 'URL', '파일ID', '상태', '초안', 'BCC', 'JSON', '수정일'],
    ['P1', '계약 구매내역서', 'me@ke.com', '', '구매', '', '', '작성', '', '', '{"items":[]}', ''],
    ['Q1', '견적요청서', 'me@ke.com', '', 'RFQ', '', '', '메일 초안', '', '', '{"items":[]}', '']] });
  const arc = fakeBook({ 완료기록: [['ID', '종류', '담당자', '작성일', '제목', 'URL', '파일ID', '상태', '초안', 'BCC', 'JSON', '수정일', '완료일'],
    ['A1', '계약 구매내역서', 'me@ke.com', '', '옛 구매', '', '', '완료', '', '', '{"items":[]}', '', new Date(2026, 9, 1)],
    ['A2', '계약 구매내역서', 'lee@ke.com', '', '남의 것', '', '', '완료', '', '', '{"items":[]}', '', new Date(2026, 9, 2)],
    ['A3', '견적요청서', 'me@ke.com', '', '최근', '', '', '완료', '', '', '{"items":[]}', '', new Date(2026, 10, 1)]] });
  const ctx = inContext({
    SpreadsheetApp: { getActive: () => db, openById: () => arc }, PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'ARCHIVE_SHEET_ID' ? 'arc' : null) }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
  });
  const r = ctx.apiSetPoStatus('P1', { stage: 'PO 발행', po: ' 4500012345 ', poDate: '2026/11/20', memo: '', bogus: 'x' });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(r.po)), { stage: 'PO 발행', po: '4500012345', poDate: '2026-11-20' });
  const rec = ctx.apiMyRecords().records.find(x => x.id === 'P1');
  assert.deepStrictEqual([rec.status, rec.po.po], ['PO 발행', '4500012345']);
  assert.throws(() => ctx.apiSetPoStatus('Q1', { stage: 'PO 발행' }), /견적요청서는/);
  const a = ctx.apiArchivedRecords();
  assert.deepStrictEqual(a.records.map(x => [x.id, x.done]), [['A3', '2026-11-01'], ['A1', '2026-10-01']]);
});

test('완료 + 구매 이력: 후보 줄(계약·견적 구매내역서), 뺀 줄 제외하고 올림, 안 올리고 완료', () => {
  const H = ['ID', '종류', '담당자', '작성일', '제목', 'URL', '파일ID', '상태', '초안', 'BCC', 'JSON', '수정일'];
  const db = fakeBook({ 진행기록: [H,
    ['C1', '계약 구매내역서', 'me@ke.com', '', '계약', 'u1', 'f1', '작성', '', '', JSON.stringify({ group: 'G', items: [{ item: 'A1', qty: 5, supplier: 'AKZO', price: 2, cur: 'USD', priceMode: 'LTA', uom: 'KT', reqDate: '2026-11-10' }, { item: 'A2', qty: 1, supplier: 'AKZO', price: 900, cur: 'USD', priceMode: 'LTA' }] }), ''],
    ['Q1', '견적요청서', 'me@ke.com', '', 'RFQ', 'u2', 'f2', '메일 초안', '', '', JSON.stringify({ group: 'G', reqDates: ['2026-11-12 EFAB'], items: [{ item: 'K1', parts: [{ reqDate: '2026-11-11' }] }] }), ''],
    ['S1', '견적 구매내역서', 'me@ke.com', '', '견적', 'u3', 'f3', '작성', '', '', JSON.stringify({ parent: 'Q1', group: 'G', items: [] }), '']] });
  const arc = fakeBook({});
  const q = Array.from({ length: 4 }, () => []);
  const r = Array(47).fill(''); r[0] = true; r[1] = '최저가'; r[5] = 'K1'; r[29] = 'CBOL'; r[30] = 3; r[33] = 11.8; r[34] = 'USD'; q.push(r);
  const quoteBook = { getSheetByName: n => (n === '견적종합' ? { getDataRange: () => ({ getValues: () => q }) } : null) };
  const ctx = inContext({
    SpreadsheetApp: { getActive: () => db, openById: id => (id === 'arc' ? arc : quoteBook) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'ARCHIVE_SHEET_ID' ? 'arc' : null) }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
  });
  const c = ctx.apiHistoryCandidates(['C1', 'Q1', 'S1']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(c.rows.map(h => [h.key, h.item, h.supplier, h.price, h.mode, h.reqDate]))),
    [['C1#0', 'A1', 'AKZO', 2, 'LTA', '2026-11-10'], ['C1#1', 'A2', 'AKZO', 900, 'LTA', ''], ['S1#0', 'K1', 'CBOL', 11.8, 'SPOT 견적', '2026-11-11']]);
  // A2 는 긴급 고가 → 빼고 올림
  const done = ctx.apiCompleteGroup({ ids: ['C1', 'Q1', 'S1'], upload: true, skip: ['C1#1'] });
  assert.deepStrictEqual([done.added, done.completed], [2, 3]);
  const hist = db.tabs['구매이력'].rows;
  assert.deepStrictEqual(hist.slice(1).map(r => [r[2], r[4], r[6], r[16]]), [['A1', 'AKZO', 2, 'me@ke.com'], ['K1', 'CBOL', 11.8, 'me@ke.com']]);
  assert.strictEqual(db.tabs['진행기록'].rows.length, 1);
  assert.strictEqual(arc.tabs['완료기록'].rows.length, 4);
  // 이력 조회 · 직전 구매
  db.insertSheet('자재정보').rows.push(['Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료'], ['K1', 'SPOT', 'Airtech', 80, 'USD', '2026-01-01', '2027-12-31']);
  const ih = ctx.apiItemHistory('k1');
  assert.deepStrictEqual([ih.contracts.length, ih.history.length, ih.history[0]['업체']], [1, 1, 'CBOL']);
  assert.strictEqual(ctx.apiLastPurchases(['K1']).K1.price, 11.8);
  // 안 올리고 완료
  db.tabs['진행기록'].rows.push(['C2', '계약 구매내역서', 'me@ke.com', '', '계약2', '', '', '작성', '', '', '{"items":[{"item":"Z"}]}', '']);
  ctx.apiCompleteGroup({ ids: ['C2'], upload: false });
  assert.strictEqual(db.tabs['구매이력'].rows.length, 3);
});
