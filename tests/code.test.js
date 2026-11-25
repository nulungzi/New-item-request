// Code.js 가 Apps Script 없이도 끝까지 실행되는지 가짜 객체로 확인
const test = require('node:test');
const assert = require('node:assert');
const { inContext, source } = require('./load.js');
const vm = require('vm');

test('apiCreatePurchase: 템플릿 복사 후 계획 적용', () => {
  const calls = [];
  const range = new Proxy({}, { get: (t, k) => (...a) => { calls.push(k); return range; } });
  const tplHead = { '7,21': '수량', '7,23': 'SITE' };   // 새 계약 템플릿 (Need-by 옆 수량)
  const sheet = name => ({ getRange: (r, c) => (tplHead[r + ',' + c] ? new Proxy({}, { get: (t, k) => (k === 'getValue' ? () => tplHead[r + ',' + c] : range[k]) }) : range), setRowHeight: () => {}, getMaxColumns: () => 200, setColumnWidth: () => calls.push('setColumnWidth'), setFrozenRows: () => {}, setFrozenColumns: () => {},
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
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
  });
  const report = ctx.checkSetup();
  assert.match(report, /✅ 자재정보 탭: 머리글 1행, 품목 1줄/);
  assert.match(report, /❌ 구매내역서 템플릿: 없는 탭 _proto/);
  assert.match(report, /RFQ_TEMPLATE_ID 가 없습니다/);
  assert.match(report, /ARCHIVE_SHEET_ID 가 없습니다/);
  assert.match(report, /— 잠금 · 캐시/);   // 묶음 제목
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
  const res = ctx.apiRegisterItem({ item: 'kern99999-001', priceMode: 'SPOT', supplier: 'Hexcel', shipFrom: 'Air', desc: 'PREPREG', uom: 'RL' });
  assert.deepStrictEqual([res.row, res.item, res.mode], [4, 'KERN99999-001', 'SPOT']);
  assert.strictEqual(written.length, 1);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(written[0].v[0])), [2, 'KERN99999-001', 'SPOT', 'Hexcel', '', '', '', '', '', 'Air', 'PREPREG']);
  assert.throws(() => ctx.apiRegisterItem({ item: 'B', priceMode: 'LTA' }), /업체명/);
  // 여러 줄: 하나라도 틀리면 아무것도 안 씀, 같은 묶음 안의 겹치는 SPOT 도 막음
  const bad = ctx.apiRegisterItems([{ item: 'C1', priceMode: 'SPOT', supplier: 'X', uom: 'EA', desc: 'D' }, { item: 'C1', priceMode: 'spot', supplier: 'x', uom: 'EA', desc: 'D' }, {}, { item: 'C2', priceMode: 'LTA' }]);
  assert.strictEqual(bad.ok, false);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(bad.errors.map(e => [e.i, e.item]))), [[1, 'C1'], [3, 'C2']]);
  assert.match(bad.errors[0].messages[0], /이번 1번째 줄/);
  assert.strictEqual(written.length, 1);
  const good = ctx.apiRegisterItems([{ item: 'D1', priceMode: 'SPOT', supplier: 'X', uom: 'EA', desc: 'D' }, { item: 'D2', priceMode: 'LTA', supplier: 'PPG', price: '2.5', cur: 'usd', start: '2027-01-01', end: '2027-12-31' }, {}]);
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
    getSheetByName: () => { let at = ''; const rec = new Proxy({}, { get: (t, k) => k === 'getMaxColumns' ? () => 200 : (...a) => { ops.push(k); if (k === 'getRange') at = a[0] + ',' + a[1]; return k === 'getValue' ? ({ '9,23': 'Need-by', '9,24': '수량', '9,26': 'Tax Code', '3,39': 'Need-by 구분', '7,21': '수량', '7,23': 'SITE' })[at] || '' : rec; } }); return rec; } });
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
  const spRec = ctx.apiMyRecords().records.find(r => r.id === sp.recordId);
  assert.deepStrictEqual([spRec.fx.rates.USD.v, typeof spRec.fx.at], [1350, 'string']);   // 확정 환율
  const cRec = ctx.apiMyRecords().records.find(r => r.id === res.files[0].recordId);
  assert.strictEqual(cRec.fx.rates.USD.v, 1350);
  assert.ok(ops.includes('insertCheckboxes') && ops.includes('deleteColumns'));
  mine = ctx.apiMyRecords().records;
  assert.deepStrictEqual(mine.map(r => [r.kind, r.status]).sort(), [['Spot 구매내역서', '작성'], ['견적 구매내역서', '작성'], ['견적요청서', '견적 구매내역서 작성'], ['계약 구매내역서', '작성']].sort());
  // 이미 만든 견적요청서로 또 누르면 한 번 더 확인 (again 이면 만듦) · 다른 창에서 견적을 먼저 저장했으면 멈춤
  const again = { recordId: rfq.recordId, baseDate: '2026-11-10', rates: { USD: 1350 }, vendors: ['Hexcel', '3M'], items: [{ item: 'MS20995C20', uom: 'RL', qty: 3, needBy: '2027-01-15', quotes: quote.items[0].quotes }] };
  assert.throws(() => ctx.apiCreateSpotPurchase(again), /^Error: SPOT_EXISTS\|.*\|이 견적요청서로 이미/);
  assert.throws(() => ctx.apiCreateSpotPurchase(Object.assign({ again: true, base: 'old-save' }, again)), /다른 창/);
  assert.ok(ctx.apiCreateSpotPurchase(Object.assign({ again: true, base: sp.savedAt }, again)).url);

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
  assert.strictEqual(JSON.stringify(saved.data), JSON.stringify({ quoteSheet: 'QSURL' }));
  assert.strictEqual(saved.log, '견적 입력용 시트');
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

test('사업부 확인 요청 메일: 초안·보내기, 받는 사람 기억, 기록 상태', () => {
  const db = fakeBook({ 진행기록: [['ID', '종류', '담당자', '작성일', '제목', 'URL', '파일ID', '상태', '초안', 'BCC', 'JSON', '수정일'],
    ['P1', '계약 구매내역서', 'me@ke.com', '', '261113 구매내역서', 'https://docs/p1', '', '작성', '', '', '{"items":[]}', ''],
    ['Q1', '견적요청서', 'me@ke.com', '', 'RFQ', '', '', '메일 초안', '', '', '{"items":[]}', ''],
    ['X1', '계약 구매내역서', 'lee@ke.com', '', '남의 것', '', '', '작성', '', '', '{"items":[]}', '']] });
  const up = {}, drafts = [], sent = [];
  const ctx = inContext({
    SpreadsheetApp: { getActive: () => db },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }),
      getUserProperties: () => ({ getProperty: k => up[k] || null, setProperty: (k, v) => { up[k] = v; }, setProperties: o => Object.assign(up, o) }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
    Utilities: { formatDate: () => '2026-11-13 09:00' },
    GmailApp: { createDraft: (to, s, t, o) => { drafts.push({ to, s, o }); }, sendEmail: (to, s, t, o) => { sent.push({ to, s, o }); } },
  });
  assert.throws(() => ctx.apiConfirmMail({ recordId: 'Q1', to: 'a@ke.com' }), /구매내역서 기록/);
  assert.throws(() => ctx.apiConfirmMail({ recordId: 'X1', to: 'a@ke.com' }), /다른 담당자/);
  assert.throws(() => ctx.apiConfirmMail({ recordId: 'P1', to: '' }), /받는 사람/);
  assert.throws(() => ctx.apiConfirmMail({ recordId: 'P1', to: 'a@ke.com, nope' }), /이메일 형식/);
  const r = ctx.apiConfirmMail({ recordId: 'P1', to: 'Biz@KE.com; b@ke.com', cc: 'c@ke.com', name: '김민형', mode: 'draft' });
  assert.strictEqual(r.sent, false);
  assert.strictEqual(sent.length, 0);
  assert.deepStrictEqual([drafts[0].to, drafts[0].s, drafts[0].o.cc], ['biz@ke.com,b@ke.com', '261113 구매내역서', 'c@ke.com']);
  assert.ok(drafts[0].o.htmlBody.indexOf('<a href="https://docs/p1">261113 구매내역서</a>') > 0);
  assert.strictEqual(db.tabs['진행기록'].rows[1][7], '사업부 확인 초안');
  assert.strictEqual(up.CONFIRM_TO, 'biz@ke.com, b@ke.com');
  ctx.apiConfirmMail({ recordId: 'P1', to: 'biz@ke.com', subject: '날인 요청', mode: 'send' });
  assert.deepStrictEqual([sent.length, sent[0].s], [1, '날인 요청']);
  assert.strictEqual(db.tabs['진행기록'].rows[1][7], '사업부 확인 요청');
  assert.strictEqual(JSON.parse(db.tabs['진행기록'].rows[1][10]).confirmMail.mode, 'send');
  assert.strictEqual(ctx.apiConfirmDefaults().to, 'biz@ke.com');
});

test('완료 기록 삭제: 본인 기록만, 구매 이력 줄도 같이 (선택)', () => {
  const H = ['ID', '종류', '담당자', '작성일', '제목', 'URL', '파일ID', '상태', '초안', 'BCC', 'JSON', '수정일', '완료일'];
  const arc = fakeBook({ 완료기록: [H,
    ['A1', '견적요청서', 'me@ke.com', '', 'T1', '', '', '완료', '', '', '{}', '', ''],
    ['A2', '계약 구매내역서', 'lee@ke.com', '', '남의 것', '', '', '완료', '', '', '{}', '', ''],
    ['A3', '계약 구매내역서', 'me@ke.com', '', 'T3', '', '', '완료', '', '', '{}', '', ''],
    ['A4', '계약 구매내역서', 'me@ke.com', '', '남김', '', '', '완료', '', '', '{}', '', '']] });
  const L = require('./load.js').load();
  const hrow = (id, who) => { const r = Array(L.HISTORY.HEAD.length).fill(''); r[L.HISTORY.HEAD.indexOf('기록 ID')] = id; r[L.HISTORY.HEAD.indexOf('담당자')] = who; return r; };
  const db = fakeBook({ 구매이력: [L.HISTORY.HEAD.slice(), hrow('A3', 'me@ke.com'), hrow('A4', 'me@ke.com'), hrow('A2', 'lee@ke.com'), hrow('A3', 'lee@ke.com')] });
  const ctx = inContext({
    SpreadsheetApp: { getActive: () => db, openById: () => arc },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'ARCHIVE_SHEET_ID' ? 'arc' : null) }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
  });
  assert.throws(() => ctx.apiDeleteArchived([]), /골라/);
  const r = ctx.apiDeleteArchived(['A1', 'A2', 'A3'], true);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(r)), { deleted: 2, skipped: 1, history: 1 });
  assert.deepStrictEqual(arc.tabs['완료기록'].rows.slice(1).map(x => x[0]), ['A2', 'A4']);
  assert.deepStrictEqual(db.tabs['구매이력'].rows.slice(1).map(x => x[L.HISTORY.HEAD.indexOf('기록 ID')]), ['A4', 'A2', 'A3']);
  assert.strictEqual(ctx.apiDeleteArchived(['A4'], false).history, 0);
  assert.strictEqual(db.tabs['구매이력'].rows.length, 4);
  // 바로 지우지 않고 삭제기록 탭으로 (같이 뺀 구매이력 줄도 보관) → 되돌리기
  const trash = arc.tabs['삭제기록'].rows;
  assert.deepStrictEqual(trash.slice(1).map(x => x[0]), ['A3', 'A1', 'A4']);
  assert.strictEqual(JSON.parse(trash[1][15]).length, 1);
  const del = ctx.apiDeletedRecords().records;
  assert.deepStrictEqual(del.map(x => x.id).sort(), ['A1', 'A3', 'A4']);
  assert.strictEqual(del.filter(x => x.id === 'A3')[0].history, 1);
  const back = ctx.apiRestoreDeleted(['A3', 'A2']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(back)), { restored: 1, history: 1 });
  assert.deepStrictEqual(arc.tabs['완료기록'].rows.slice(1).map(x => x[0]), ['A2', 'A3']);
  assert.deepStrictEqual(arc.tabs['삭제기록'].rows.slice(1).map(x => x[0]), ['A1', 'A4']);
  assert.ok(db.tabs['구매이력'].rows.slice(1).some(x => x[L.HISTORY.HEAD.indexOf('기록 ID')] === 'A3' && x[L.HISTORY.HEAD.indexOf('담당자')] === 'me@ke.com'));
  // 변경기록
  const log = db.tabs['변경기록'].rows.slice(1).map(x => x[2] + ':' + x[3]);
  assert.deepStrictEqual(log.filter(x => /삭제/.test(x)), ['삭제 (삭제기록으로):A3', '삭제 (삭제기록으로):A1', '삭제 (삭제기록으로):A4', '삭제 되돌리기:A3']);
});

test('완료 취소: 진행 중으로 되돌리고 상태 복원, 올린 구매 이력 줄 지움 · 업체 추천은 이 기록 빼고', () => {
  const H = ['ID', '종류', '담당자', '작성일', '제목', 'URL', '파일ID', '상태', '초안', 'BCC', 'JSON', '수정일'];
  const L = require('./load.js').load();
  const hrow = (id, who, item, sup) => { const r = Array(L.HISTORY.HEAD.length).fill(''); r[2] = item; r[4] = sup; r[L.HISTORY.HEAD.indexOf('기록 ID')] = id; r[L.HISTORY.HEAD.indexOf('담당자')] = who; return r; };
  const quote = { rates: { USD: 1400 }, items: [{ quotes: { OMI: { price: '10', cur: 'USD' }, CBOL: { price: '9', cur: 'USD' } } }] };
  const arc = fakeBook({ 완료기록: [H.concat(['완료일']),
    ['A1', '견적요청서', 'me@ke.com', '2026-10-01', 'RFQ old', '', '', '완료', 'D', '', JSON.stringify({ items: [{ item: 'K1' }], vendors: ['OMI', 'CBOL'], quote }), '', ''],
    ['A2', '계약 구매내역서', 'me@ke.com', '', '구매', '', '', '완료', '', '', JSON.stringify({ items: [], po: { stage: 'PO 발행' } }), '', ''],
    ['A3', '계약 구매내역서', 'lee@ke.com', '', '남의 것', '', '', '완료', '', '', '{}', '', '']] });
  const db = fakeBook({ 진행기록: [H, ['Q9', '견적요청서', 'me@ke.com', '2026-11-15', 'RFQ new', '', '', '메일 초안', 'D', '', JSON.stringify({ items: [{ item: 'k1' }], vendors: ['OMI'] }), '']],
    구매이력: [L.HISTORY.HEAD.slice(), hrow('A2', 'me@ke.com', 'K1', 'CBOL'), hrow('A2', 'lee@ke.com', 'K1', 'CBOL'), hrow('Z', 'me@ke.com', 'K1', 'PPG')] });
  const ctx = inContext({
    SpreadsheetApp: { getActive: () => db, openById: () => arc },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'ARCHIVE_SHEET_ID' ? 'arc' : null) }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
  });
  // 추천: A1(완료) 견적 + 구매이력, Q9 자신은 뺌
  const hints = JSON.parse(JSON.stringify(ctx.apiVendorHints('Q9')));
  assert.deepStrictEqual(hints.map(h => [h.name, h.quoted, h.lowest, h.bought]), [['CBOL', 1, 1, 2], ['PPG', 0, 0, 1], ['OMI', 1, 0, 0]]);
  assert.strictEqual(ctx.apiVendorStats().length, 3);
  const r = ctx.apiRestoreArchived(['A1', 'A2', 'A3']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(r)), { restored: 2, skipped: 1, history: 1 });
  const rows = db.tabs['진행기록'].rows;
  assert.deepStrictEqual(rows.slice(1).map(x => [x[0], x[7]]), [['Q9', '메일 초안'], ['A2', 'PO 발행'], ['A1', '메일 초안']]);
  assert.deepStrictEqual(arc.tabs['완료기록'].rows.slice(1).map(x => x[0]), ['A3']);
  assert.strictEqual(db.tabs['구매이력'].rows.length, 3);
  assert.ok(ctx.apiMyRecords().records.some(x => x.id === 'A1'));
});

test('파일 만들기 속도: 붙은 칸 쓰기를 묶어도 결과는 칸마다 하나씩 쓸 때와 똑같음', () => {
  const L = require('./load.js').load();
  const d = s => L.toDate_(s);
  // 시트 호출을 칸 단위 사건으로 기록하는 가짜 스프레드시트
  const mkBook = () => {
    const log = [], calls = { n: 0 };
    const norm = v => (typeof v === 'string' && v.charAt(0) === '=') ? ['f', v] : v instanceof Date ? ['d', v.toISOString()] : ['v', v];
    const rng = (s, r, c, nr = 1, nc = 1) => {
      const each = (f) => { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) f(r + i, c + j, i, j); };
      const ev = (k, x) => each((rr, cc) => log.push([s, rr, cc, k, x]));
      const range = {
        setValue: v => { calls.n++; if (v === undefined) throw new Error('undefined'); ev(...norm(v)); return range; },
        setFormula: f => { calls.n++; ev('f', f); return range; },
        setValues: m => { calls.n++; each((rr, cc, i, j) => { if (m[i][j] === undefined) throw new Error('undefined in setValues'); log.push([s, rr, cc, ...norm(m[i][j])]); }); return range; },
        setBackground: x => { calls.n++; ev('bg', x); return range; },
        setBackgrounds: m => { calls.n++; each((rr, cc, i, j) => log.push([s, rr, cc, 'bg', m[i][j]])); return range; },
        setNumberFormat: x => { calls.n++; ev('nf', x); return range; },
        setNumberFormats: m => { calls.n++; each((rr, cc, i, j) => log.push([s, rr, cc, 'nf', m[i][j]])); return range; },
        copyTo: (t) => { calls.n++; log.push(['copy', s, r, c, nr, nc, t._at]); },
        merge: () => { calls.n++; log.push(['merge', s, r, c, nr, nc]); }, breakApart: () => { calls.n++; log.push(['unmerge', s, r, c, nr, nc]); },
        setBorder: (...a) => { calls.n++; log.push(['border', s, r, c, nr, nc, JSON.stringify(a)]); return range; },
        setDataValidation: () => { calls.n++; log.push(['dv', s, r, c, nr, nc]); }, insertCheckboxes: () => { calls.n++; log.push(['cb', s, r, c, nr, nc]); },
        setWrap: x => { calls.n++; log.push(['wrap', s, r, c, x]); }, clear: () => { calls.n++; log.push(['clear', s, r, c, nr, nc]); },
        _at: [s, r, c],
      };
      return range;
    };
    const sheets = {};
    const sheet = n => sheets[n] || (sheets[n] = { getRange: (r, c, nr, nc) => rng(n, r, c, nr, nc),
      setRowHeight: (r, h) => { calls.n++; log.push(['h', n, r, h]); }, setColumnWidth: (c, w) => { calls.n++; log.push(['w', n, c, w]); },
      setFrozenRows: x => { calls.n++; log.push(['fr', n, x]); }, setFrozenColumns: x => { calls.n++; log.push(['fc', n, x]); },
      getMaxColumns: () => 200, deleteColumns: (c, k) => { calls.n++; log.push(['delc', n, c, k]); }, getConditionalFormatRules: () => [], setConditionalFormatRules: x => { calls.n++; log.push(['cf', n, x.length]); } });
    return { log, calls, ss: { getSheetByName: sheet, getSpreadsheetTimeZone: () => 'Asia/Seoul' } };
  };
  const chain = new Proxy({}, { get: () => () => chain });
  const ctx = inContext({
    SpreadsheetApp: { CopyPasteType: { PASTE_FORMAT: 'F' }, BorderStyle: { SOLID: 'S', SOLID_MEDIUM: 'M' }, newDataValidation: () => chain, newConditionalFormatRule: () => chain },
    Utilities: { parseDate: (s) => new Date(s + 'T00:00:00+09:00') },
  });
  const lines = Array.from({ length: 22 }, (_, i) => ({ item: 'KERC' + (100 + i), qty: 5 + i, needBy: d('2027-0' + (1 + i % 8) + '-10'), supplier: i % 2 ? 'AKZO' : 'PPG', cur: i % 3 ? 'USD' : 'KRW',
    price: 10 + i, priceMode: 'LTA', material: '화공', desc: 'PRIMER ' + i, product: 'P', shipFrom: i % 4 ? 'Ocean' : '', leadtime: i % 5 ? '10wks' : '', uom: 'KT', remark: i % 6 ? '' : '분납', program: i % 7 ? 'AH-64E' : '',
    needs: i === 3 ? [{ needBy: '2027-01-10', qty: 3 }, { needBy: '2027-04-10', qty: 5 }] : undefined }));
  const base = d('2026-11-15');
  const spotItems = [
    { item: 'KERN00278-001', desc: 'FEP FILM', material: 'FEP', qpl: 'Airtech', uom: 'RL', qty: 64, needBy: d('2027-05-10'), program: '공통', prev: { year: 2026, qty: 24, cur: 'usd', price: 964.19 },
      quotes: { Airtech: { sel: true, review: 'Sole source', qty: 64, mfg: 'A4000V', uom: 'RL', price: 1, cur: 'usd', lt: '4~6' }, Hexcel: { price: 'no bid' }, '3M': { price: 2, cur: 'EUR', qty: 64 } } },
    { item: 'MS20995C20', desc: 'Safety Wire', uom: 'RL', qty: 3, needBy: d('2027-01-15'), program: '', quotes: {} },
    L.fixedSpotItem_({ item: 'KERC00064', supplier: 'AKZO', qty: 1100, uom: 'KT', price: 78.49, cur: 'USD', priceMode: 'LTA', needBy: '2027-05-14', needs: [{ needBy: '2027-05-14', qty: 500 }, { needBy: '2027-06-20', qty: 600 }], desc: 'PRIMER' })];
  const rfqItems = [{ item: 'MS20995C20', registered: false, name: 'MS20995C20', desc: 'Safety Wire', qpl: 'MALIN', uom: 'RL', qty: 3, needBy: d('2027-01-15') },
    { item: 'KERC00016', registered: true, name: 'PR-1776', desc: 'PRIMER', qpl: 'PPG', uom: 'KT', remark: '', qty: 5, needBy: d('2027-03-01'), parts: [{ qty: 2, needBy: '2027-03-01' }, { qty: 3, needBy: '2027-06-01' }] }];
  const plans = {
    contract: L.buildPurchasePlan_(lines, ['USD'], { USD: 1400 }, base).concat(L.scheduleConfirmOps_(base, null, {})),
    spot: L.buildSpotPlan_(spotItems, ['Airtech', 'Hexcel', '3M'], { USD: 1355.6, EUR: 1740.98 }, base, null, { title: 'T', confirmer: '홍길동' }),
    rfq: L.buildRfqPlan_(rfqItems, base),
    rfq2: L.buildRfq2Plan_(L.splitRfqItems_(rfqItems)),
  };
  for (const [name, ops] of Object.entries(plans)) {
    const one = mkBook(), many = mkBook();
    ctx.applyPlan_(one.ss, ops, { batch: false });
    ctx.applyPlan_(many.ss, ops);
    assert.strictEqual(JSON.stringify(many.log), JSON.stringify(one.log), name + ': 결과가 달라짐');
    assert.ok(many.calls.n < one.calls.n, name + ': 호출이 줄지 않음');
    console.log('#', name, '시트 호출', one.calls.n, '→', many.calls.n);
  }
});

test('SPOT 신규 자재: UOM · Description · Type(열이 있을 때) 필요', () => {
  const values = [['Item', '계약 구분', 'UOM', 'QPL', 'Description', 'Type'], ['A1', 'SPOT', 'EA', '', 'X', 'Chemical']];
  const written = [];
  const sheet = { getDataRange: () => ({ getValues: () => values }), getLastRow: () => values.length, getRange: () => ({ setValues: v => written.push(v), copyTo: () => {}, setNumberFormat: () => {} }) };
  const ctx = inContext({ SpreadsheetApp: { getActive: () => ({ getSheetByName: () => sheet }), CopyPasteType: { PASTE_FORMAT: 'F' } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) }, LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) } });
  const bad = ctx.apiRegisterItems([{ item: 'N1', priceMode: 'SPOT' }]);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(bad.errors[0].messages)), ['SPOT은 UOM이 필요합니다.', 'SPOT은 Description이 필요합니다.', 'SPOT은 Type이 필요합니다.']);
  const ok = ctx.apiRegisterItems([{ item: 'N1', priceMode: 'SPOT', uom: 'RL', desc: 'TAPE', 'x:Type': 'Tape' }]);
  assert.strictEqual(ok.ok, true);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(written[0][0])), ['N1', 'SPOT', 'RL', '', 'TAPE', 'Tape']);
  // 잠금을 오래 못 받으면 한국어 안내
  const busy = inContext({ SpreadsheetApp: { getActive: () => ({ getSheetByName: () => sheet }) }, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    LockService: { getScriptLock: () => ({ waitLock: () => { throw new Error('Lock timeout'); }, releaseLock: () => {} }) } });
  assert.throws(() => busy.apiRegisterItems([{ item: 'N2', priceMode: 'SPOT', uom: 'EA', desc: 'D', 'x:Type': 'T' }]), /다른 분이 지금 저장하는 중/);
});

test('견적요청서 수정: 새 파일(예전 파일은 이름만 바꿈), 견적 옮김, 다음 메일은 [Revised] · 업체 더함', () => {
  const db = fakeBook({ 업체목록: [['업체명', '이메일'], ['Hexcel', 'a@hexcel.com'], ['3M', 'q@3m.com']] });
  let drafts = [], n = 0, clock = '2026-11-16 10:00';
  const rfqBook = id => ({ getUrl: () => 'URL-' + id, getId: () => id, getSheetByName: () => new Proxy({}, { get: () => () => new Proxy({}, { get: (t) => () => t }) }) });
  const files = {}, renamed = [];
  const file = id => files[id] || (files[id] = { getId: () => id, setName: x => renamed.push([id, x]), getParents: () => ({ hasNext: () => true, next: () => 'FOLDER' }),
    makeCopy: (name, folder) => { const nid = 'F' + (++n); files[nid] = file(nid); files[nid].copiedInto = folder; return files[nid]; } });
  const ctx = inContext({
    console: { warn: () => {}, log: () => {} },
    SpreadsheetApp: { getActive: () => db, openById: id => rfqBook(id), CopyPasteType: { PASTE_FORMAT: 'F' }, BorderStyle: { SOLID: 'S', SOLID_MEDIUM: 'M' } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => ({ RFQ_TEMPLATE_ID: 'tpl' })[k] || null }), getUserProperties: () => ({ getProperty: () => '', setProperty: () => {} }) },
    DriveApp: { getFileById: id => file(id) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
    Utilities: { formatDate: (d, tz, f) => (f === 'yyyy-MM-dd HH:mm' ? clock : '261116-100000') },
    UrlFetchApp: { fetch: () => ({ getResponseCode: () => 200, getBlob: () => ({ setName: x => ({ name: x }) }) }) },
    ScriptApp: { getOAuthToken: () => 't' },
    GmailApp: { createDraft: (to, subj, text, opt) => { drafts.push({ subj, opt }); return { getId: () => 'D' + drafts.length }; } },
  });
  const title = '[KE] RFQ // 261116 Test Material Spot (due 11/20)';
  const r = ctx.apiCreateRfq({ baseDate: '2026-11-16', title, due: '2026-11-20', items: [
    { item: 'K1', name: 'A', uom: 'EA', qty: 2, needBy: '2027-01-10' }, { item: 'K2', name: 'B', uom: 'RL', qty: 1, needBy: '2027-02-01' }] });
  clock = '2026-11-16 10:05';
  ctx.apiCreateRfqDraft({ recordId: r.recordId, bcc: ['a@hexcel.com'], vendors: ['Hexcel'] });
  ctx.apiSaveQuotes(r.recordId, { vendors: ['Hexcel'], items: [{ quotes: { Hexcel: { price: '10' } } }, { quotes: { Hexcel: { price: '3' } } }] });
  clock = '2026-11-16 11:00';
  assert.throws(() => ctx.apiReviseRfq({ recordId: r.recordId, items: [{ item: 'K9', qty: 1 }] }), /Item name/);
  const v = ctx.apiReviseRfq({ recordId: r.recordId, baseDate: '2026-11-16', items: [{ item: 'K1', name: 'A', uom: 'EA', qty: 5, needBy: '2027-01-10' }, { item: 'K3', name: 'C', uom: 'KT', qty: 4, needBy: '2027-03-01' }] });
  assert.deepStrictEqual(JSON.parse(JSON.stringify([v.rev.no, v.rev.added, v.rev.removed])), [1, ['K3'], ['K2']]);
  assert.deepStrictEqual(renamed, [['F1', title + ' (수정 전 261116-1100)']]);
  assert.strictEqual(files.F2.copiedInto, 'FOLDER');                      // 같은 폴더에 새 파일
  let rec = ctx.apiMyRecords().records.find(x => x.id === r.recordId);
  assert.deepStrictEqual(JSON.parse(JSON.stringify([rec.url, rec.fileId, rec.status, rec.items.map(i => [i.item, i.qty]), rec.quote.items.map(q => q.quotes.Hexcel ? q.quotes.Hexcel.price : '')])),
    ['URL-F2', 'F2', '견적요청서 수정', [['K1', 5], ['K3', 4]], ['10', '']]);
  clock = '2026-11-16 11:10';
  const d = ctx.apiCreateRfqDraft({ recordId: r.recordId, bcc: ['q@3m.com'], vendors: ['3M'] });
  assert.strictEqual(d.subject, '[Revised] ' + title);
  assert.match(drafts[1].opt.htmlBody, /has been revised/);
  rec = ctx.apiMyRecords().records.find(x => x.id === r.recordId);
  assert.deepStrictEqual(JSON.parse(JSON.stringify([rec.status, rec.vendors, rec.mails.map(m => m.kind)])), ['수정 메일 초안', ['Hexcel', '3M'], ['견적 요청', '수정 요청']]);
  // 그다음 메일은 다시 보통 (수정 뒤 이미 보냄)
  clock = '2026-11-16 11:20';
  assert.strictEqual(ctx.apiCreateRfqDraft({ recordId: r.recordId, bcc: ['q@3m.com'], vendors: ['3M'] }).subject, title);
});

test('자재 정보: 열 키 + Item 있는 줄, 버전은 DB 새로고침 · 자재 등록 때만 바뀜', () => {
  const values = [['', '', ''], ['Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '메모'], ['A1', 'LTA', 'PPG', 1, 'USD', new Date(2026, 0, 1), 'x'], ['', '', '', '', '', '', ''], ['A2', 'SPOT', '', '', '', '', '']];
  const props = { MASTER_VER: '100' };
  const sheet = { getDataRange: () => ({ getValues: () => values }), getLastRow: () => values.length, getRange: () => ({ setValues: () => {}, copyTo: () => {}, setNumberFormat: () => {} }) };
  const ctx = inContext({ SpreadsheetApp: { getActive: () => ({ getSheetByName: () => sheet }), CopyPasteType: { PASTE_FORMAT: 'F' } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) }, Utilities: { formatDate: () => '2026-11-16 09:00' } });
  const m = JSON.parse(JSON.stringify(ctx.apiMasterList()));
  assert.deepStrictEqual(m.cols.map(c => c.key), ['item', 'priceMode', 'supplier', 'price', 'cur', 'start', 'x:메모']);
  assert.deepStrictEqual(m.rows.map(r => r[0]), ['A1', 'A2']);
  assert.strictEqual(m.rows[0][5], '2026-01-01');
  assert.strictEqual(m.ver, '100');
  assert.strictEqual(ctx.apiInit().masterVer, '100');              // 그냥 열면 그대로
  const v1 = ctx.apiInit(true).masterVer;                           // DB 새로고침
  assert.notStrictEqual(v1, '100');
  const reg = ctx.apiRegisterItem({ item: 'A3', priceMode: 'SPOT', uom: 'EA', desc: 'D' });
  assert.ok(reg.masterVer && reg.masterVer === props.MASTER_VER);
});

test('견적요청서 없이 견적 바로 입력: 파일·메일 없는 기록, 업체 고르기, 수정은 품목만, 메일은 막음', () => {
  const db = fakeBook({});
  const ctx = inContext({
    console: { warn: () => {}, log: () => {} },
    SpreadsheetApp: { getActive: () => db },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
    Utilities: { formatDate: (d, tz, f) => (f === 'yyMMdd' ? '261118' : f === 'yyyy-MM-dd HH:mm' ? '2026-11-18 10:00' : '261118-100000') },
  });
  const r = ctx.apiCreateDirectQuote({ items: [{ item: 'KERC02112', name: 'DAPCO 2100', uom: 'CA', qty: 10, needBy: '2026-12-01' }, { item: 'KERC02112', name: 'DAPCO 2100', uom: 'CA', qty: 20, needBy: '2027-01-01' }] });
  assert.strictEqual(r.name, '261118 Material Spot (견적 직접 입력)');
  let rec = ctx.apiMyRecords().records[0];
  assert.deepStrictEqual(JSON.parse(JSON.stringify([rec.kind, rec.direct, rec.url, rec.status, rec.items.length, rec.items[0].parts.length])), ['견적요청서', true, '', '견적 직접 입력', 1, 2]);
  assert.throws(() => ctx.apiSetQuoteVendors(r.recordId, []), /한 곳 이상/);
  ctx.apiSetQuoteVendors(r.recordId, ['Syensqo', ' Syensqo', 'PPG']);
  rec = ctx.apiMyRecords().records[0];
  assert.deepStrictEqual(JSON.parse(JSON.stringify([rec.vendors, rec.status])), [['Syensqo', 'PPG'], '견적 입력']);
  assert.throws(() => ctx.apiCreateRfqDraft({ recordId: r.recordId, bcc: ['a@b.com'] }), /보낼 메일이 없습니다/);
  // Need-by 별로 나눈 견적 저장 → 수정(같은 건수)은 그대로 옮김, 파일 없이
  ctx.apiSaveQuotes(r.recordId, { vendors: ['Syensqo'], items: [{ at: 0, part: 0, quotes: { Syensqo: { price: '5' } } }, { at: 0, part: 1, quotes: { Syensqo: { price: '4.5' } } }] });
  const v = ctx.apiReviseRfq({ recordId: r.recordId, items: [{ item: 'KERC02112', qty: 10, needBy: '2026-12-01' }, { item: 'KERC02112', qty: 30, needBy: '2027-01-01' }] });
  assert.strictEqual(v.url, '');
  rec = ctx.apiMyRecords().records[0];
  assert.deepStrictEqual(JSON.parse(JSON.stringify(rec.quote.items.map(x => [x.at, x.part, x.quotes.Syensqo.price]))), [[0, 0, '5'], [0, 1, '4.5']]);
  assert.strictEqual(rec.items[0].qty, 40);
  assert.strictEqual(ctx.recordAlert_(rec, new Date(2026, 10, 18)).level, '');
});

test('메일 보낸 뒤 견적 입력 업체 더하기 · 빼기: 상태 그대로, 더한 · 뺀 업체', () => {
  const H = ['ID', '종류', '담당자', '작성일', '제목', 'URL', '파일ID', '상태', '초안', 'BCC', 'JSON', '수정일'];
  const db = fakeBook({ 진행기록: [H, ['Q1', '견적요청서', 'me@ke.com', '', 'RFQ', 'u', 'f', '메일 초안', 'D', '', JSON.stringify({ items: [{ item: 'K1' }], vendors: ['OMI', 'CBOL'], mails: [{ no: 1, vendors: ['OMI', 'CBOL'] }] }), '']] });
  const ctx = inContext({ SpreadsheetApp: { getActive: () => db }, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) } });
  const r = ctx.apiSetQuoteVendors('Q1', ['CBOL', '클레비오', '에스피에어로']);
  assert.deepStrictEqual(JSON.parse(JSON.stringify([r.added, r.removed])), [['클레비오', '에스피에어로'], ['OMI']]);
  const rec = ctx.apiMyRecords().records[0];
  assert.deepStrictEqual(JSON.parse(JSON.stringify([rec.status, rec.vendors, rec.mails.length])), ['메일 초안', ['CBOL', '클레비오', '에스피에어로'], 1]);
});

test('환율 탭: 수식(GOOGLEFINANCE) 환율은 덮어쓰지 않고, 손으로 적은 환율만 값·수정일을 바꿈', () => {
  const writes = [];
  const vals = [['CUR', 'KRW 환율', '수정일'], ['USD', 1350, new Date(2026, 10, 23)], ['EUR', 1473.47, '=TODAY()'], ['SGD', 1000, new Date(2026, 10, 1)]];
  const fx = [['', '', ''], ['', '=GOOGLEFINANCE("CURRENCY:USDKRW")', '=TODAY()'], ['', '', '=TODAY()'], ['', '', '']];
  const sh = { getDataRange: () => ({ getValues: () => vals, getFormulas: () => fx }), appendRow: r => writes.push(['append', r[0], r[1]]),
    getRange: (r, c, nr, nc) => ({ setValues: v => writes.push(['set', r, c, nc, v[0][0]]), setValue: v => writes.push(['one', r, c, v]) }) };
  const ctx = inContext({ SpreadsheetApp: { openById: () => ({ getSheetByName: () => sh }) }, PropertiesService: { getScriptProperties: () => ({ getProperty: () => 'X' }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, tryLock: () => true, releaseLock: () => {} }) } });
  ctx.db_ = () => ({ getSheetByName: () => sh });
  ctx.saveRates_({ USD: 1400, EUR: 1500, SGD: 1010, JPY: 9 });
  assert.deepStrictEqual(writes.map(w => w.slice(0, 4)), [['one', 3, 2, 1500], ['set', 4, 2, 2], ['append', 'JPY', 9]]);
  const t = ctx.rateTable_();
  delete t.info.JPY;   // 새로 붙인 줄
  assert.deepStrictEqual(JSON.parse(JSON.stringify(t.info)), { USD: { at: '2026-11-23', auto: true }, EUR: { at: '', auto: false }, SGD: { at: '2026-11-01', auto: false } });
});

test('업체 추가: 같은 업체 줄 아래에 끼워 넣고 위 줄 서식을 복사 · 견적서 파일 붙이기/빼기', () => {
  const log = [];
  const vals = [['업체명', '이메일', 'SUPPLIER', 'SITE', 'CONTACT'], ['CBOL', 'a@cbol.com', 'CBOL', 'CBOL', ''], ['GWT', 'b@gwt.com', 'GWT', 'GWT', '']];
  const sh = { getDataRange: () => ({ getValues: () => vals }), insertRowAfter: r => log.push(['ins', r]), appendRow: () => log.push(['append']),
    getRange: (r, c, nr, nc) => ({ copyTo: (d) => log.push(['fmt', r, d.row]), setValues: v => log.push(['val', r, v[0].join('|')]), row: r }) };
  const ctx = inContext({ SpreadsheetApp: { CopyPasteType: { PASTE_FORMAT: 'F' } }, LockService: { getScriptLock: () => ({ waitLock: () => {}, tryLock: () => true, releaseLock: () => {} }) } });
  ctx.vendorSheet_ = () => sh; ctx.apiVendors = () => ({ ok: 1 });
  ctx.apiAddVendor({ name: 'cbol', email: 'c@cbol.com', contact: 'Jin' });
  assert.deepStrictEqual(log, [['ins', 2], ['fmt', 2, 3], ['val', 3, 'CBOL|c@cbol.com|CBOL|CBOL|Jin']]);

  // 견적서 파일
  const rec = { id: 'Q1', kind: '견적요청서', fileId: 'F1', qfiles: {} }, data = {}, trashed = [];
  Object.assign(ctx, {
    Utilities: { base64Decode: s => Buffer.from(s, 'base64'), formatDate: () => '261123', newBlob: (b, m, n) => ({ n }) },
    DriveApp: { getFileById: id => ({ getParents: () => ({ next: () => ({ createFile: b => ({ getId: () => 'NEW', getUrl: () => 'URL', name: b.n }) }) }), setTrashed: t => trashed.push(id) }) }
  });
  ctx.findRecord_ = () => ({ rec });
  ctx.updateRecord_ = (f, patch) => { if (patch.edit) patch.edit(data); };
  const r = ctx.apiAttachQuoteFile({ recordId: 'Q1', vendor: 'CBOL', name: 'q/1.pdf', mime: 'application/pdf', data: Buffer.from('PDF').toString('base64') });
  assert.strictEqual(r.file.name, 'CBOL_견적서_261123_q_1.pdf');
  assert.deepStrictEqual(Object.keys(data.qfiles), ['CBOL']);
  assert.throws(() => ctx.apiAttachQuoteFile({ recordId: 'Q1', vendor: 'CBOL', name: 'x', data: '' }), /빈 파일/);
  const out = ctx.apiRemoveQuoteFile('Q1', 'CBOL', 'NEW');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(out.qfiles)), {});
  assert.deepStrictEqual(trashed, ['NEW']);
});

test('큰 견적: 진행기록 한 칸(50,000자)을 넘으면 견적데이터 탭에 나눠 담고, 읽을 때 다시 붙임 · 변경기록에 바뀐 단가', () => {
  const H = ['ID', '종류', '담당자', '작성일', '제목', 'URL', '파일ID', '상태', '초안', 'BCC', 'JSON', '수정일'];
  const items = Array.from({ length: 30 }, (_, i) => ({ item: 'KERN' + (1000 + i), name: 'N' + i, qty: 5, needBy: '2027-01-10' }));
  const db = fakeBook({ 진행기록: [H, ['Q1', '견적요청서', 'me@ke.com', '', 'RFQ', '', '', '메일 초안', '', '', JSON.stringify({ items, vendors: [] }), '']] });
  const ctx = inContext({
    SpreadsheetApp: { getActive: () => db }, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
  });
  const vendors = Array.from({ length: 12 }, (_, i) => 'Vendor' + i);
  const mk = p => ({ vendors, items: items.map((it, i) => ({ at: i, program: '787', prev: {}, quotes: Object.fromEntries(vendors.map(v => [v, { price: String(p + i), cur: 'USD', qty: '5', uom: 'EA', mfg: 'MFG-' + v, lt: '8wks', moq: '1', remark: 'remark text for ' + v, shelf: '12M' }])) })) });
  const r1 = ctx.apiSaveQuotes('Q1', Object.assign(mk(10), { base: '' }));
  const row = db.tabs['진행기록'].rows[1];
  assert.ok(row[10].length < 49000, '진행기록 칸 ' + row[10].length);
  assert.strictEqual(JSON.parse(row[10]).quoteRef, 1);
  const parts = db.tabs['견적데이터'].rows.slice(1);
  assert.ok(parts.length >= 2 && parts.every(p => p[0] === 'Q1' && p[2].charAt(0) === '~' && p[2].length <= 45001), '조각 ' + parts.length);
  // 다른 실행에서 읽기 (메모 없이)
  ctx.qstoreMemo_ = null;
  const rec = ctx.findRecord_('Q1', true).rec;
  assert.strictEqual(rec.quote.items.length, 30);
  assert.strictEqual(rec.quote.items[29].quotes.Vendor11.price, '39');
  assert.strictEqual(rec.quote.savedAt, r1.savedAt);
  // 다시 저장: 저장 시각 확인(guard)은 견적데이터의 값으로, 단가 하나 바꿈 → 변경기록
  const q2 = mk(10); q2.items[3].quotes.Vendor2.price = '99'; q2.base = r1.savedAt;
  ctx.qstoreMemo_ = null;
  ctx.apiSaveQuotes('Q1', q2);
  assert.throws(() => ctx.apiSaveQuotes('Q1', Object.assign(mk(10), { base: r1.savedAt })), /다른 창/);
  ctx.qstoreMemo_ = null;
  assert.strictEqual(ctx.findRecord_('Q1', true).rec.quote.items[3].quotes.Vendor2.price, '99');
  const log = db.tabs['변경기록'].rows.slice(1).filter(x => x[2] === '견적 저장');
  assert.strictEqual(log.length, 2);
  assert.ok(/KERN1003 Vendor2: 단가 13→99/.test(log[1][6]), log[1][6]);
  // 바뀐 것이 없는 저장은 변경기록에 남기지 않음
  ctx.qstoreMemo_ = null;
  const now = ctx.findRecord_('Q1', true).rec.quote.savedAt, q3 = mk(10); q3.items[3].quotes.Vendor2.price = '99'; q3.base = now;
  ctx.apiSaveQuotes('Q1', q3);
  assert.strictEqual(db.tabs['변경기록'].rows.slice(1).filter(x => x[2] === '견적 저장').length, 2);
  // 견적 말고도 너무 크면 한국어 안내
  assert.throws(() => ctx.updateRecord_(ctx.findRecord_('Q1', true), { data: { items: Array.from({ length: 900 }, (_, i) => ({ item: 'KERX' + i, name: 'x'.repeat(60) })) } }), /품목이 너무 많아/);
});

test('같은 요청 두 번: 10분 안에 같은 내용이면 다시 보내지 않음 · 보낸 뒤 기록 실패는 오류 대신 안내', () => {
  const H = ['ID', '종류', '담당자', '작성일', '제목', 'URL', '파일ID', '상태', '초안', 'BCC', 'JSON', '수정일'];
  const db = fakeBook({ 진행기록: [H, ['P1', '계약 구매내역서', 'me@ke.com', '', '261113 구매내역서', 'https://docs/p1', '', '작성', '', '', '{"items":[]}', '']] });
  const store = new Map(), sent = [];
  let lockOk = true;
  const ctx = inContext({
    SpreadsheetApp: { getActive: () => db },
    CacheService: { getUserCache: () => ({ get: k => store.get(k) || null, put: (k, v) => store.set(k, v), remove: k => store.delete(k) }), getScriptCache: () => null },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }), getUserProperties: () => ({ getProperty: () => null, setProperties: () => {} }) },
    LockService: { getScriptLock: () => ({ waitLock: () => { if (!lockOk) throw new Error('timeout'); }, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
    Utilities: { formatDate: () => '2026-11-13 09:00', sleep: () => {} },
    GmailApp: { createDraft: () => {}, sendEmail: (to, s) => sent.push(s) },
  });
  const p = { recordId: 'P1', to: 'biz@ke.com', subject: '날인 요청', mode: 'send' };
  const a = ctx.apiConfirmMail(p);
  assert.strictEqual(sent.length, 1);
  assert.ok(!a.repeated && !a.warn);
  const b = ctx.apiConfirmMail(Object.assign({}, p));     // 같은 내용으로 다시 누름
  assert.strictEqual(sent.length, 1);
  assert.strictEqual(b.repeated, true);
  // 내용이 다르면 보냄. 보낸 뒤 기록 잠금을 못 받으면 오류가 아니라 warn (다시 누르면 두 번 나가므로)
  lockOk = false;
  const c = ctx.apiConfirmMail(Object.assign({}, p, { subject: '날인 요청 2' }));
  assert.strictEqual(sent.length, 2);
  assert.ok(/메일은 보냈지만 진행 기록에는 남기지 못했습니다/.test(c.warn), c.warn);
  assert.ok(!/아무것도 바뀌지 않았습니다/.test(c.warn));
  // 보내기 전에 실패하면 캐시에 남지 않아 고쳐서 다시 누를 수 있음
  lockOk = true;
  assert.throws(() => ctx.apiConfirmMail({ recordId: 'P1', to: 'nope', mode: 'send' }), /이메일 형식/);
  assert.throws(() => ctx.apiConfirmMail({ recordId: 'P1', to: 'nope', mode: 'send' }), /이메일 형식/);
  // 처리하는 중에 또 누르면 안내
  const key = [...store.keys()][0];
  store.set(key, JSON.stringify({ state: 'run' }));
  assert.throws(() => ctx.apiConfirmMail(p), /지금 처리하고 있습니다/);
});

test('견적데이터 기록: 잠금 안에서는 시트를 새로 읽음 · 못 읽으면 지우지 않고 멈춤 · 수정 중 들어온 견적도 옮김 · 같은 요청은 기록이 바뀌었으면 다시', () => {
  const H = ['ID', '종류', '담당자', '작성일', '제목', 'URL', '파일ID', '상태', '초안', 'BCC', 'JSON', '수정일'];
  const items = Array.from({ length: 30 }, (_, i) => ({ item: 'KERN' + (1000 + i), name: 'N' + i, qty: 5, needBy: '2027-01-10' }));
  const db = fakeBook({ 진행기록: [H, ['Q1', '견적요청서', 'me@ke.com', '', 'RFQ', '', '', '메일 초안', '', '', JSON.stringify({ items, vendors: ['Vendor0'], direct: true }), '']] });
  const store = new Map();
  const ctx = inContext({
    SpreadsheetApp: { getActive: () => db }, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    CacheService: { getUserCache: () => ({ get: k => store.get(k) || null, put: (k, v) => store.set(k, v), remove: k => store.delete(k) }), getScriptCache: () => null },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) },
    Utilities: { formatDate: () => '2026-11-24 10:00', sleep: () => {} },
  });
  const vendors = Array.from({ length: 12 }, (_, i) => 'Vendor' + i);
  const mk = p => ({ vendors, items: items.map((it, i) => ({ at: i, quotes: Object.fromEntries(vendors.map(v => [v, { price: String(p), cur: 'USD', qty: '5', mfg: 'M-' + v, remark: '=비고 ' + 'r'.repeat(60) }])) })) });
  const s1 = ctx.apiSaveQuotes('Q1', Object.assign(mk(10), { base: '' }));
  assert.strictEqual(JSON.parse(db.tabs['진행기록'].rows[1][10]).quoteRef, 1);
  // 1) 잠금 밖에서 읽어 둔 뒤 다른 창이 단가 500 저장 → 견적과 상관없는 수정이 옛 견적으로 덮어쓰지 않음
  ctx.qstoreMemo_ = null;
  const found = ctx.findRecord_('Q1', true);                 // 메모에 단가 10
  const other = inContext({ SpreadsheetApp: { getActive: () => db }, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@ke.com' }), getEffectiveUser: () => ({ getEmail: () => 'me@ke.com' }) } });
  other.apiSaveQuotes('Q1', Object.assign(mk(500), { base: s1.savedAt }));
  ctx.updateRecord_(found, { data: { quoteSheet: 'X' } });
  ctx.qstoreMemo_ = null;
  assert.strictEqual(ctx.findRecord_('Q1', true).rec.quote.items[0].quotes.Vendor0.price, '500');
  assert.ok(db.tabs['견적데이터'].rows.slice(1).every(r => r[2].charAt(0) === '~'));   // = 로 시작하는 조각도 글자로
  // 2) 수정 중 다른 창 견적도 같이 옮김 (잠금 안의 값으로)
  const rows = items.map(it => ({ item: it.item, name: it.name, qty: 5, needBy: '2027-01-10' }));
  const rv = ctx.apiReviseRfq({ recordId: 'Q1', items: rows });
  ctx.qstoreMemo_ = null;
  assert.strictEqual(ctx.findRecord_('Q1', true).rec.quote.items[0].quotes.Vendor0.price, '500');
  assert.strictEqual(rv.rev.no, 1);
  // 같은 내용으로 다시 수정: 기록이 그대로면 repeated, 그 사이 기록이 바뀌었으면 다시 함
  assert.strictEqual(ctx.apiReviseRfq({ recordId: 'Q1', items: rows }).repeated, true);
  db.tabs['진행기록'].rows[1][11] = new Date(Date.now() + 5000);
  const rv2 = ctx.apiReviseRfq({ recordId: 'Q1', items: rows });
  assert.ok(!rv2.repeated && rv2.rev.no === 2);
  // 3) 견적데이터 조각이 깨지면 저장을 멈추고 줄을 지우지 않음
  const qd = db.tabs['견적데이터'].rows;
  qd[1][2] = '#ERROR!';
  const n = qd.length;
  ctx.qstoreMemo_ = null;
  assert.throws(() => ctx.apiSetQuoteVendors('Q1', ['Vendor0', 'Vendor1']), /견적데이터 탭\)을 읽지 못해/);
  assert.strictEqual(db.tabs['견적데이터'].rows.length, n);
  assert.strictEqual(JSON.parse(db.tabs['진행기록'].rows[1][10]).quoteRef, 1);
});
