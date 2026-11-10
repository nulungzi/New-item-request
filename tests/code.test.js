// Code.js 가 Apps Script 없이도 끝까지 실행되는지 가짜 객체로 확인
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs'), vm = require('vm'), path = require('path');

test('apiCreatePurchase: 템플릿 복사 후 계획 적용', () => {
  const calls = [];
  const range = new Proxy({}, { get: (t, k) => (...a) => { calls.push(k); return range; } });
  const sheet = name => ({ getRange: () => range, setRowHeight: () => {}, setFrozenRows: () => {}, setFrozenColumns: () => {},
    getDataRange: () => ({ getValues: () => name === '자재정보'
      ? [['Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료'], ['A1', 'LTA', 'AKZO', 2, 'USD', new Date(2026, 0, 1), new Date(2028, 11, 31)]]
      : [['CUR', 'KRW 환율', '수정일']] }), appendRow: () => {} });
  const ss = { getSheetByName: n => sheet(n), deleteSheet: () => calls.push('deleteSheet'), setActiveSheet: () => {}, getUrl: () => 'URL', getId: () => 'new' };
  const ctx = {
    SpreadsheetApp: { getActive: () => ss, openById: () => ss, CopyPasteType: { PASTE_FORMAT: 'F' }, BorderStyle: { SOLID: 'S' },
      newDataValidation: () => range },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'TEMPLATE_ID' ? 'tpl' : null) }) },
    DriveApp: { getFileById: () => ({ makeCopy: () => ({ getId: () => 'new' }) }) },
  };
  vm.createContext(ctx);
  for (const f of ['Logic.js', 'Code.js']) vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'apps-script', f), 'utf8'), ctx);
  const res = ctx.apiCreatePurchase({ baseDate: '2026-11-10', rates: { USD: 1350 }, remarks: {}, lines: [{ item: 'A1', qty: 3, needBy: '2027-06-10', contractRow: 2 }] });
  assert.strictEqual(res.name, '261110 구매내역서_AKZO');
  assert.strictEqual(res.id, 'new');
  assert.ok(calls.includes('copyTo') && calls.includes('merge') && calls.includes('deleteSheet'));
});

test('파일 읽는 순서와 상관없이 로드된다 (Code.js 먼저)', () => {
  for (const order of [['Code.js', 'Logic.js'], ['Logic.js', 'Code.js'], ['Code.js']]) {
    const ctx = {};
    vm.createContext(ctx);
    assert.doesNotThrow(() => {
      for (const f of order) vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'apps-script', f), 'utf8'), ctx);
    }, order.join(' → '));
  }
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
  vm.createContext(ctx);
  for (const f of ['Logic.js', 'Code.js']) vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'apps-script', f), 'utf8'), ctx);
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
