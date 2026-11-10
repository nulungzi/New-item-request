// dist/ 배포 파일 검사: 최신인지, Apps Script처럼 한 파일로 읽혀도 동작하는지
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs'), vm = require('vm'), path = require('path');
const { bundle } = require('../tools/bundle.js');
const dist = path.join(__dirname, '..', 'dist');

test('dist 가 원본과 같다 (npm run build 를 잊지 않았는지)', () => {
  const fresh = bundle();
  for (const name of Object.keys(fresh)) {
    assert.strictEqual(fs.readFileSync(path.join(dist, name), 'utf8'), fresh[name], name + ' 이 오래되었습니다. npm run build 를 실행하세요.');
  }
});

test('dist/Code.gs: 첫 글자, 금지 문자, 문법', () => {
  const src = fs.readFileSync(path.join(dist, 'Code.gs'), 'utf8');
  assert.ok(src.startsWith('// ====='), '첫 줄은 주석이어야 합니다');
  const bad = [...src].filter(ch => { const c = ch.codePointAt(0); return (c < 32 && c !== 9 && c !== 10 && c !== 13) || [0xa0, 0x2028, 0x2029, 0xfeff].includes(c); });
  assert.deepStrictEqual(bad, [], '보이지 않는 문자가 있습니다');
  assert.ok(!/module\.exports/.test(src), 'Node 전용 코드가 남아 있습니다');
  assert.doesNotThrow(() => new vm.Script(src, { filename: 'Code.gs' }));
});

test('dist/Code.gs 단독으로 apiInit · checkSetup 이 동작', () => {
  const values = [
    ['', 'Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료', 'Shelf Life', 'Description'],
    [1, 'A1', 'LTA', 'PPG', 1, 'USD', new Date(2026, 0, 1), new Date(2028, 11, 31), '12MM', 'PRIMER'],
    [2, 'A1', 'SPOT', 'ACME', '', '', '', '', '', ''],
  ];
  const sheet = { getDataRange: () => ({ getValues: () => values }), getRange: () => ({ setValues: () => {} }), appendRow: () => {} };
  const tplSheets = ['구매요약', '구매세부내역', '구매일정', '사업부 담당 확인', '_proto'];
  const logs = [];
  const ctx = {
    console: { log: m => logs.push(m), warn: () => {} },
    SpreadsheetApp: {
      getActive: () => ({ getSheetByName: n => (n === '자재정보' || n === '환율' ? sheet : null), insertSheet: () => sheet }),
      openById: () => ({ getName: () => '구매내역서 템플릿', getSheetByName: n => (tplSheets.includes(n) ? {} : null) }),
    },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k === 'TEMPLATE_ID' ? 'tpl' : null) }) },
    DriveApp: {},
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(dist, 'Code.gs'), 'utf8'), ctx);
  const init = ctx.apiInit();
  assert.deepStrictEqual(JSON.parse(JSON.stringify(init.items)), [['A1', 'PRIMER', '', 'LTA/SPOT', '12MM']]);
  const report = ctx.checkSetup();
  assert.match(report, /✅ 자재정보: 머리글 1행, 품목 2줄/);
  assert.match(report, /✅ 템플릿: 구매내역서 템플릿/);
  assert.ok(!/❌/.test(report), report);
});
