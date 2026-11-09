// node --test tests/
const test = require('node:test');
const assert = require('node:assert');
const L = require('../apps-script/Logic.js');

const sheet = [
  [null, null, null, null, null, null, null, null],
  [null, 'Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료', 'MOQ', 'UOM', 'Description', 'Product\n Name'],
  [null, 'KERC00067', 'LTA', 'AKZO', 1, 'USD', '2026-01-01', '2028-12-31', '없음', 'KT', 'TOP COAT', 'A'],
  [null, 'KERC00016', 'LTA', 'PPG', 68, 'USD', '2025-01-01', '2026-12-31', '없음', 'KT', 'PRIMER', ''],
  [null, 'KERN00089-001', '3PP', 'Syensqo', 105, 'USD', '2026-01-01', '2026-12-31', 6000, 'FT2', 'FILM', ''],
  [null, 'KERN00089-001', '3PP', 'Syensqo', 130, 'USD', '2027-01-01', '2027-12-31', 6000, 'FT2', 'FILM', ''],
  [null, 'KERC00999', 'LTA', 'AKZO', 9, 'EUR', '2026-01-01', '2028-12-31', '없음', 'LI', 'X', ''],
  [null, 'KERC00999', 'LTA', 'Mapaero', 7, 'EUR', '2026-01-01', '2028-12-31', '없음', 'LI', 'X', ''],
];
const contracts = L.parseContracts_(sheet);

test('머리글 위치·줄바꿈 머리글을 인식한다', () => {
  assert.strictEqual(contracts.length, 6);
  assert.strictEqual(contracts[0].product, 'A');
  assert.strictEqual(contracts[0].row, 3);
});

test('소요제기 붙여넣기 파싱 (탭·공백, 머리글 건너뜀)', () => {
  const p = L.parseRequest_('품번\t수량\nkerc00067\t1,200\t787AB / 737WL\nKERC00016 10\n\nKERX abc');
  assert.deepStrictEqual(p.requests.map(r => [r.item, r.qty, r.program]),
    [['KERC00067', 1200, '787AB / 737WL'], ['KERC00016', 10, '']]);
  assert.deepStrictEqual(p.skipped, [1, 5]);
});

test('기준일로 계약/Spot과 연도별 단가를 판정한다', () => {
  const q = L.parseRequest_('KERC00067 1\nKERC00016 1\nKERN00089-001 1\nNOPE 1\nKERC00999 1').requests;
  const r26 = L.classify_(q, contracts, L.toDate_('2026-10-07'));
  assert.deepStrictEqual(r26.map(r => r.status), ['계약', '계약', '계약', 'Spot', '계약']);
  assert.strictEqual(r26[2].candidates[0].price, 105);
  assert.strictEqual(r26[3].reason, '계약 없음');
  assert.deepStrictEqual(r26[4].candidates.map(c => c.supplier), ['Mapaero', 'AKZO']); // 단가 낮은 순
  const r27 = L.classify_(q, contracts, L.toDate_('2027-02-01'));
  assert.strictEqual(r27[1].status, 'Spot');
  assert.match(r27[1].reason, /계약 만료 \(PPG, 2026-12-31 종료\)/);
  assert.strictEqual(r27[2].candidates[0].price, 130);
});

test('업체·통화별 요약', () => {
  const g = L.summarize_([
    { supplier: 'PPG', cur: 'USD', priceMode: 'LTA', qty: 300, price: 1, material: 'Primer' },
    { supplier: 'PPG', cur: 'USD', priceMode: 'LTA', qty: 100, price: 2, material: 'Sealant' },
    { supplier: 'PPG', cur: 'USD', priceMode: 'LTA', qty: 10, price: 3, material: 'Sealant' },
  ]);
  assert.strictEqual(g.length, 1);
  assert.deepStrictEqual([g[0].count, g[0].qty, g[0].amount, g[0].mainMaterial], [3, 410, 530, 'Sealant']);
});

test('Need-by-date 여러 표기 인식, Program은 나머지 칸', () => {
  const p = L.parseRequest_('KERC00067\t5\t2027-01-15\nKERC00016\t3\t2027. 2. 1.\t787AB\nKERC00016 3 12/20/2026 737WL FSF\nKERC00016 1');
  assert.deepStrictEqual(p.requests.map(r => [L.fmtDate_(r.needBy), r.program]),
    [['2027-01-15', ''], ['2027-02-01', '787AB'], ['2026-12-20', '737WL FSF'], ['', '']]);
});

test('리드타임·계약종료 기준 Need-by 경고', () => {
  assert.strictEqual(L.leadDays_('10-12wks'), 84);
  assert.strictEqual(L.leadDays_('30 days'), 30);
  assert.strictEqual(L.leadDays_(''), null);
  const c = L.parseContracts_([
    ['Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료', 'Leadtime'],
    ['A1', 'LTA', 'PPG', 1, 'USD', '2025-01-01', '2026-12-31', '10-12wks'],
  ]);
  const base = L.toDate_('2026-10-07');
  const r = L.classify_(L.parseRequest_('A1 1 2026-11-01\nA1 1 2027-02-01\nA1 1 2026-12-31').requests, c, base);
  assert.deepStrictEqual(r[0].warnings, ['리드타임(10-12wks) 부족']);
  assert.deepStrictEqual(r[1].warnings, ['Need-by가 계약종료(2026-12-31) 이후']);
  assert.deepStrictEqual(r[2].warnings, []);
});
