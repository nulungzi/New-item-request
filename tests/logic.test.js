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

test('품목별 Need-by-date로 계약/Spot과 연도별 단가를 판정한다', () => {
  const today = L.toDate_('2026-10-07');
  const q = L.parseRequest_('KERC00067 1 2027-03-01\nKERC00016 1 2026-12-15\nKERC00016 1 2027-02-01\nKERN00089-001 1 2026-11-30\nKERN00089-001 1 2027-05-01\nNOPE 1 2026-12-01\nKERC00999 1 2026-12-01').requests;
  const r = L.classify_(q, contracts, today);
  assert.deepStrictEqual(r.map(x => x.status), ['계약', '계약', 'Spot', '계약', '계약', 'Spot', '계약']);
  assert.match(r[2].reason, /Need-by 시점 계약 만료 \(PPG, 2026-12-31 종료\)/);
  assert.strictEqual(r[3].candidates[0].price, 105);
  assert.strictEqual(r[4].candidates[0].price, 130);
  assert.strictEqual(r[5].reason, '계약 없음');
  assert.deepStrictEqual(r[6].candidates.map(c => c.supplier), ['Mapaero', 'AKZO']);
});

test('Need-by-date가 없으면 작성일로 판정하고 경고', () => {
  const r = L.classify_(L.parseRequest_('KERC00016 1').requests, contracts, L.toDate_('2026-10-07'));
  assert.strictEqual(r[0].status, '계약');
  assert.match(r[0].warnings[0], /Need-by-date 없음/);
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

test('리드타임 기준 Need-by 경고', () => {
  assert.strictEqual(L.leadDays_('10-12wks'), 84);
  assert.strictEqual(L.leadDays_('30 days'), 30);
  assert.strictEqual(L.leadDays_(''), null);
  const c = L.parseContracts_([
    ['Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료', 'Leadtime'],
    ['A1', 'LTA', 'PPG', 1, 'USD', '2025-01-01', '2028-12-31', '10-12wks'],
  ]);
  const r = L.classify_(L.parseRequest_('A1 1 2026-11-01\nA1 1 2027-02-01\nA1 1 2026-10-01').requests, c, L.toDate_('2026-10-07'));
  assert.deepStrictEqual(r.map(x => x.warnings), [['리드타임(10-12wks) 부족'], [], ['Need-by가 작성일보다 이전']]);
});

test('Program·Ship Form·Lead Time 열 이름 인식', () => {
  const c = L.parseContracts_([
    ['Item', '업체명', 'Program', 'Ship Form', 'Lead Time', 'Material'],
    ['A1', 'PPG', '787AB / 737WL', 'Ocean', '10-12wks', 'Primer'],
  ])[0];
  assert.deepStrictEqual([c.program, c.shipFrom, c.leadtime, c.material], ['787AB / 737WL', 'Ocean', '10-12wks', 'Primer']);
});

test('공백이 들어간 품번 (탭·공백 붙여넣기 모두)', () => {
  const p = L.parseRequest_('10-60754-279 42L\t50\t2027-06-10\n10-60754-279  42L 50 2027-06-10 787AB\nKERC00053 50 2027-06-10');
  assert.deepStrictEqual(p.requests.map(r => [r.item, r.qty, L.fmtDate_(r.needBy), r.program]),
    [['10-60754-279 42L', 50, '2027-06-10', ''], ['10-60754-279 42L', 50, '2027-06-10', '787AB'], ['KERC00053', 50, '2027-06-10', '']]);
});

test('MOQ 미달 경고', () => {
  const c = L.parseContracts_([['Item', '업체명', '계약시작', '계약종료', 'MOQ'], ['A1', 'S', '2026-01-01', '2028-12-31', '6,000'], ['B1', 'S', '2026-01-01', '2028-12-31', '-']]);
  const r = L.classify_(L.parseRequest_('A1 50 2027-06-10\nA1 6000 2027-06-10\nB1 1 2027-06-10').requests, c, L.toDate_('2026-11-09'));
  assert.deepStrictEqual(r.map(x => x.warnings), [['수량 50 < MOQ 6000'], [], []]);
});

test('구매내역서 계획: 통화 2개, 업체 2개', () => {
  const base = L.toDate_('2026-11-10');
  const lines = [
    { item: 'A', qty: 10, needBy: L.toDate_('2027-06-10'), supplier: 'AKZO', cur: 'USD', price: 2, priceMode: 'LTA', material: 'Primer', shipFrom: 'Air', leadtime: '10-12wks' },
    { item: 'B', qty: 5, needBy: L.toDate_('2027-06-10'), supplier: 'Mapaero', cur: 'EUR', price: 3, priceMode: 'LTA', material: 'Base', shipFrom: '', leadtime: '' },
  ];
  const ops = L.buildPurchasePlan_(lines, ['USD', 'EUR'], { USD: 1350, EUR: 1500 }, base);
  const get = (s, r, c) => (ops.filter(o => o.op === 'set' && o.s === s && o.r === r && o.c === c).pop() || {}).v;
  // 세부내역: 통화 2개 → 머리글 8행, 데이터 9~10행
  assert.strictEqual(get('구매세부내역', 8, 2), 'L/I');
  assert.strictEqual(get('구매세부내역', 9, 4), 'A');
  assert.strictEqual(get('구매세부내역', 5, 11), '=SUMIF($I$9:$I$10,"KRW",$K$9:$K$10)');
  assert.strictEqual(get('구매세부내역', 6, 11), '=K3*$P$2+K4*$Q$2+K5');
  assert.deepStrictEqual(get('구매세부내역', 9, 20), { date: '2027-06-10' });
  // 비어 있는 Ship From·Leadtime·Program은 노란 칸
  const bg = ops.filter(o => o.op === 'bg' && o.s === '구매세부내역').map(o => o.r + ',' + o.c);
  assert.deepStrictEqual(bg, ['9,3', '10,3', '10,12', '10,13']);
  // 요약: 업체 2줄, CUR 합계 2줄, 총합계
  assert.strictEqual(get('구매요약', 3, 8), '구매금액');
  assert.strictEqual(get('구매요약', 6, 2), 'CUR 별 합계');
  assert.strictEqual(get('구매요약', 8, 8), "=H6*'구매세부내역'!$P$2+H7*'구매세부내역'!$Q$2");
  // 구매일정: 11월부터 11개월, 2026년(2개월) / 2027년(9개월)
  assert.strictEqual(get('구매일정', 5, 5), 11);
  assert.strictEqual(get('구매일정', 5, 25), 9);
  const merges = ops.filter(o => o.op === 'merge' && o.s === '구매일정').map(o => [o.c1, o.c2]);
  assert.deepStrictEqual(merges, [[5, 8], [9, 26]]);
});
