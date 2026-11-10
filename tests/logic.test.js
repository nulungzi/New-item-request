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
  assert.strictEqual(r[5].reason, '계약 없음 · DB 미등록 자재');
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
  // 구매일정: 11월부터 12개월, 2026년(2개월) / 2027년(10개월)
  assert.strictEqual(get('구매일정', 5, 5), 11);
  assert.strictEqual(get('구매일정', 5, 27), 10);
  const merges = ops.filter(o => o.op === 'merge' && o.s === '구매일정').map(o => [o.c1, o.c2]);
  assert.deepStrictEqual(merges, [[5, 8], [9, 28]]);
});

test('화면 표 입력: Program·Ship·Tax 전달, 빈 행 무시', () => {
  const p = L.parseRows_([
    { item: ' kerc00053 ', qty: '50', needBy: '2027-06-10', program: '787AB', shipFrom: 'Ocean', taxCode: '1' },
    { item: '', qty: '' },
    { item: 'X', qty: '' },
  ]);
  assert.deepStrictEqual(p.skipped, [3]);
  const r = L.classify_(p.requests, contracts, L.toDate_('2026-11-10'))[0];
  assert.deepStrictEqual([r.item, r.qty, L.fmtDate_(r.needBy), r.program, r.shipFrom, r.taxCode], ['KERC00053', 50, '2027-06-10', '787AB', 'Ocean', '1']);
});

test('구매일정 막대: 예시·사용자 값·빈 줄·범위 보정', () => {
  assert.deepStrictEqual(L.normSchedule_(null), L.SCHEDULE_EXAMPLE);
  assert.deepStrictEqual(L.normSchedule_([{ start: 2, len: 5 }, { start: 20, len: 9 }, { start: -3, len: 2 }, { start: 4, len: 0 }]),
    [{ start: 2, len: 5 }, { start: 20, len: 4 }, { start: 0, len: 2 }, { start: 0, len: 0 }, { start: 0, len: 0 }]);
  const lines = [{ item: 'A', qty: 1, supplier: 'S', cur: 'USD', price: 1, priceMode: 'LTA', shipFrom: 'Air', leadtime: '1wks' }];
  const ops = L.buildPurchasePlan_(lines, ['USD'], { USD: 1 }, L.toDate_('2026-11-10'), [{ start: 2, len: 3 }, { start: 0, len: 0 }]);
  const row6 = ops.filter(o => o.op === 'bg' && o.s === '구매일정' && o.r === 6);
  assert.strictEqual(row6.length, 24);
  assert.deepStrictEqual(row6.filter(o => o.color !== '#ffffff').map(o => o.c), [7, 8, 9]);
  assert.strictEqual(ops.filter(o => o.op === 'bg' && o.s === '구매일정' && o.r === 7 && o.color !== '#ffffff').length, 0);
  const lt = ops.filter(o => o.op === 'set' && o.s === '구매일정' && o.c === 4).map(o => o.v);
  assert.deepStrictEqual(lt, ['1.5M', '', '', '', '']);
});

test('SPOT 등록 행은 계약으로 보지 않는다', () => {
  const c = L.parseContracts_([
    ['Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료', 'Description'],
    ['S1', 'SPOT', 'ACME', '', '', '', '', 'SPOT ONLY'],
    ['S2', 'SPOT', 'ACME', '', '', '', '', 'X'],
    ['S2', 'LTA', 'PPG', 5, 'USD', '2026-01-01', '2028-12-31', 'X'],
  ]);
  const r = L.classify_(L.parseRequest_('S1 1 2027-01-01\nS2 1 2027-01-01\nNEW 1 2027-01-01').requests, c, L.toDate_('2026-11-10'));
  assert.deepStrictEqual(r.map(x => [x.status, x.registered]), [['Spot', true], ['계약', true], ['Spot', false]]);
  assert.strictEqual(r[0].reason, 'Spot 자재 (DB 등록: ACME)');
  assert.strictEqual(r[0].desc, 'SPOT ONLY');
  assert.deepStrictEqual(r[1].candidates.map(x => x.supplier), ['PPG']);
});

test('신규 자재 등록 검사와 행 만들기', () => {
  const c = L.parseContracts_([
    ['Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료'],
    ['A1', 'LTA', 'PPG', 1, 'USD', '2026-01-01', '2026-12-31'],
  ]);
  // SPOT: Item·구분만 있어도 됨
  let v = L.validateNewItem_({ item: ' new-1 42l ', priceMode: 'spot', desc: 'SEAL' }, c);
  assert.deepStrictEqual(v.errors, []);
  assert.strictEqual(v.rec.item, 'NEW-1 42L');
  // LTA: 필수값 누락
  v = L.validateNewItem_({ item: 'B1', priceMode: 'LTA' }, c);
  assert.strictEqual(v.errors.length, 4);
  // LTA: 기간 겹침
  v = L.validateNewItem_({ item: 'A1', priceMode: 'LTA', supplier: 'ppg', price: '2', cur: 'usd', start: '2026-06-01', end: '2027-05-31' }, c);
  assert.match(v.errors[0], /겹치는 계약.*2행/);
  // 다음 해 계약은 가능, 값 정리
  v = L.validateNewItem_({ item: 'A1', priceMode: 'LTA', supplier: 'PPG', price: '1,234.5', cur: 'usd', start: '2027-01-01', end: '2027-12-31', moq: '6,000', 'x:비고2': ' memo ' }, c);
  assert.deepStrictEqual(v.errors, []);
  assert.deepStrictEqual([v.rec.price, v.rec.cur, v.rec.moq, L.fmtDate_(v.rec.start)], [1234.5, 'USD', 6000, '2027-01-01']);
  const headers = ['', 'Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료', 'MOQ', 'Lead Time', 'Ship Form', '비고2'];
  const row = L.newItemRow_(headers, v.rec);
  assert.deepStrictEqual(row.slice(0, 6), ['', 'A1', 'LTA', 'PPG', 1234.5, 'USD']);
  assert.strictEqual(row[11], 'memo');
  assert.deepStrictEqual(headers.map(L.keyForHeader_).slice(1), ['item', 'priceMode', 'supplier', 'price', 'cur', 'start', 'end', 'moq', 'leadtime', 'shipFrom', 'x:비고2']);
});

test('구매내역서 파일 이름: 작업자 입력 우선, 비우면 자동', () => {
  const base = L.toDate_('2026-11-10');
  assert.strictEqual(L.purchaseFileName_('', base, ['Syensqo', 'AKZO']), '261110 구매내역서_Syensqo_AKZO');
  assert.strictEqual(L.purchaseFileName_('  보잉 화공자재\n LTA 품의  ', base, ['AKZO']), '보잉 화공자재 LTA 품의');
  assert.strictEqual(L.purchaseFileName_('x'.repeat(200), base, []).length, 150);
});

test('사업부 담당 확인: 담당자·제목, 확인 날짜는 비워 둠', () => {
  const lines = [{ item: 'A', qty: 1, supplier: 'S', cur: 'USD', price: 1, priceMode: 'LTA', shipFrom: 'Air', leadtime: '1wks', shelfLife: '12MM' }];
  const ops = L.buildPurchasePlan_(lines, ['USD'], { USD: 1 }, L.toDate_('2026-11-10'), null, { title: '보잉 LTA 품의', confirmer: ' 김정환   차장 ' });
  const c = ops.filter(o => o.s === '사업부 담당 확인');
  assert.deepStrictEqual(c.filter(o => o.op === 'set').map(o => [o.r, o.c, o.v]), [[2, 1, '김정환 차장'], [1, 2, '구매내역서 제목'], [2, 2, '보잉 LTA 품의']]);
  assert.deepStrictEqual(c.filter(o => o.op === 'bg').map(o => o.c), [3]);
  // Shelf Life는 구매세부내역 O열(15)
  assert.ok(ops.some(o => o.op === 'set' && o.s === '구매세부내역' && o.c === 15 && o.v === '12MM'));
  // 담당자 비우면 A2도 노란 칸
  const ops2 = L.buildPurchasePlan_(lines, ['USD'], { USD: 1 }, L.toDate_('2026-11-10'), null, { title: 't' });
  assert.deepStrictEqual(ops2.filter(o => o.s === '사업부 담당 확인' && o.op === 'bg').map(o => o.c), [1, 3]);
  assert.ok(L.isPersonFormat_('김정환 차장') && L.isPersonFormat_('남궁민수 대리') && !L.isPersonFormat_('김정환') && !L.isPersonFormat_('차장 김정환 님'));
});

test('uniq_ · itemIndex_', () => {
  assert.deepStrictEqual(L.uniq_(['b', 'a', 'b', '', null, 'a', 'c']), ['b', 'a', 'c']);
  const c = L.parseContracts_([
    ['Item', '계약 구분', '업체명', 'Ship Form', 'Shelf Life', 'Description'],
    ['B2', 'spot', 'X', '', '', ''], ['A1', 'LTA', 'PPG', 'Air', '', 'PRIMER'], ['A1', '3PP', 'Q', 'Ocean', '6+3(DOP)', 'P2'], ['B2', 'LTA', 'Y', '', '', 'SEAL'],
  ]);
  assert.deepStrictEqual(L.itemIndex_(c), [['A1', 'PRIMER', 'Air', 'LTA/3PP', '6+3(DOP)'], ['B2', 'SEAL', '', 'SPOT/LTA', '']]);
});
