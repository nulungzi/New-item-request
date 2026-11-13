// node --test tests/
const test = require('node:test');
const assert = require('node:assert');
const L = require('./load.js').load();

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
  assert.deepStrictEqual(r.map(x => x.warnings), [['수량 50 → MOQ(최소 6000) · 6000개로 주문'], [], []]);
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
  const merges = ops.filter(o => o.op === 'merge' && o.s === '구매일정');
  assert.deepStrictEqual(merges.filter(o => o.r1 === 4).map(o => [o.c1, o.c2]), [[5, 8], [9, 28]]);   // 연도
  assert.strictEqual(merges.filter(o => o.r1 === 5).length, 12);                                        // 달마다 2칸
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
  assert.strictEqual(r[0].reason, 'Spot 자재 · DB 단가 없음');
  assert.deepStrictEqual(r.map(x => x.route), ['rfq', 'contract', 'rfq']);
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

test('구매일정 E~AB열 너비 23px', () => {
  const lines = [{ item: 'A', qty: 1, supplier: 'S', cur: 'USD', price: 1, priceMode: 'LTA' }];
  const w = L.buildPurchasePlan_(lines, ['USD'], { USD: 1 }, L.toDate_('2026-11-10')).filter(o => o.op === 'width' && o.s === '구매일정');
  assert.deepStrictEqual(w.map(o => o.c), Array.from({ length: 24 }, (_, i) => 5 + i));
  assert.ok(w.every(o => o.s === '구매일정' && o.px === 23));
});

test('드라이브 폴더 주소 → ID', () => {
  const id = '1AbCdEfGhIjKlMnOpQrStUvWxYz012345';
  assert.strictEqual(L.parseFolderId_('https://drive.google.com/drive/folders/' + id + '?usp=sharing'), id);
  assert.strictEqual(L.parseFolderId_('https://drive.google.com/drive/u/1/folders/' + id), id);
  assert.strictEqual(L.parseFolderId_('https://drive.google.com/open?id=' + id), id);
  assert.strictEqual(L.parseFolderId_('  ' + id + ' '), id);
  assert.strictEqual(L.parseFolderId_(''), '');
  assert.strictEqual(L.parseFolderId_('내 폴더'), '');
});

test('입력용 시트: 열 순서가 웹앱 표와 같고, 머리글 줄을 알아본다', () => {
  assert.deepStrictEqual(L.INPUT_SHEET.COLUMNS.map(c => c[0]), ['Item', '수량', 'Need-by', 'Program', 'Shelf Life', 'Ship', 'Tax']);
  assert.ok(L.isInputHeader_(['Item', '수량']) && L.isInputHeader_([' item ']) && !L.isInputHeader_(['KERC00053', '50']));
});

test('Spot 세분화: DB 단가 있는 Spot은 spotPriced, 없으면 rfq (Airtech 형식)', () => {
  const c = L.parseContracts_([
    ['', 'Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료', 'MOQ', 'UOM', 'Lead Time', 'Ship Form', 'Material', 'Shelf-Life', 'QPL', 'Description'],
    [1, 'KERN00286-003', 'Spot', 'Airtech', 462, 'USD', '2025-01-01', '2027-12-31', 'Min 28, Multi 2', 'RL', '6~8wks', '', 'Nylon Film', '12MM', 'Airtech', 'NYLON FILM 2MIL'],
    [2, 'KERN00286-004', 'Spot', 'Airtech', 463, 'USD', '2025-01-01', '2027-12-31', '-', 'RL', '1~2wks', '', 'Nylon Film', '', 'Airtech', 'NYLON FILM 2MIL 54'],
    [3, 'MS20995C20', 'Spot', '', '', '', '', '', '', 'RL', '', '', 'Safety Wire', '', 'MALIN Co / GERARD DANIEL', 'Safety Wire'],
  ]);
  assert.deepStrictEqual([c[0].shelfLife, c[0].qpl, c[2].qpl], ['12MM', 'Airtech', 'MALIN Co / GERARD DANIEL']);
  const today = L.toDate_('2026-11-10');
  const r = L.classify_(L.parseRows_([
    { item: 'KERN00286-003', qty: 30, needBy: '2027-06-10' },
    { item: 'KERN00286-004', qty: 1, needBy: '2028-02-01' },   // 단가 기간 지남
    { item: 'MS20995C20', qty: 3, needBy: '2027-01-15' },       // DB에 있지만 단가 없음
    { item: 'NEW-1', qty: 1, needBy: '2027-01-15' },            // DB에 없음
    { item: 'KERN00286-004', qty: 5, needBy: '2027-03-01' },
  ]).requests, c, today);
  assert.deepStrictEqual(r.map(x => [x.status, x.route]), [['Spot', 'spotPriced'], ['Spot', 'rfq'], ['Spot', 'rfq'], ['Spot', 'rfq'], ['Spot', 'spotPriced']]);
  assert.strictEqual(r[0].candidates[0].price, 462);
  assert.deepStrictEqual(r[0].warnings, []);                  // 30개: Min 28, Multi 2 를 만족, 리드타임 충분
  assert.match(r[1].reason, /단가 기간 지남/);
  assert.deepStrictEqual([r[2].reason, r[2].qpl, r[2].uom], ['Spot 자재 · DB 단가 없음', 'MALIN Co / GERARD DANIEL', 'RL']);
  assert.strictEqual(r[3].reason, '계약 없음 · DB 미등록 자재');
});

test('MOQ 읽기: 숫자 = 최소, Min x = 최소, Multi x = x개 단위', () => {
  assert.deepStrictEqual(L.parseMoq_('Min 28, Multi 2'), { min: 28, multi: 2 });
  assert.deepStrictEqual(L.parseMoq_('Min 8'), { min: 8, multi: 0 });
  assert.deepStrictEqual(L.parseMoq_('Multi 16'), { min: 0, multi: 16 });
  assert.deepStrictEqual(L.parseMoq_(6000), { min: 6000, multi: 0 });
  assert.deepStrictEqual(L.parseMoq_('12,500'), { min: 12500, multi: 0 });
  assert.deepStrictEqual(L.parseMoq_('Min 819, Multi 819'), { min: 819, multi: 819 });
  for (const v of ['-', '없음', '', null]) assert.deepStrictEqual(L.parseMoq_(v), { min: 0, multi: 0 });
  assert.strictEqual(L.moqWarning_(25, 'Min 28, Multi 2'), '수량 25 → MOQ(최소 28, 2개 단위) · 28개로 주문');
  assert.strictEqual(L.moqWarning_(31, 'Min 28, Multi 2'), '수량 31 → MOQ(최소 28, 2개 단위) · 32개로 주문');
  assert.strictEqual(L.moqWarning_(30, 'Min 28, Multi 2'), '');
  assert.strictEqual(L.moqWarning_(20, 'Multi 16'), '수량 20 → MOQ(16개 단위) · 32개로 주문');
  assert.strictEqual(L.moqWarning_(900, 'Min 819, Multi 819'), '수량 900 → MOQ(최소 819, 819개 단위) · 1638개로 주문');
  assert.strictEqual(L.moqWarning_(5, '-'), '');
});

test('견적요청서 제목', () => {
  const base = L.toDate_('2026-11-10');
  assert.strictEqual(L.rfqTitle_(base, 'Boeing', 'Chemical', undefined, '2026-11-20'), '[KE] RFQ // 261110 Boeing Chemical Material Spot (due 11/20)');
  assert.strictEqual(L.rfqTitle_(base, '', ' Safety  wire ', 'Spot material', '2026-11-09'), '[KE] RFQ // 261110 Safety wire Spot material (due 11/9)');
  assert.strictEqual(L.rfqTitle_(base, '', '', '', ''), '[KE] RFQ // 261110');
});

test('견적요청서 월 칸: 첫 달부터 12개월', () => {
  const base = L.toDate_('2026-11-10');
  assert.deepStrictEqual(L.rfqMonth_(base, L.toDate_('2026-11-30')), { k: 0, note: '' });
  assert.deepStrictEqual(L.rfqMonth_(base, L.toDate_('2027-06-10')), { k: 7, note: '' });
  assert.strictEqual(L.rfqMonth_(base, L.toDate_('2027-10-31')).k, 11);
  assert.match(L.rfqMonth_(base, L.toDate_('2027-11-01')).note, /12개월 넘게 뒤/);
  assert.match(L.rfqMonth_(base, L.toDate_('2026-10-01')).note, /이전/);
});

test('견적요청서 작성 계획', () => {
  const base = L.toDate_('2026-11-10');
  const ops = L.buildRfqPlan_([
    { item: 'MS20995C20', registered: false, name: 'MS20995C20', desc: 'Safety Wire', qpl: 'MALIN Co', uom: 'RL', qty: 3, needBy: L.toDate_('2027-01-15') },
    { item: 'KERC00016', registered: true, name: 'PR-1776', desc: 'PRIMER', qpl: 'PPG', uom: 'KT', remark: '', qty: 5, needBy: L.toDate_('2027-03-01') },
  ], base);
  const get = (r, c) => (ops.filter(o => o.op === 'set' && o.r === r && o.c === c).pop() || {}).v;
  // 월 칸은 가장 이른 Need-by(2027-01-15) 달부터 12개월
  assert.deepStrictEqual(get(2, 9), { date: '2027-01-01' });
  assert.deepStrictEqual(get(2, 20), { date: '2027-12-01' });
  assert.deepStrictEqual([get(3, 3), get(3, 4), get(3, 8), get(3, 9), get(3, 21)], ['TBD', 'MS20995C20', 'MFG COC/TR', 3, '=SUM(I3:T3)']);
  assert.deepStrictEqual([get(4, 2), get(4, 3), get(4, 8), get(4, 11), get(4, 9)], [2, 'KERC00016', '', 5, '']);
  assert.deepStrictEqual(ops.filter(o => o.op === 'rowFmt').map(o => [o.fr, o.r]), [[3, 4]]);
  assert.deepStrictEqual(ops.filter(o => o.op === 'width').map(o => [o.c, o.px]), Array.from({ length: 12 }, (_, i) => [9 + i, 64]).concat([[21, 80], [8, 110]]));
});

test('견적요청 메일 본문', () => {
  const m = L.rfqEmail_([{ name: 'LOCTITE EA 9837.1', desc: 'BMS 8-341 Type III <A&B>', qpl: 'Henkel', uom: 'EA', qty: 20000, needBy: '' }], '11/20', '<b>김민형 사원</b>');
  assert.match(m.html, /as you can find in the below due 11\/20\./);
  assert.match(m.html, /LOCTITE EA 9837\.1/);
  assert.match(m.html, /&lt;A&amp;B&gt;/);
  assert.match(m.html, />20,000</);
  assert.match(m.html, />ASAP</);
  assert.match(m.html, /MFG COC\/TR/);
  assert.ok(m.html.indexOf('Thank you') < m.html.indexOf('김민형'));
  assert.match(m.text, /1\. LOCTITE EA 9837\.1 \| BMS 8-341/);
});

test('Spot 견적 구매내역서 계획: 업체 수만큼 줄·비교 칸, 선정 체크, NO BID, 요약 순서 수식', () => {
  const base = L.toDate_('2026-11-10');
  const items = [
    { item: 'KERN00278-001', desc: 'FEP FILM 2MIL', material: 'FEP Film', qpl: 'Airtech', uom: 'RL', qty: 64, needBy: L.toDate_('2027-05-10'), program: '공통 / 18M DOS',
      prev: { year: 2026, qty: 24, cur: 'usd', price: 964.19 },
      quotes: { Airtech: { sel: true, review: 'Sole source', qty: 64, mfg: 'A4000V', uom: 'RL', price: 1, cur: 'usd', lt: '4~6' }, Hexcel: { price: 'no bid' } } },
    { item: 'MS20995C20', desc: 'Safety Wire', uom: 'RL', qty: 3, needBy: L.toDate_('2027-01-15'), program: '', quotes: {} },
  ];
  const vendorInfo = L.vendorInfo_([['', '업체명', '이메일', 'SUPPLIER', 'SITE'], ['', 'Airtech', 'a@x.com', 'AIRTECH INTERNATIONAL INC', 'HUNTINGTON BEACH'], ['', 'Hexcel', 'h@x.com', 'HEXCEL CORP', '']]);
  const ops = L.buildSpotPlan_(items, ['Airtech', 'Hexcel', '3M'], { USD: 1355.6, EUR: 1740.98 }, base, null, { title: 'T', confirmer: '안성미 과장', vendorInfo });
  const get = (s, r, c) => (ops.filter(o => o.op === 'set' && o.s === s && o.r === r && o.c === c).pop() || {}).v;
  // 견적종합: 품목 2 × 업체 3 = 6줄 (5~10행)
  assert.deepStrictEqual([5, 6, 7, 8, 9, 10].map(r => get('견적종합', r, 30)), ['Airtech', 'Hexcel', '3M', 'Airtech', 'Hexcel', '3M']);
  assert.deepStrictEqual([get('견적종합', 5, 1), get('견적종합', 5, 2), get('견적종합', 6, 1)], [true, 'Sole source', false]);
  assert.deepStrictEqual([get('견적종합', 5, 34), get('견적종합', 5, 35), get('견적종합', 6, 34)], [1, 'USD', 'NO BID']);
  // 월 칸은 가장 이른 Need-by(2027-01) 달부터: 2027-01 → L열, 2027-05 → P열
  assert.deepStrictEqual(get('견적종합', 3, 12), { date: '2027-01-01' });
  assert.strictEqual(get('견적종합', 5, 16), 64);
  assert.strictEqual(get('견적종합', 8, 12), 3);
  assert.strictEqual(get('견적종합', 5, 29), '=IF(AB5="", "", COUNTIFS($E$5:$E$10, E5, $AB$5:$AB$10, "<"&AB5)+1)');   // 같은 품목 안 순위
  assert.match(get('견적종합', 5, 28), /IF\(AI5="SGD", AE5\*AH5\*\$AS\$2/);
  assert.deepStrictEqual(ops.filter(o => o.op === 'checkbox').map(o => [o.r1, o.r2]), [[5, 10]]);
  assert.deepStrictEqual(ops.filter(o => o.op === 'cf').map(o => [o.formula, o.bg, o.r2, o.c1, o.c2]),
    [['=$A5=TRUE', '#ffff00', 10, 1, 40], ['=$AH5="NO BID"', '#d9d9d9', 10, 1, 40], ['=$A5=TRUE', '#ffff00', 10, 46, 47], ['=$AH5="NO BID"', '#d9d9d9', 10, 46, 47]]);
  // 견적종합 AT·AU: 업체정보 탭 SUPPLIER·SITE (업체마다)
  assert.deepStrictEqual([get('견적종합', 5, 46), get('견적종합', 5, 47), get('견적종합', 6, 46), get('견적종합', 6, 47), get('견적종합', 7, 46)],
    ['AIRTECH INTERNATIONAL INC', 'HUNTINGTON BEACH', 'HEXCEL CORP', undefined, undefined]);
  // 구매내역서: W Need-by · X SUPPLIER · Y SITE (선정 업체), 비교 칸은 Z 부터 3개 업체(Z~AK)만 남기고 AL부터 지움
  assert.strictEqual(get('구매내역서', 11, 23), '2027-05-10');
  assert.strictEqual(get('구매내역서', 11, 24), "=IFERROR(INDEX('견적종합'!$AT:$AT, MATCH(TRUE&$B11, '견적종합'!$D:$D, 0))&\"\", \"\")");
  assert.strictEqual(get('구매내역서', 11, 25), "=IFERROR(INDEX('견적종합'!$AU:$AU, MATCH(TRUE&$B11, '견적종합'!$D:$D, 0))&\"\", \"\")");
  assert.deepStrictEqual([get('구매내역서', 8, 26), get('구매내역서', 8, 30), get('구매내역서', 8, 34)], ['Airtech', 'Hexcel', '3M']);
  assert.deepStrictEqual(ops.filter(o => o.op === 'deleteCols').map(o => [o.c, o.n]), [[38, 68]]);
  assert.ok(!ops.some(o => o.op === 'width' && o.s === '구매내역서'));     // 열 너비는 템플릿이 정함
  assert.strictEqual(get('구매내역서', 11, 2), "=IFERROR(INDEX('견적종합'!$F:$F, MATCH($A11, '견적종합'!$E:$E, 0)), \"\")");
  assert.strictEqual(get('구매내역서', 11, 16), "=IFERROR(INDEX('견적종합'!$AD:$AD, MATCH(TRUE&$B11, '견적종합'!$D:$D, 0))&\"\", \"미정\")");
  assert.strictEqual(get('구매내역서', 11, 26), "=IFERROR(INDEX('견적종합'!$AH:$AH, MATCH($B11&Z$8, '견적종합'!$C:$C, 0)), \"\")");
  assert.deepStrictEqual([6, 7, 8, 9].map(c => get('구매내역서', 11, c)), [2026, 24, 'USD', 964.19]);
  assert.strictEqual(get('구매내역서', 12, 14), undefined);       // Program 비면 노란 칸
  assert.ok(ops.some(o => o.op === 'bg' && o.s === '구매내역서' && o.r === 12 && o.c === 14));
  assert.deepStrictEqual([get('구매내역서', 3, 16), get('구매내역서', 3, 17), get('구매내역서', 3, 18)], [1355.6, undefined, 1740.98]);
  // 전체구매요약: K열 L/I → L=K+10 → 구매내역서 행
  assert.deepStrictEqual([get('전체구매요약', 5, 11), get('전체구매요약', 5, 12), get('전체구매요약', 6, 12)], [1, '=K5+10', '=K6+10']);
  assert.strictEqual(get('전체구매요약', 7, 3), 'CUR 별 합계');
  // CUR 별 합계: 환율 넣은 USD·EUR + KRW 만 (JPY·SGD 줄 없음), 총 합계도 그 통화만
  assert.deepStrictEqual([7, 8, 9].map(r => get('전체구매요약', r, 8)), ['USD', 'EUR', 'KRW']);
  assert.strictEqual(get('전체구매요약', 10, 3), '총 합계 (KRW 표기)');
  assert.strictEqual(get('전체구매요약', 10, 9), "=I7*'구매내역서'!$P$3+I8*'구매내역서'!$R$3+I9");
  assert.match(get('전체구매요약', 11, 3), /USD "&TEXT\('구매내역서'!\$P\$3,"#,##0\.00"\)&" · EUR "&TEXT\('구매내역서'!\$R\$3,"#,##0\.00"\)$/);
  assert.ok(!/JPY|SGD/.test(get('전체구매요약', 11, 3)));
  // 테두리: 견적종합 X~AN, 구매내역서 표 전체 (업체 3곳 → V+12 = AH 까지)
  assert.deepStrictEqual(ops.filter(o => o.op === 'border').map(o => [o.s, o.r1, o.c1, o.r2, o.c2, o.edge || '']),
    [['견적종합', 5, 24, 10, 40, ''], ['견적종합', 5, 24, 10, 24, 'left'], ['견적종합', 5, 46, 10, 47, ''], ['구매내역서', 11, 1, 12, 37, '']]);
  // 사업부 담당 확인 · 구매일정도 들어감
  assert.strictEqual(get('사업부 담당 확인', 2, 1), '안성미 과장');
  assert.ok(ops.some(o => o.s === '구매일정' && o.op === 'width'));
  assert.throws(() => L.buildSpotPlan_(items, [], {}, base), /업체가 없습니다/);
});

test('업체정보 탭 읽기 (머리글 위치 자동, 한 업체 여러 줄) · "-" 는 빈칸', () => {
  const v = L.parseVendors_([[null, null, null], [null, '업체명', '이메일'], [null, 'OMI', 'a@omi.com'], [null, 'OMI', 'b@omi.com; a@omi.com'], [null, '3M', '잘못된주소'], [null, 'PPG', 'p@ppg.com']]);
  assert.deepStrictEqual(v, [{ name: 'OMI', emails: ['a@omi.com', 'b@omi.com'] }, { name: 'PPG', emails: ['p@ppg.com'] }]);
  const c = L.parseContracts_([['Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료', 'Ship Form', 'QPL'], ['K1', 'Spot', '-', '-', '-', '-', '-', '-', 'MOMENTIVE_NY']]);
  assert.deepStrictEqual([c[0].supplier, c[0].price, c[0].cur, c[0].start, c[0].shipFrom, c[0].qpl], ['', null, '', null, '', 'MOMENTIVE_NY']);
  const r = L.classify_(L.parseRequest_('K1 5 2027-01-01').requests, c, L.toDate_('2026-11-10'))[0];
  assert.deepStrictEqual([r.route, r.qpl], ['rfq', 'MOMENTIVE_NY']);
});

test('회신 없음 업체는 견적종합에 NO BID · "회신 없음" 으로', () => {
  const ops = L.buildSpotPlan_([{ item: 'A', qty: 5, uom: 'EA', needBy: L.toDate_('2027-01-10'), quotes: { C: { noReply: true, qty: '5', uom: 'EA', price: '12' } } }], ['C'], {}, L.toDate_('2026-11-10'));
  const get = (r, c) => (ops.filter(o => o.op === 'set' && o.s === '견적종합' && o.r === r && o.c === c).pop() || {}).v;
  assert.deepStrictEqual([get(5, 34), get(5, 40), get(5, 31)], ['NO BID', '회신 없음', 5]);
});

test('같은 자재는 Need-by 가 달라도 견적요청서 한 줄, 해당 월 칸마다 수량', () => {
  const rows = [
    { item: 'KERN00534-001', name: 'PC938', qty: 244, needBy: '2027-09-10', program: '787' },
    { item: 'KERN00718-001', name: 'EDDING 751', qty: 360, needBy: '2027-05-01' },
    { item: 'kern00534-001', name: 'PC938', qty: 244, needBy: '2027-03-15', program: '737' },
    { item: 'KERN00534-001', name: 'PC938', qty: 10, needBy: '2027-03-30' }];
  const m = L.mergeRfqItems_(rows);
  assert.strictEqual(m.length, 2);
  assert.strictEqual(m[0].item, 'KERN00534-001');
  assert.strictEqual(m[0].qty, 498);
  assert.strictEqual(m[0].needBy, '2027-03-15');               // 가장 이른 날
  assert.deepStrictEqual(m[0].parts.map(p => p.needBy), ['2027-03-15', '2027-03-30', '2027-09-10']);
  assert.strictEqual(m[0].program, '737, 787');
  assert.strictEqual(L.mergeRfqItems_(m).length, 2);            // 다시 합쳐도 그대로
  assert.strictEqual(L.mergeRfqItems_(m)[0].qty, 498);

  const base = L.toDate_('2026-11-10');
  const ops = L.buildRfqPlan_(m, base);
  const cell = (r, c) => ops.filter(o => o.op === 'set' && o.r === r && o.c === c).pop();
  const M0 = L.RFQ.C.M0;
  // 첫 달 = 가장 이른 Need-by 2027-03
  assert.deepStrictEqual(cell(2, M0).v, { date: '2027-03-01' });
  assert.strictEqual(cell(3, M0).v, 254);                       // 2027-03: 244 + 10
  assert.strictEqual(cell(3, M0 + 6).v, 244);                   // 2027-09
  assert.strictEqual(cell(3, M0 + 1).v, '');
  assert.strictEqual(cell(4, M0 + 2).v, 360);                   // 2027-05
  assert.strictEqual(cell(4, L.RFQ.C.LI).v, 2);                 // L/I 1, 2 …

  const mail = L.rfqEmail_(m, '11/18', '');
  assert.match(mail.html, /2027-03-15 \(244\)<br>2027-03-30 \(10\)<br>2027-09-10 \(244\)/);
  assert.match(mail.html, />498</);
  assert.match(mail.text, /\| 498 \| 2027-03-15 \(244\), 2027-03-30 \(10\), 2027-09-10 \(244\)/);

  const sp = L.buildSpotPlan_([Object.assign({ quotes: {} }, m[0])], ['A'], {}, base);
  const q = (c) => sp.filter(o => o.op === 'set' && o.s === L.SPOT.QUOTE && o.r === 5 && o.c === c).pop();
  assert.strictEqual(q(L.SPOT.Q_M0).v, 254);
  assert.strictEqual(q(L.SPOT.Q_M0 + 6).v, 244);
});

test('예전 기록(parts 없음)은 Need-by 한 건으로', () => {
  const ps = L.rfqParts_({ qty: '1,200', needBy: '2027-01-05' });
  assert.strictEqual(ps.length, 1);
  assert.strictEqual(ps[0].qty, 1200);
  assert.strictEqual(L.fmtDate_(ps[0].needBy), '2027-01-05');
});

test('로그인 유지 시간 4시간', () => {
  const t = Date.now();
  assert.strictEqual(L.SESSION_HOURS, 4);
  assert.ok(L.sessionValid_(t - 3.9 * 3600000, t));
  assert.ok(!L.sessionValid_(t - 4 * 3600000, t));
  assert.ok(!L.sessionValid_(0, t));
  assert.ok(L.isEmail_(' A.b@ke.co.kr ') && !L.isEmail_('ab@') && !L.isEmail_('a b@x.com'));
});

test('견적요청서 템플릿 2 (목록형): Need-by 마다 한 줄, 같은 자재도 다음 줄', () => {
  const items = [
    { item: 'KERN00534-001', registered: true, name: 'PC938', desc: 'MARKING PENCIL', qpl: 'Newell', uom: 'EA', remark: 'COC & TR', qty: 488, needBy: '2027-03-01',
      parts: [{ qty: 244, needBy: '2027-03-01', program: 'AH-64E' }, { qty: 244, needBy: '2027-09-01', program: 'EFAB' }] },
    { item: 'MS20995C20', name: 'MS20995C20', desc: 'Safety Wire', uom: 'RL', qty: 3, needBy: '', program: '787' }];
  const rows = L.splitRfqItems_(items);
  assert.deepStrictEqual(rows.map(r => [r.name, r.qty, r.needBy, r.program]), [['PC938', 244, '2027-03-01', 'AH-64E'], ['PC938', 244, '2027-09-01', 'EFAB'], ['MS20995C20', 3, '', '787']]);
  assert.ok(rows.every(r => !r.parts));
  const ops = L.buildRfq2Plan_(rows);
  const get = (r, c) => (ops.filter(o => o.op === 'set' && o.r === r && o.c === c).pop() || {}).v;
  assert.deepStrictEqual([1, 2, 3, 4, 5, 7, 10, 11, 12, 14].map(c => get(2, c)),
    [1, 'PC938', 'MARKING PENCIL', 244, 'EA', '=IF(ISNUMBER(F2), D2*F2, "")', 'Newell', '2027.03.01', 'COC & TR', 'AH-64E']);
  assert.deepStrictEqual([get(3, 1), get(3, 2), get(3, 11)], [2, 'PC938', '2027.09.01']);
  assert.deepStrictEqual([get(4, 11), get(4, 12), get(4, 6), get(4, 15)], ['ASAP', 'MFG COC/TR', undefined, undefined]);   // U/P·Supplier 는 업체가 적음
  assert.deepStrictEqual(ops.filter(o => o.op === 'rowFmt').map(o => [o.fr, o.r]), [[2, 3], [2, 4]]);
  assert.strictEqual(L.rfq2FileName_(L.toDate_('2026-11-10'), 'HW', 'ywk', 'AH-64E EFAB'), 'RFQ_261110_KE_HW_YWK_AH-64E_EFAB');
  assert.strictEqual(L.rfq2FileName_(L.toDate_('2026-11-10'), ' ', 'YWK', ''), 'RFQ_261110_KE_YWK');
});

// 구글 시트처럼: 서식 복사(fmt·cellFmt·rowFmt) 대상이 병합 칸과 일부만 겹치면 오류
function simulateMerges(ops, initial) {
  const merges = {};
  Object.keys(initial).forEach(s => { merges[s] = initial[s].map(m => Object.assign({}, m)); });
  const list = s => (merges[s] = merges[s] || []);
  const partial = (m, r1, c1, r2, c2) => {
    const inter = !(m.r2 < r1 || m.r1 > r2 || m.c2 < c1 || m.c1 > c2);
    const inside = m.r1 >= r1 && m.r2 <= r2 && m.c1 >= c1 && m.c2 <= c2;
    return inter && !inside;
  };
  ops.forEach((o, i) => {
    if (o.op === 'unmerge') merges[o.s] = list(o.s).filter(m => m.r2 < o.r1 || m.r1 > o.r2 || m.c2 < o.c1 || m.c1 > o.c2);
    else if (o.op === 'merge') list(o.s).push({ r1: o.r1, c1: o.c1, r2: o.r2, c2: o.c2 });
    else if (o.op === 'cellFmt' || o.op === 'fmt' || o.op === 'rowFmt') {
      const c1 = o.op === 'cellFmt' ? o.c : o.c1, c2 = o.op === 'cellFmt' ? o.c : o.c2;
      const bad = list(o.s).find(m => partial(m, o.r, c1, o.r, c2));
      if (bad) throw new Error('op ' + i + ' ' + JSON.stringify(o) + ' 가 병합 ' + JSON.stringify(bad) + ' 와 일부 겹침');
    }
  });
}
const pair = (r, c1, c2) => ({ r1: r, c1, r2: r, c2 });
test('구매일정: 서식 복사가 템플릿 병합 칸과 겹치지 않는다 (계약·Spot 템플릿)', () => {
  const months = []; for (let c = 5; c <= 27; c += 2) months.push(pair(5, c, c + 1));
  // 계약 템플릿: 4행 E4:J4 · K4:Z4, 5행 달마다 2칸 (E5:F5 … AA5:AB5)
  const contract = { 구매일정: [pair(4, 5, 10), pair(4, 11, 26)].concat(months) };
  // Spot 템플릿: 4행 E4:H4 · I4:X4, 5행 E5:F5 … W5:X5 · AA5:AB5
  const spot = { 구매일정: [pair(4, 5, 8), pair(4, 9, 24)].concat(months.filter(m => m.c1 !== 25)) };
  const base = L.toDate_('2026-11-11');
  const ops = L.scheduleConfirmOps_(base, null, {});
  simulateMerges(ops, contract);
  simulateMerges(ops, spot);
  const line = { item: 'A', qty: 1, needBy: L.toDate_('2027-06-10'), supplier: 'AKZO', cur: 'USD', price: 2, priceMode: 'LTA' };
  simulateMerges(L.buildPurchasePlan_([line], ['USD'], { USD: 1350 }, base), contract);
});

test('업체정보 SUPPLIER·SITE 열 (업체명 열이 SUPPLIER 보다 앞)', () => {
  const v = [['', '업체명', '이메일', 'SUPPLIER', 'SITE', 'CONTACT'], ['', 'OMI', 'a@omi.com', 'OMI SPECIALTY PRODUCTS, LLC', 'LA', 'x'], ['', 'OMI', 'b@omi.com', '', '', ''],
    ['', 'GV엔지니어링', 'g@gv.com', 'GV 엔지니어링', 6138159311, '']];
  assert.deepStrictEqual(JSON.parse(JSON.stringify(L.parseVendors_(v))), [{ name: 'OMI', emails: ['a@omi.com', 'b@omi.com'] }, { name: 'GV엔지니어링', emails: ['g@gv.com'] }]);
  const info = L.vendorInfo_(v);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(L.vendorOf_(info, 'omi'))), { supplier: 'OMI SPECIALTY PRODUCTS, LLC', site: 'LA' });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(L.vendorOf_(info, 'GV 엔지니어링'))), { supplier: 'GV 엔지니어링', site: '6138159311' });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(L.vendorOf_(info, '없는업체'))), { supplier: '', site: '' });
});

test('계약 구매내역서: 같은 자재·같은 계약은 한 줄, Need-by 칸에 날짜(수량)', () => {
  const base = L.toDate_('2026-11-11');
  const c = { item: 'KERC00064', row: 7, supplier: 'AKZO', cur: 'USD', price: 78.49, priceMode: 'LTA', uom: 'KT' };
  const lines = L.mergePurchaseLines_([
    Object.assign({}, c, { qty: 600, needBy: L.toDate_('2027-06-20'), program: '787' }),
    Object.assign({}, c, { item: 'KERC00149', row: 9, supplier: 'Syensqo', price: 39.23, qty: 200, needBy: L.toDate_('2027-03-01') }),
    Object.assign({}, c, { qty: 500, needBy: L.toDate_('2027-05-14'), program: '737' }),
    Object.assign({}, c, { row: 8, price: 80, qty: 10, needBy: L.toDate_('2028-02-01') })]);   // 다른 해 단가 → 따로
  assert.deepStrictEqual(lines.map(l => [l.item, l.qty, l.needs.length, l.program]), [['KERC00064', 1100, 2, '787, 737'], ['KERC00149', 200, 1, ''], ['KERC00064', 10, 1, '']]);
  assert.strictEqual(L.fmtDate_(lines[0].needBy), '2027-05-14');
  lines[0].supplierName = 'AKZO NOBEL'; lines[0].site = 'US';
  const ops = L.buildPurchasePlan_(lines, ['USD'], { USD: 1350 }, base);
  const get = (r, c2) => (ops.filter(o => o.op === 'set' && o.s === '구매세부내역' && o.r === r && o.c === c2).pop() || {}).v;
  assert.deepStrictEqual([get(7, 20), get(7, 21), get(7, 22)], ['Need-by\nDate', 'SUPPLIER', 'SITE']);
  assert.deepStrictEqual([get(8, 8), get(8, 20), get(8, 21), get(8, 22)], [1100, '2027-05-14(500)\n2027-06-20(600)', 'AKZO NOBEL', 'US']);
  assert.deepStrictEqual(get(9, 20), { date: '2027-03-01' });
  assert.ok(ops.some(o => o.op === 'wrap' && o.r === 8 && o.c === 20));
  assert.ok(ops.some(o => o.op === 'height' && o.s === '구매세부내역' && o.r === 8 && o.pt === 30));
  assert.ok(ops.filter(o => o.op === 'fmt' && o.s === '구매세부내역').every(o => o.c2 === 22));
  // 템플릿 예시 모양은 맨 먼저 지우고 다시 그림
  assert.deepStrictEqual(ops.find(o => o.s === '구매세부내역'), { op: 'clear', s: '구매세부내역', r1: 1, c1: 1, r2: 30, c2: 22 });
  assert.ok(!ops.some(o => o.op === 'width' && o.s === '구매세부내역'));   // 열 너비는 템플릿이 정함
});

test('하나로 만들기: 계약 품목(fixed)은 견적종합에 계약 업체 한 줄로 미리 선정', () => {
  const base = L.toDate_('2026-11-13');
  const items = [
    { item: 'KERC00064', name: 'PRIMER', desc: 'BMS10-11', uom: 'KT', qty: 1020, needBy: L.toDate_('2027-03-15'), parts: [{ qty: 500, needBy: '2027-03-15' }, { qty: 520, needBy: '2027-04-15' }],
      program: '787', fixed: { vendor: 'AKZO', price: 78.49, cur: 'USD', lt: '10-12wKs', reason: 'LTA' } },
    { item: 'KERN00736-001', name: '88TB0110', uom: 'SH', qty: 1, needBy: L.toDate_('2027-03-01'), quotes: { OMI: { price: 5, cur: 'USD', qty: 1 }, CBOL: { price: 'NO BID' } } }];
  const ops = L.buildSpotPlan_(items, ['OMI', 'CBOL'], { USD: 1350 }, base, null, { vendorInfo: L.vendorInfo_([['업체명', '이메일', 'SUPPLIER'], ['AKZO', 'a@x.com', 'AKZO NOBEL']]) });
  const get = (s2, r, c) => (ops.filter(o => o.op === 'set' && o.s === s2 && o.r === r && o.c === c).pop() || {}).v;
  // 5행: 계약 품목 한 줄 (선정 TRUE, 사유 LTA, AKZO 78.49 USD), 6~7행: 견적 품목 업체 2줄
  assert.deepStrictEqual([1, 2, 5, 30, 31, 34, 35, 38, 46].map(c => get('견적종합', 5, c)), [true, 'LTA', 1, 'AKZO', 1020, 78.49, 'USD', '10-12wKs', 'AKZO NOBEL']);
  assert.deepStrictEqual([get('견적종합', 6, 5), get('견적종합', 6, 30), get('견적종합', 7, 30), get('견적종합', 8, 30)], [2, 'OMI', 'CBOL', undefined]);
  assert.deepStrictEqual(ops.filter(o => o.op === 'checkbox').map(o => [o.r1, o.r2]), [[5, 7]]);
  assert.match(get('견적종합', 6, 29), /COUNTIFS\(\$E\$5:\$E\$7, E6, \$AB\$5:\$AB\$7/);
  assert.strictEqual(get('구매내역서', 12, 2), "=IFERROR(INDEX('견적종합'!$F:$F, MATCH($A12, '견적종합'!$E:$E, 0)), \"\")");
  assert.strictEqual(get('구매내역서', 11, 23), '2027-03-15(500)\n2027-04-15(520)');
});

test('견적 입력용 시트: 사이트 표와 같은 열, Need-by 여러 건은 날짜(수량)', () => {
  const rows = L.quoteSheetRows_([
    { item: 'KERN00736-001', name: '88TB0110', qty: 1, uom: 'SH', needBy: '2027-03-01' },
    { item: 'KERC00064', name: 'PRIMER', qty: 1020, uom: 'KT', needBy: '2027-03-15', parts: [{ qty: 500, needBy: '2027-03-15' }, { qty: 520, needBy: '2027-04-15' }] }]);
  assert.deepStrictEqual(rows[0].slice(0, 6), ['No', 'Item code', 'Item name', '요청 수량', 'Need-by', 'U/P (단가)']);
  assert.strictEqual(rows[0].length, 14);
  assert.deepStrictEqual(rows[1], [1, 'KERN00736-001', '88TB0110', 1, '2027-03-01', '', '', 1, 'SH', '', '', '', '', '']);
  assert.strictEqual(rows[2][4], '2027-03-15(500), 2027-04-15(520)');
  const used = {};
  assert.deepStrictEqual(['A/B:C', 'a/b:c', '사용 방법'].map(n => L.sheetTabName_(n, used)), ['A B C', 'a b c (2)', '사용 방법']);
});

test('구매일정 표시 개월: 기본 12, 늘리면 그만큼 칸·병합·서식 (예전 배열 형식도 그대로)', () => {
  const base = L.toDate_('2026-11-13');
  const ops18 = L.scheduleConfirmOps_(base, { months: 18, bars: [{ start: 0, len: 30 }, { start: 30, len: 6 }] }, {});
  const set = (r, c) => (ops18.filter(o => o.op === 'set' && o.s === '구매일정' && o.r === r && o.c === c).pop() || {}).v;
  const w = ops18.filter(o => o.op === 'width');
  assert.strictEqual(w.length, 36);                                   // 18개월 × 2칸 (E~AN)
  assert.deepStrictEqual([set(5, 5), set(5, 39), set(4, 5), set(4, 9)], [11, 4, '2026년', '2027년']);
  assert.deepStrictEqual(ops18.filter(o => o.op === 'merge' && o.r1 === 4).map(o => [o.c1, o.c2]), [[5, 8], [9, 32], [33, 40]]);   // 2026 · 2027 · 2028(1~4월)
  assert.strictEqual(ops18.filter(o => o.op === 'merge' && o.r1 === 5).length, 18);
  assert.ok(ops18.some(o => o.op === 'cellFmt' && o.r === 6 && o.c === 39 && o.fc === 23));
  assert.deepStrictEqual([set(6, 4), set(7, 4)], ['15.0M', '3.0M']);   // 30칸(15개월)도 잘리지 않음
  const bg = ops18.filter(o => o.op === 'bg' && o.r === 7 && o.color !== '#ffffff').map(o => o.c);
  assert.deepStrictEqual([bg[0], bg[bg.length - 1]], [35, 40]);
  // 예전 형식(막대 배열)·범위 밖 값
  assert.strictEqual(L.scheduleConfirmOps_(base, [{ start: 0, len: 4 }], {}).filter(o => o.op === 'width').length, 24);
  assert.strictEqual(L.schedMonths_({ months: 99 }), 36);
  assert.strictEqual(L.schedMonths_({ months: 3 }), 12);
});
