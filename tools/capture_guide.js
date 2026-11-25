// 사용법 팝업용 실제 화면 사진 찍기 (예시 데이터, 가짜 서버). 화면을 바꾼 뒤 다시 찍을 때:
//   LANG=ko_KR.UTF-8 node tools/capture_guide.js "$PWD" <작업 폴더>
//   python3 -I tools/pack_guide.py <작업 폴더>/guide apps-script/Guide.html
// Guide.html = JSON {img: {id: {src, w, h}}, hot: {단계 id: [x%, y%]}} — 단계 글은 Index.html 의 GUIDES
const { chromium } = require('playwright');
const fs = require('fs'), path = require('path');
const [repo, S] = process.argv.slice(2);
const OUT = path.join(S, 'guide'); fs.mkdirSync(OUT, { recursive: true }); fs.mkdirSync(path.join(S, 'ui'), { recursive: true });
fs.copyFileSync(path.join(repo, 'apps-script/Index.html'), path.join(S, 'ui/index.html'));
const L = require(path.join(repo, 'tests/load.js')).load();
const shots = {}, hots = {};
const TODAY = L.toDate_('2026-11-13');
const dl = (needBy, lt, ship) => JSON.parse(JSON.stringify(L.poDeadline_(needBy, lt, ship, TODAY)));   // 발주 마감일

const cand = (row, supplier, mode, price, desc, uom, need, lt, ship) => ({ row, supplier, priceMode: mode, price, cur: 'USD', uom, desc, start: '2026-01-01', end: '2028-12-31', leadtime: lt, deadline: dl(need, lt, ship) });
const res = (line, item, qty, needBy, route, alts, desc, uom, program) => ({ line, item, qty, needBy, program: program || '', route, candidates: alts.slice(0, 1), alts, warnings: [], reason: route === 'rfq' ? 'Spot 자재 · DB 단가 없음' : '', desc, uom, registered: true });
const RESULTS = [
  res(1, 'KERC00064', 500, '2027-05-14', 'contract', [cand(7, 'AKZO', 'LTA', 78.49, 'BMS10-11,TY-1,CL-A GREEN PRIMER', 'KT', '2027-05-14', '12wks', 'Ocean'), cand(8, 'PPG', '3PP', 80.1, 'BMS10-11,TY-1,CL-A GREEN PRIMER', 'KT', '2027-05-14', '10wks', 'Ocean')], 'BMS10-11,TY-1,CL-A GREEN PRIMER', 'KT', '787'),
  res(2, 'KERC00064', 600, '2027-06-20', 'contract', [cand(7, 'AKZO', 'LTA', 78.49, 'BMS10-11,TY-1,CL-A GREEN PRIMER', 'KT', '2027-06-20', '12wks', 'Ocean'), cand(8, 'PPG', '3PP', 80.1, 'BMS10-11,TY-1,CL-A GREEN PRIMER', 'KT', '2027-06-20', '10wks', 'Ocean')], 'BMS10-11,TY-1,CL-A GREEN PRIMER', 'KT', '787'),
  res(3, 'KERN00278-014', 64, '2027-05-10', 'spotPriced', [cand(20, 'Airtech', 'SPOT', 83.58, 'FEP FILM 2MIL A5000V', 'RL', '2027-05-10', '3wks', 'Air')], 'FEP FILM 2MIL A5000V', 'RL', 'AH-64E'),
  res(4, 'KERN00736-001', 3, '2027-03-01', 'rfq', [], '88TB0110 TAPE, MASKING', 'SH', 'AH-64E'),
  res(5, 'KERN02113-001', 10, '2027-03-15', 'rfq', [], 'A88TB0092-6 SPACER', 'PC', 'EFAB')];
const VSHEET = [[], ['', '업체명', '이메일', 'SUPPLIER', 'SITE', 'CONTACT'],
  ['', 'OMI', 'sales@omi.example.com', 'OMI SPECIALTY PRODUCTS, LLC', 'LA', 'Order desk'], ['', 'OMI', 'quote@omi.example.com', '', '', ''],
  ['', 'CBOL', 'paul@cbol.example.com', 'CBOL CORPORATION', 'CBOL', 'Paul (Sales)'], ['', 'CBOL', 'jun@cbol.example.com', '', '', 'Jun (Order)'], ['', 'CBOL', 'julie@cbol.example.com', '', '', ''],
  ['', 'Armacorp', 'sales@armacorp.example.com', 'ARMACORP INC', 'ARMACORP INC', ''], ['', 'Aerovac', 'sales@aerovac.example.com', 'AEROVAC SYSTEMS', 'KEIGHLEY', ''],
  ['', 'GV엔지니어링', 'sales@gv.example.com', 'GV 엔지니어링', '6138159311', '']];
const rfqItems = [{ item: 'KERN00736-001', registered: true, name: '88TB0110', desc: 'TAPE, MASKING', qty: 3, uom: 'SH', needBy: '2027-03-01' },
  { item: 'KERN02113-001', registered: true, name: 'A88TB0092-6', desc: 'SPACER', qty: 10, uom: 'PC', needBy: '2027-03-15' },
  { item: 'KERN00278-014', registered: true, name: 'A5000V 2MIL', desc: 'FEP FILM', qty: 64, uom: 'RL', needBy: '2027-05-10', parts: [{ qty: 32, needBy: '2027-05-10' }, { qty: 32, needBy: '2027-08-10' }] }];
const quote = { vendors: ['OMI', 'CBOL', 'Armacorp', 'Aerovac'], noReply: { Armacorp: true }, valid: { OMI: '2026-12-31', CBOL: '2026-11-18' }, extra: { OMI: { costs: [{ name: 'Freight', amount: '120', cur: 'USD' }, { name: 'Hazmat', amount: '45', cur: 'USD' }], mov: { amount: '300', cur: 'USD' } } }, rates: { USD: 1350 }, items: [
  { prev: { year: 2026, qty: 5, cur: 'USD', price: 9.4 }, quotes: { OMI: { price: 12.5, cur: 'USD', qty: 3, uom: 'SH', lt: '4wks', mfg: '3M' }, CBOL: { price: 11.8, cur: 'USD', qty: 3, uom: 'SH', lt: '6wks', mfg: '3M', sel: true }, Aerovac: { price: 'NO BID', qty: 3, uom: 'SH' } } },
  { quotes: { OMI: { price: 2.1, cur: 'USD', qty: 10, uom: 'PC', lt: '8wks', sel: true }, CBOL: { price: 1.95, cur: 'USD', qty: 10, uom: 'PC', lt: '10wks' } } },
  { quotes: { OMI: { price: 'NO BID' }, Aerovac: { price: 79, cur: 'USD', qty: 64, uom: 'RL', lt: '3wks', mfg: 'AEROVAC', sel: true } } }] };
let records = [
  { id: 'R1', kind: '견적요청서', title: '[KE] RFQ // 261113 Boeing Process Material Spot (due 11/20)', created: '2026-11-13', status: '메일 초안', url: '#', fileId: 'F1', draft: 'D', due: '2026-11-14',
    qfiles: { OMI: [{ id: 'qf1', name: 'OMI_견적서_261114_Quotation Q-2611.pdf', url: '#', at: '2026-11-14 10:20' }] }, vendors: quote.vendors.concat(['GV엔지니어링']), quote, items: rfqItems,
    mails: [{ no: 1, kind: '견적 요청', at: '2026-11-13 09:10', vendors: quote.vendors, count: 6 }], group: 'G1', fixed: [{ item: 'KERC00064', supplier: 'AKZO', qty: 1100, uom: 'KT', price: 78.49, cur: 'USD', priceMode: 'LTA', needBy: '2027-05-14', needs: [{ needBy: '2027-05-14', qty: 500 }, { needBy: '2027-06-20', qty: 600 }], desc: 'GREEN PRIMER' }] },
  { id: 'R2', fx: { at: '2026-11-13 10:12', rates: { USD: { v: 1350, src: 'auto', tabAt: '2026-11-13' } } }, kind: '견적 구매내역서', title: '261113 Boeing Process Material Spot (due 11/20) 견적 구매내역서', created: '2026-11-13', status: '작성', url: '#', items: [{}, {}, {}, {}], parent: 'R1', group: 'G1' },
  { id: 'R3', fx: { at: '2026-11-13 10:12', rates: { USD: { v: 1342.5, src: 'typed', tab: 1350, tabAt: '2026-11-13' } } }, kind: '계약 구매내역서', title: '261112 구매내역서_AKZO_Syensqo', created: '2026-11-12', status: '사업부 확인 요청', url: 'https://docs.google.com/spreadsheets/d/example', items: [{ deadline: '2026-11-24' }, { deadline: '2027-01-08' }], group: 'G0',
    confirmMail: { mode: 'send', at: '2026-11-12 16:40', to: ['biz@ke.com'] } },
  { id: 'R4', kind: 'Spot 구매내역서', title: '261112 Spot 구매내역서_Airtech', created: '2026-11-12', status: '품의 중', url: '#', items: [{ deadline: '2027-03-30' }], group: 'G0', po: { stage: '품의 중', approval: 'AP-2026-1188' } }];
records.forEach(r => { r.alert = JSON.parse(JSON.stringify(L.recordAlert_(r, TODAY))); });
const fieldsH = ['Item', '계약 구분', '업체명', '단가', 'CUR', '계약시작', '계약종료', 'MOQ', 'UOM', 'Ship From', 'Leadtime', 'Description', 'QPL', 'Type'];
const fields = fieldsH.map(h => ({ header: h, key: L.keyForHeader_(h) }));
// 업체 추천 · 성적표 · 자재 정보 · 완료 기록 (예시)
const st = (name, requested, replied, quoted, lowest, bought, ltWeeks, last) => ({ name, requested, replied, quoted, lowest, selected: lowest, bought, replyRate: requested ? Math.round(replied / requested * 100) : null, ltWeeks, last, items: ['KERN00736-001'] });
const HINTS = [st('CBOL', 3, 3, 3, 2, 1, 6, '2026-11-12'), st('OMI', 3, 2, 2, 1, 0, 4, '2026-11-12'), st('Aerovac', 2, 1, 1, 0, 0, 3, '2026-10-02')];
const STATS = [st('CBOL', 13, 6, 9, 7, 3, 3.1, '2026-11-15'), st('OMI', 9, 2, 4, 1, 1, 7.5, '2026-11-13'), st('Armacorp', 7, 2, 3, 2, 0, 12, '2026-11-13'), st('Aerovac', 4, 3, 3, 1, 1, 3, '2026-10-02'), st('GV엔지니어링', 7, 3, 4, 1, 0, 15, '2026-11-13')];
const mrow = (item, mode, sup, price, start, end, desc, uom, type) => [item, mode, sup, price, price === '' ? '' : 'USD', start, end, '', uom, mode === 'SPOT' ? 'Air' : 'Ocean', mode === 'SPOT' ? '' : '10wks', desc, sup, type];
const MASTER = { ver: '1', at: '2026-11-19 09:00', cols: fieldsH.map(h => ({ key: L.keyForHeader_(h), header: h })), rows: [
  mrow('KERC00064', 'LTA', 'AKZO', 72.1, '2024-01-01', '2024-12-31', 'BMS10-11,TY-1,CL-A GREEN PRIMER', 'KT', 'Chemical'),
  mrow('KERC00064', 'LTA', 'AKZO', 75.3, '2025-01-01', '2025-12-31', 'BMS10-11,TY-1,CL-A GREEN PRIMER', 'KT', 'Chemical'),
  mrow('KERC00064', 'LTA', 'AKZO', 78.49, '2026-01-01', '2026-12-31', 'BMS10-11,TY-1,CL-A GREEN PRIMER', 'KT', 'Chemical'),
  mrow('KERC00064', '3PP', 'PPG', 74, '2027-01-01', '2027-12-31', 'BMS10-11,TY-1,CL-A GREEN PRIMER', 'KT', 'Chemical'),
  mrow('KERC00149', '3PP', 'Syensqo', 390, '2026-01-01', '2026-11-30', 'PR-1776 SEALANT', 'EA', 'Sealant'),
  mrow('KERN00278-014', 'SPOT', 'Airtech', 83.58, '2026-01-01', '2028-12-31', 'FEP FILM 2MIL A5000V', 'RL', 'Film'),
  mrow('KERN00736-001', 'SPOT', '', '', '', '', '88TB0110 TAPE, MASKING', 'SH', 'Tape'),
  mrow('KERN02113-001', 'SPOT', '', '', '', '', 'A88TB0092-6 SPACER', 'PC', 'Hardware')] };
const ARCHIVED = [{ id: 'A1', kind: '견적요청서', title: '[KE] RFQ // 261105 Boeing HW Spot (due 11/12)', created: '2026-11-05', done: '2026-11-14', url: '#', items: [{}, {}] },
  { id: 'A2', kind: '견적 구매내역서', title: '261105 Boeing HW Spot (due 11/12) 견적 구매내역서', created: '2026-11-12', done: '2026-11-14', url: '#', items: [{}, {}] },
  { id: 'A3', kind: '계약 구매내역서', title: '261104 구매내역서_PPG', created: '2026-11-04', done: '2026-11-10', url: '#', items: [{}] }];
const vendorsResp = () => { const info = L.vendorInfo_(VSHEET); return { vendors: JSON.parse(JSON.stringify(L.parseVendors_(VSHEET).map(v => Object.assign(v, L.vendorOf_(info, v.name))))), missing: false }; };

(async () => {
  const b = await chromium.launch({ args: ['--lang=ko-KR'] });
  const pg = await b.newPage({ viewport: { width: 1440, height: 940 }, deviceScaleFactor: 1.25, reducedMotion: 'reduce', locale: 'ko-KR', timezoneId: 'Asia/Seoul' });
  pg.on('pageerror', e => console.log('PAGEERR', e.message));
  await pg.exposeFunction('srv', (fn, args) => {
    switch (fn) {
      case 'apiAuthState': return { account: 'buyer@ke.com', deployOk: true, loggedIn: true, expiresAt: Date.now() + 1e7, now: Date.now(), hours: 4, url: 'u' };
      case 'apiQuoteDups': return { dups: { 'KERN02113-001': [{ title: '261110 구매내역서_OMI', owner: 'lee@ke.example.com', mine: false, created: '2026-11-10' }] } };
      case 'apiInit': return { today: '2026-11-13', rates: { USD: 1350 }, rateInfo: { USD: { at: '2026-11-13', auto: true } }, items: [['KERC00064', 'BMS10-11,TY-1,CL-A GREEN PRIMER', 'Ocean', 'LTA, 3PP', '12MM'], ['KERN00278-014', 'FEP FILM 2MIL A5000V', 'Air', 'SPOT', ''], ['KERN00736-001', '88TB0110 TAPE, MASKING', '', 'SPOT', ''], ['KERN02113-001', 'A88TB0092-6 SPACER', '', 'SPOT', '']], fields, schedule: { tasks: L.SCHEDULE_TASKS, example: L.SCHEDULE_EXAMPLE, slots: 24 } };
      case 'apiMyRecords': return { me: '', records, archiveReady: true };
      case 'apiClassify': return { results: JSON.parse(JSON.stringify(RESULTS)), skipped: [], rates: { USD: 1350 } };
      case 'apiVendors': return vendorsResp();
      case 'apiSignature': return { html: '', source: 'gmail', gmailService: true };
      case 'apiRegisterItems': return { ok: false, errors: [{ i: 2, item: 'KERC09001', messages: ['LTA·3PP는 단가가 필요합니다.', 'LTA·3PP는 계약시작과 계약종료가 필요합니다.'] }] };
      case 'apiSaveQuotes': return { ok: true };
      case 'apiGuide': return '{"img":{},"hot":{}}';
      case 'apiDigestState': return { on: false, hour: 8 };
      case 'apiVendorHints': return HINTS;
      case 'apiVendorStats': return STATS;
      case 'apiMasterList': return MASTER;
      case 'apiArchivedRecords': return { records: ARCHIVED, ready: true };
      case 'apiDeletedRecords': return { records: [{ id: 'T1', kind: '견적요청서', title: '[KE] RFQ // 261020 테스트 견적요청서', created: '2026-10-20', done: '2026-10-28', deleted: '2026-11-10', url: '#', history: 0 },
        { id: 'T2', kind: '계약 구매내역서', title: '261021 구매내역서_PPG (테스트)', created: '2026-10-21', done: '2026-10-30', deleted: '2026-11-10', url: '#', history: 2 }] };
      case 'apiSystemCheck': return { at: '2026-11-13 09:05', ms: 9200, list: [
        { group: '배포 · 로그인', level: 'ok', msg: '웹앱이 "웹 앱에 액세스하는 사용자" 권한으로 실행됩니다' },
        { group: '자재정보 (DB)', level: 'ok', msg: '자재정보 스프레드시트: 소요제기 DB · 편집 권한' },
        { group: '자재정보 (DB)', level: 'ok', msg: '자재정보 탭: 머리글 2행, 품목 1,284줄 · 버전 1763000000000' },
        { group: '템플릿', level: 'ok', msg: '구매내역서 템플릿: 구매내역서_템플릿' },
        { group: '템플릿', level: 'ok', msg: 'Spot 견적 구매내역서 템플릿: Spot_구매내역서_템플릿' },
        { group: '완료 보관 파일', level: 'bad', msg: '보관 파일에 편집 권한이 없습니다 (VIEW) — 완료 · 되돌리기 · 삭제가 실패합니다' },
        { group: '잠금 · 캐시', level: 'ok', msg: '잠금: 41ms 에 받음 (여러 명 동시 저장 보호)' },
        { group: '진행기록 · 견적데이터', level: 'ok', msg: '견적데이터 탭: 큰 견적 2건 모두 읽힘' },
        { group: '환율', level: 'ok', msg: 'USD 1350 (수식 · 자동)' },
        { group: '환율', level: 'warn', msg: 'EUR 1473.47 — 12일 전 값 (손으로 적은 환율)' }] };
      case 'apiConfirmDefaults': return { to: 'biz.team@ke.com', cc: '', name: '홍길동', body: '', signature: '<div style="font-family:Arial,sans-serif;font-size:10pt;color:#444">홍길동 사원<br>자재구매팀</div>' };
      default: return null;
    }
  });
  await pg.addInitScript(() => {
    const mk = (ok, fail) => new Proxy({}, { get: (t, k) => k === 'withSuccessHandler' ? f => mk(f, fail) : k === 'withFailureHandler' ? f => mk(ok, f) : (...a) => window.srv(k, a).then(ok, e => fail(new Error(e.message))) });
    window.google = { script: { run: mk(() => {}, () => {}) } };
    try { localStorage.setItem('skipHome', '1'); } catch (e) {}
  });
  await pg.goto('file://' + path.join(S, 'ui/index.html')); await pg.waitForTimeout(500);
  await pg.addStyleTag({ content: '*,*::before,*::after{animation:none!important;transition:none!important} .gd-hot{display:none}' });

  // 사진 찍기: 기준 요소(sel)의 영역, top/bottom 으로 자르기, hot: {stepId: 요소}
  async function shot(id, sel, hot, opt = {}) {
    const el = await pg.$(sel); await el.scrollIntoViewIfNeeded();
    let box = await el.boundingBox();
    if (opt.top) { const t = await (await pg.$(opt.top)).boundingBox(); const y = t.y - 10; box = { x: box.x, y, width: box.width, height: box.y + box.height - y }; }
    if (opt.bottom) { const t = await (await pg.$(opt.bottom)).boundingBox(); box.height = t.y + t.height + 14 - box.y; }
    if (opt.maxH) box.height = Math.min(box.height, opt.maxH);
    const sy = await pg.evaluate(() => window.scrollY), sx = await pg.evaluate(() => window.scrollX);
    const file = path.join(OUT, id + '.png');
    await pg.screenshot({ path: file, clip: { x: Math.max(0, box.x + sx), y: Math.max(0, box.y + sy), width: box.width, height: box.height }, fullPage: true });
    shots[id] = { file, w: Math.round(box.width), h: Math.round(box.height) };
    for (const [step, hs] of Object.entries(hot || {})) {
      const h = await pg.$(hs); if (!h) { console.log('no hot', step, hs); continue; }
      const hb = await h.boundingBox();
      hots[step] = [+(((hb.x + hb.width / 2) - box.x) / box.width * 100).toFixed(1), +(((hb.y + hb.height / 2) - box.y) / box.height * 100).toFixed(1)];
    }
    console.log('shot', id, shots[id].w + 'x' + shots[id].h);
  }

  // 1. 소요제기 입력
  await pg.evaluate(() => {
    document.querySelectorAll('.req-row').forEach(r => r.remove());
    [{ item: 'KERC00064', qty: 500, need: '2027-05-14', program: '787' }, { item: 'KERC00064', qty: 600, need: '2027-06-20', program: '787' },
     { item: 'KERN00278-014', qty: 64, need: '2027-05-10', program: 'AH-64E', ship: 'Air' }, { item: 'KERN00736-001', qty: 3, need: '2027-03-01', program: 'AH-64E' },
     { item: 'KERN02113-001', qty: 10, need: '2027-03-15', program: 'EFAB' }].forEach(v => addRow(v));
  });
  await shot('req', '#step1 .panel', { n1: '#btnClassify', n2: '#btnInputSheet' });
  // 2. 판정 결과
  await pg.evaluate(() => call('apiClassify').then(d => { state.results = d.results; prepResults(state.results); $('step1').classList.add('done'); render(); }));
  await pg.waitForTimeout(300);
  await shot('result', '#step2 .panel', { n3: '#tbl [data-pick="3"]', n4: '#modeBox', n5: '#confirmBox', a1: '#tbl .po-dl' });
  await pg.check('#pickOk');
  // 3. 구매내역서
  await pg.evaluate(() => { const f = document.getElementById('flowBar'); if (f) f.style.visibility = 'hidden'; });
  await shot('ptabs', '#purchaseCard .panel', { p1: '.doc-tab[data-doc="spot"]', p7: '#rates .rate-src' }, { bottom: '#remarks' });
  await pg.evaluate(() => { const f = document.getElementById('flowBar'); if (f) f.style.visibility = ''; });
  await pg.click('#btnSchedExample'); await pg.click('#moUp'); await pg.click('#moUp'); await pg.click('#moUp');
  await shot('gantt', '#schedSub', { p2: '#gbar-1', p3: '#moUp', p6: '#schedUnit' });
  await pg.fill('#confirmer', '홍길동 과장');
  await shot('pfoot', '#purchaseCard .panel', { p4: '#btnCreate' }, { top: '#fileName' });
  // 만드는 중 창 (14초째, 3단계)
  await pg.evaluate(() => { const m = makingStart('apiCreatePurchase'); m.t0 = Date.now() - 14200; mkTick(); });
  await shot('making', '#mk .mk-box', { p5: '#mkHide' });
  await pg.evaluate(() => makingEnd(mkCur));
  // 4. 견적요청서
  const names = ['88TB0110', 'A88TB0092-6', 'A5000V 2MIL'];
  const inputs = await pg.$$('#spotTbl input[data-f="name"]'); for (let i = 0; i < inputs.length; i++) await inputs[i].fill(names[i] || 'ITEM');
  await pg.fill('#rfqBiz', 'Boeing'); await pg.fill('#rfqKind', 'Process'); await pg.fill('#rfqDue', '2026-11-20'); await pg.dispatchEvent('#rfqDue', 'change');
  await shot('rfqbox', '#rfqBox', { r1: '#spotTbl input[data-f="name"]', r5: '#btnDirect' });
  await pg.click('#btnRfq'); await pg.waitForTimeout(200);
  await pg.click('.rp-opt[data-tpl="list"]'); await pg.fill('#rp2Kind', 'HW'); await pg.fill('#rp2Init', 'YWK'); await pg.dispatchEvent('#rp2Init', 'input');
  await shot('rfqpick', '#rfqPick .modal-box', { r2: '.rp-opt[data-tpl="list"]' });
  await pg.click('#rpCancel');
  // 5. 진행 중
  await pg.evaluate(() => showView('records')); await pg.waitForTimeout(300);
  await shot('records', '#recPanel', { c1: '.rec-ghead', c2: '[data-gdone]', c3: '.rec[data-id="R1"] [data-step="revise"]', c7: '.rec[data-id="R2"] .fx-meta', c10: '[data-po="R3"]' }, { top: '.rec-group' });
  await pg.evaluate(() => openRevise(recById('R1'))); await pg.waitForTimeout(200);
  await pg.fill('#rvTbl tr[data-i="0"] [data-f="qty"]', '5');
  await pg.click('#rvAdd'); await pg.fill('#rvTbl tr:last-child [data-f="item"]', 'KERN02200-001'); await pg.fill('#rvTbl tr:last-child [data-f="name"]', 'A88TB0100'); await pg.fill('#rvTbl tr:last-child [data-f="qty"]', '20');
  await shot('revise', '#rvModal .modal-box', { c4: '#rvSave' });
  await pg.click('#rvCancel');
  await pg.click('#recArchive'); await pg.waitForTimeout(400);
  await pg.check('[data-arc="A3"]');
  await shot('archive', '#recPanel', { c5: '[data-arcback="A1"]', c6: '#arcDel' }, { top: '#recPanel .rec-tools', maxH: 560 });
  await pg.click('#trashOpen'); await pg.waitForTimeout(400);
  await shot('trash', '.trash-box', { c8: '[data-untrash="T2"]' });
  await pg.click('#trashOpen'); await pg.waitForTimeout(200);
  await pg.click('#recArchive'); await pg.waitForTimeout(300);
  await shot('alerts', '#recPanel', { a2: '.al-chip', a3: '#dgToggle', m1: '[data-cf="R3"]' }, { top: '#recPanel .rec-tools', maxH: 760 });
  // 사이트 점검
  await pg.click('#btnCheck'); await pg.waitForTimeout(400);
  await shot('check', '#ckModal .modal-box', { c9: '#ckSum' });
  await pg.click('#ckClose');
  // 사업부 확인 메일
  await pg.click('[data-cf="R3"]'); await pg.waitForTimeout(400);
  await shot('confirm', '#cfModal .modal-box', { m2: '#cfPrev', m3: '#cfSend' });
  await pg.click('#cfCancel');
  await pg.click('[data-toggle="R1"]'); await pg.waitForTimeout(150);
  await shot('vlist', '.rec[data-id="R1"]', { q1: '[data-vendor="OMI"]', q2: '[data-qsheet]', q10: '.rec[data-id="R1"] .qprog' });
  // 메일 초안
  await pg.click('[data-draft="R1"]'); await pg.waitForTimeout(300);
  await pg.check('[data-vall="0"]'); await pg.click('[data-more="1"]'); await pg.check('[data-e="jun@cbol.example.com"]'); await pg.check('[data-vall="3"]');
  await shot('draft', '#draftModal .modal-box', { r3: '[data-more="1"]', r4: '#dmOk', r6: '#dmHintAll' });
  await pg.click('#dmCancel');
  // 견적 입력 표
  await pg.click('[data-vendor="OMI"]'); await pg.waitForTimeout(400);
  await pg.click('#vdExtra summary').catch(() => {}); await pg.waitForTimeout(150);
  await shot('qtable', '#vDrawer .drawer', { q3: '#vdBody tr[data-i="0"] [data-f="price"]', q8: '#vdAttach', q9: '#vdValid', q12: '#vdExtra summary', q13: '#vdBody [data-qsplit]' });
  await pg.click('#vdDone'); await pg.waitForTimeout(300);
  // 견적 종합
  await pg.click('[data-sum="R1"]'); await pg.waitForTimeout(300);
  await pg.selectOption('#qItems .q-item[data-i="0"] tr[data-v="CBOL"] [data-f="nbMode"]', 'adj'); await pg.waitForTimeout(100);
  await shot('qnb', '#qItems .q-item[data-i="0"]', { q4: '#qItems .q-item[data-i="0"] tr[data-v="CBOL"] input[type="radio"]', q7: '#qItems .q-item[data-i="0"] tr[data-v="CBOL"] [data-f="nbMode"]' });
  await shot('qdup', '#qItems .q-item[data-i="1"]', { q11: '#qItems .q-item[data-i="1"] .dup' });
  await shot('qbest', '#qItems', { q5: '#qBest .warn-txt', q14: '#qVendorTot h3' }, { top: '#qBest', maxH: 900 });
  await pg.click('#qBack'); await pg.waitForTimeout(200);
  await pg.click('[data-vendor="Aerovac"]'); await pg.waitForTimeout(400);
  await pg.click('[data-split="2"]'); await pg.waitForTimeout(200);
  await shot('qsplit', '#vDrawer .drawer', { q6: '#vdBody .vq-part' });
  await pg.click('[data-merge="2"]'); await pg.click('#vdDone'); await pg.waitForTimeout(300);
  // SPOT 신규 등록 (판정 결과의 등록 창)
  await pg.evaluate(() => { showView('new'); openRegister('KERN02200-001'); }); await pg.waitForTimeout(300);
  await shot('reg', '#regPanel', { i4: '#regNote' });
  await pg.click('#btnRegCancel');
  // 6. 자재 등록
  await pg.evaluate(() => {
    state.itRows = [{ item: 'KERN00278-020', priceMode: 'SPOT', supplier: 'Aerovac', uom: 'RL', shipFrom: 'Air', desc: 'FEP FILM 2MIL', qpl: 'AEROVAC' },
      { item: 'KERC09000', priceMode: 'LTA', supplier: 'PPG', price: '12.5', cur: 'USD', start: '2027-01-01', end: '2027-12-31', uom: 'KT', shipFrom: 'Ocean', leadtime: '10wks', desc: 'EPOXY PRIMER', qpl: 'PPG' },
      { item: 'KERC09001', priceMode: 'LTA', supplier: 'PPG', uom: 'KT', desc: 'TOPCOAT' }];
    showView('items');
  });
  await pg.click('#itSave'); await pg.waitForTimeout(300);
  await shot('items', '#viewItems .panel', { i1: '#itSheet', i2: '.row-err', i3: '#itSave' });
  // 7. 업체 등록
  await pg.evaluate(() => showView('vendors')); await pg.waitForTimeout(300);
  await pg.check('input[name="vnMode"][value="new"]'); await pg.fill('#vnName', 'Aerovac'); await pg.fill('#vnEmail', 'sales@aerovac.example.com'); await pg.fill('#vnSup', 'AEROVAC SYSTEMS');
  await pg.waitForTimeout(300);
  await shot('vendors', '#viewVendors .panel', { v1: 'input[name="vnMode"][value="new"]', v2: '[data-addto]', v3: '#vnTbl .vs' }, { maxH: 760 });
  // 8. 자재 정보
  await pg.evaluate(() => showView('master')); await pg.waitForTimeout(400);
  await pg.fill('#msFind', 'kerc'); await pg.waitForTimeout(100);
  await pg.click('[data-msopen="KERC00064"]'); await pg.waitForTimeout(100);
  await shot('master', '#viewMaster .panel', { k1: '#msFind', k2: '#msTbl .ms-trend' }, { maxH: 700 });
  await b.close();
  fs.writeFileSync(path.join(OUT, 'shots.json'), JSON.stringify({ shots, hots }));
})();
