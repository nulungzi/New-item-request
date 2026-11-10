// ===== 소요제기 판정 웹앱 · Apps Script =====
// 이 파일 전체를 Apps Script 편집기의 Code.gs 에 붙여넣습니다. (화면은 Index.html)
// 1부: 판정·계산 함수 / 2부: 시트·드라이브 입출력, 웹앱 함수(api*), 설정 점검(checkSetup)

// ================================================================
// 1부. 판정·계산 (Google 서비스를 쓰지 않는 순수 함수 — Node 테스트로 검증)
// ================================================================
/**
 * 순수 로직 (Google 서비스 호출 없음)
 */

// 자재정보 시트 열 이름 → 내부 키. 공백/대소문자/줄바꿈은 무시하고 비교한다.
var MASTER_COLUMNS = {
  item: ['Item', 'Item No.', '품번'],
  priceMode: ['계약 구분', 'Price Mode'],
  supplier: ['업체명', 'Supplier'],
  price: ['단가', 'Unit Price'],
  cur: ['CUR'],
  start: ['계약시작'],
  end: ['계약종료'],
  moq: ['MOQ'],
  uom: ['UOM'],
  desc: ['Description'],
  material: ['Material'],
  program: ['Program'],
  product: ['Product Name'],
  shipFrom: ['Ship From', 'Ship Form', 'Shipment'],
  leadtime: ['Leadtime', 'Lead Time', 'L/T'],
  shelfLife: ['Shelf Life'],
  taxCode: ['Tax Code']
};

/** 중복 없이 처음 나온 순서대로 (빈 값 제외) */
function uniq_(arr) {
  var out = [];
  (arr || []).forEach(function (x) { if (x !== '' && x !== null && x !== undefined && out.indexOf(x) < 0) out.push(x); });
  return out;
}

function headerKey_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, '').toLowerCase();
}

function itemKey_(s) {
  return String(s == null ? '' : s).trim().replace(/\s+/g, ' ').toUpperCase();
}

function toDate_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var m = String(v == null ? '' : v).trim().match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}

function fmtDate_(d) {
  if (!d) return '';
  var p = function (n) { return (n < 10 ? '0' : '') + n; };
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}

function toNumber_(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v == null ? '' : v).replace(/,/g, ''));
  return isNaN(n) ? null : n;
}

/** 2차원 배열(시트 값)에서 머리글 행을 찾아 계약 목록으로 변환. 머리글은 상위 10행 안 어디에 있어도 된다. */
function parseContracts_(values) {
  var hRow = -1, idx = {};
  for (var r = 0; r < Math.min(values.length, 10) && hRow < 0; r++) {
    var keys = values[r].map(headerKey_);
    if (keys.indexOf(headerKey_('Item')) >= 0 || keys.indexOf(headerKey_('품번')) >= 0) {
      hRow = r;
      Object.keys(MASTER_COLUMNS).forEach(function (k) {
        idx[k] = -1;
        MASTER_COLUMNS[k].forEach(function (alias) {
          var i = keys.indexOf(headerKey_(alias));
          if (i >= 0 && idx[k] < 0) idx[k] = i;
        });
      });
    }
  }
  if (hRow < 0) throw new Error('자재정보 시트에서 "Item" 머리글을 찾지 못했습니다.');

  var out = [];
  for (var i = hRow + 1; i < values.length; i++) {
    var row = values[i];
    var get = function (k) { return idx[k] >= 0 ? row[idx[k]] : ''; };
    var item = itemKey_(get('item'));
    if (!item) continue;
    out.push({
      row: i + 1,
      item: item,
      priceMode: String(get('priceMode')).trim(),
      supplier: String(get('supplier')).trim(),
      price: toNumber_(get('price')),
      cur: String(get('cur')).trim().toUpperCase(),
      start: toDate_(get('start')),
      end: toDate_(get('end')),
      moq: get('moq'),
      uom: String(get('uom')).trim(),
      desc: String(get('desc')).trim(),
      material: String(get('material')).trim(),
      program: String(get('program')).trim(),
      product: String(get('product')).trim(),
      shipFrom: String(get('shipFrom')).trim(),
      leadtime: String(get('leadtime')).trim(),
      shelfLife: String(get('shelfLife')).trim(),
      taxCode: get('taxCode') === '' ? '' : get('taxCode')
    });
  }
  return out;
}

/** 날짜 표기 통일: 2026-12-15, 2026.12.15, 2026. 12. 15., 2026/12/15, 12/15/2026 → 2026-12-15 */
function normDates_(line) {
  return line
    .replace(/(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})\.?/g, '$1-$2-$3')
    .replace(/\b(\d{1,2})\/(\d{1,2})\/(\d{4})\b/g, '$3-$1-$2');
}

/**
 * 소요제기 붙여넣기 파싱. 한 줄에 "Item  수량  Need-by-date  Program(선택)".
 * 엑셀에서 복사한 탭 구분, 쉼표, 공백 구분 모두 허용. 수량이 숫자가 아닌 줄(머리글 등)은 건너뛴다.
 * 날짜처럼 생긴 칸은 위치와 상관없이 Need-by-date로 본다.
 */
function parseRequest_(text) {
  var out = [], skipped = [];
  String(text || '').split(/\r?\n/).forEach(function (line, n) {
    if (!line.trim()) return;
    line = normDates_(line);
    var cells;
    if (line.indexOf('\t') >= 0) {
      cells = line.split('\t').map(function (c) { return c.trim(); });
    } else {
      // 공백 구분: 품번에 공백이 있을 수 있음 (예: "10-60754-279 42L 50 2027-06-10").
      // 날짜 바로 앞(날짜가 없으면 둘째 이후 첫) 숫자 칸을 수량으로 보고, 그 앞은 모두 품번으로 합친다.
      var t = line.trim().split(/\s*,\s*|\s+/);
      var isNum = function (x) { return /^\d[\d,]*(\.\d+)?$/.test(x); };
      var di = -1;
      t.forEach(function (x, i) { if (di < 0 && i > 0 && /^\d{4}-\d{1,2}-\d{1,2}$/.test(x)) di = i; });
      var k = -1;
      if (di > 1 && isNum(t[di - 1])) k = di - 1;
      for (var i = 1; k < 0 && i < t.length; i++) if (isNum(t[i])) k = i;
      cells = k > 0 ? [t.slice(0, k).join(' ')].concat(t.slice(k)) : t;
    }
    var qty = toNumber_(cells[1]);
    if (!cells[0] || qty === null) { skipped.push(n + 1); return; }
    var needBy = null, rest = [];
    cells.slice(2).forEach(function (c) {
      var d = !needBy && /^\d{4}-\d{1,2}-\d{1,2}$/.test(c) ? toDate_(c) : null;
      if (d) needBy = d; else if (c) rest.push(c);
    });
    out.push({ line: n + 1, item: itemKey_(cells[0]), qty: qty, needBy: needBy, program: rest.join(' ') });
  });
  return { requests: out, skipped: skipped };
}

/**
 * 화면 표 입력 → 요청 목록. rows: [{item, qty, needBy:'yyyy-mm-dd', program, shipFrom, taxCode}]
 * Item·수량이 없는 행은 건너뛰고 행 번호를 skipped 에 담는다.
 */
function parseRows_(rows) {
  var out = [], skipped = [];
  (rows || []).forEach(function (r, i) {
    var item = itemKey_(r.item), qty = toNumber_(r.qty);
    if (!item && (qty === null || qty === undefined)) return;
    if (!item || qty === null) { skipped.push(i + 1); return; }
    out.push({
      line: i + 1, item: item, qty: qty, needBy: toDate_(r.needBy),
      program: String(r.program || '').trim(),
      shipFrom: String(r.shipFrom || '').trim(),
      shelfLife: String(r.shelfLife || '').trim(),
      taxCode: String(r.taxCode == null ? '' : r.taxCode).trim()
    });
  });
  return { requests: out, skipped: skipped };
}

/** "10-12wks", "8 weeks", "30 days" → 최대 일수. 모르면 null */
function leadDays_(s) {
  var m = String(s || '').toLowerCase().match(/(\d+)(?:\s*[-~]\s*(\d+))?\s*(wk|week|w\b|day|d\b|주|일)/);
  if (!m) return null;
  var n = +(m[2] || m[1]);
  return /^(day|d|일)/.test(m[3]) ? n : n * 7;
}

/** 계약 품목 확인 메시지: MOQ 미달, Need-by 누락·지남, 리드타임 부족 (today = 작성일) */
function needByWarnings_(q, c, today) {
  var w = [];
  var moq = toNumber_(c.moq);
  if (moq && q.qty < moq) w.push('수량 ' + q.qty + ' < MOQ ' + moq);
  if (!q.needBy) { w.push('Need-by-date 없음 — 작성일 기준으로 판정'); return w; }
  if (q.needBy < today) { w.push('Need-by가 작성일보다 이전'); return w; }
  var ld = leadDays_(c.leadtime);
  if (ld && q.needBy - today < ld * 86400000) w.push('리드타임(' + c.leadtime + ') 부족');
  return w;
}

/**
 * 각 소요 품목을 그 품목의 Need-by-date(출고일) 기준으로 계약/Spot 판정.
 * Need-by-date가 없으면 today(작성일)로 판정. 유효 계약이 여럿이면 단가 낮은 순으로 후보를 준다.
 */
function classify_(requests, contracts, today) {
  var byItem = {};
  contracts.forEach(function (c) { (byItem[c.item] = byItem[c.item] || []).push(c); });
  return requests.map(function (q) {
    var d = q.needBy || today;
    var all = byItem[q.item] || [];
    var deals = all.filter(function (c) { return !isSpotMode_(c.priceMode); });   // SPOT 등록 행은 계약이 아님
    var valid = deals.filter(function (c) {
      return (!c.start || c.start <= d) && (!c.end || d <= c.end);
    }).sort(function (a, b) { return (a.price || 0) - (b.price || 0); });
    var res = { line: q.line, item: q.item, qty: q.qty, needBy: q.needBy, program: q.program,
      shipFrom: q.shipFrom || '', shelfLife: q.shelfLife || '', taxCode: q.taxCode || '', candidates: valid, warnings: [], registered: all.length > 0 };
    if (valid.length) {
      res.status = '계약';
      res.reason = valid.length > 1 ? '유효 계약 ' + valid.length + '건 — 업체 선택 필요' : '';
      res.warnings = needByWarnings_(q, valid[0], today);
    } else {
      res.status = 'Spot';
      if (!all.length) {
        res.reason = '계약 없음 · DB 미등록 자재';
      } else if (!deals.length) {
        var sups = [];
        all.forEach(function (c) { if (c.supplier && sups.indexOf(c.supplier) < 0) sups.push(c.supplier); });
        res.reason = 'Spot 자재 (DB 등록' + (sups.length ? ': ' + sups.join(', ') : '') + ')';
        res.desc = all[0].desc;
      } else {
        var last = deals.reduce(function (a, b) { return (b.end || 0) > (a.end || 0) ? b : a; });
        res.reason = last.end && last.end < d
          ? 'Need-by 시점 계약 만료 (' + last.supplier + ', ' + fmtDate_(last.end) + ' 종료)'
          : 'Need-by 시점 계약 시작 전 (' + last.supplier + ', ' + fmtDate_(last.start) + ' 시작)';
        res.desc = last.desc;
      }
    }
    return res;
  });
}

/** 사업부 담당자 표기 다듬기: "김정환   차장" → "김정환 차장" */
function normPerson_(s) { return String(s || '').replace(/\s+/g, ' ').trim().slice(0, 40); }
/** "이름 직급" 형식인지 (예: 김정환 차장, 남궁민수 대리) */
function isPersonFormat_(s) { return /^[가-힣]{2,5} [가-힣A-Za-z]{1,10}$/.test(normPerson_(s)); }

/** 구매내역서 파일 이름: 작업자가 적은 이름(다듬어서) 또는 "yymmdd 구매내역서_업체_업체" */
function purchaseFileName_(custom, base, suppliers) {
  var n = String(custom || '').replace(/[\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 150);
  return n || fmtDate_(base).replace(/-/g, '').slice(2) + ' 구매내역서_' + suppliers.join('_');
}

/* ------------------------------------------------------------------
 * 신규 자재 등록
 * ------------------------------------------------------------------ */
var PRICE_MODES = ['LTA', '3PP', 'SPOT'];
function isSpotMode_(m) { return /^spot$/i.test(String(m || '').trim()); }

/** 시트 머리글 → 내부 키 (모르는 머리글은 'x:머리글') */
function keyForHeader_(h) {
  var hk = headerKey_(h);
  if (!hk) return '';
  var found = '';
  Object.keys(MASTER_COLUMNS).forEach(function (k) {
    if (!found && MASTER_COLUMNS[k].some(function (a) { return headerKey_(a) === hk; })) found = k;
  });
  return found || 'x:' + String(h).trim();
}

/**
 * 등록 값 검사. rec: {item, priceMode, supplier, price, cur, start, end, moq, ... , 'x:머리글': 값}
 * 반환: {errors: [...], rec: 정리된 값(날짜는 Date, 숫자는 Number)}
 */
function validateNewItem_(rec, contracts) {
  var e = [], out = {};
  Object.keys(rec || {}).forEach(function (k) { out[k] = typeof rec[k] === 'string' ? rec[k].trim() : rec[k]; });
  out.item = itemKey_(rec.item);
  out.priceMode = String(rec.priceMode || '').trim().toUpperCase();
  out.supplier = String(rec.supplier || '').trim();
  out.cur = String(rec.cur || '').trim().toUpperCase();
  var isDeal = out.priceMode === 'LTA' || out.priceMode === '3PP';

  if (!out.item) e.push('Item을 입력해 주세요.');
  if (PRICE_MODES.indexOf(out.priceMode) < 0) e.push('계약 구분을 LTA, 3PP, SPOT 중에서 골라 주세요.');
  var hasPrice = rec.price !== undefined && rec.price !== null && String(rec.price).trim() !== '';
  out.price = hasPrice ? toNumber_(rec.price) : '';
  if (hasPrice && out.price === null) e.push('단가는 숫자로 입력해 주세요.');
  out.start = toDate_(rec.start) || '';
  out.end = toDate_(rec.end) || '';
  if (rec.start && !out.start) e.push('계약시작 날짜 형식을 확인해 주세요.');
  if (rec.end && !out.end) e.push('계약종료 날짜 형식을 확인해 주세요.');
  if (isDeal) {
    if (!out.supplier) e.push('LTA·3PP는 업체명이 필요합니다.');
    if (!hasPrice) e.push('LTA·3PP는 단가가 필요합니다.');
    if (!out.cur) e.push('LTA·3PP는 통화(CUR)가 필요합니다.');
    if (!out.start || !out.end) e.push('LTA·3PP는 계약시작과 계약종료가 필요합니다.');
  }
  if (out.start && out.end && out.start > out.end) e.push('계약종료가 계약시작보다 빠릅니다.');
  var moqN = toNumber_(rec.moq);
  if (moqN !== null && /^[\d,.\s]+$/.test(String(rec.moq))) out.moq = moqN;

  if (out.item && !e.length) {
    (contracts || []).forEach(function (c) {
      if (c.item !== out.item || c.supplier.toUpperCase() !== out.supplier.toUpperCase()) return;
      if (isDeal && !isSpotMode_(c.priceMode) && (!c.end || out.start <= c.end) && (!c.start || c.start <= out.end)) {
        e.push('같은 업체·기간에 겹치는 계약이 이미 있습니다 (자재정보 ' + c.row + '행).');
      }
      if (!isDeal && isSpotMode_(c.priceMode)) e.push('같은 업체로 이미 SPOT 등록되어 있습니다 (자재정보 ' + c.row + '행).');
    });
  }
  return { errors: e, rec: out };
}

/** 머리글 목록(시트 열 순서)과 정리된 값으로 시트 한 행을 만든다. 머리글 없는 열은 '' */
function newItemRow_(headers, rec) {
  return headers.map(function (h) {
    var k = keyForHeader_(h);
    if (!k) return '';
    var v = rec[k];
    return v === undefined || v === null ? '' : v;
  });
}

/** 화면 Item 안내용 목록: [Item, Description, Ship From, 계약 구분들('LTA/SPOT'), Shelf Life] (Item 순) */
function itemIndex_(contracts) {
  var items = {};
  contracts.forEach(function (c) {
    var x = items[c.item] = items[c.item] || { desc: '', ship: '', shelf: '', modes: [] };
    if (!x.desc && c.desc) x.desc = c.desc;
    if (!x.ship && c.shipFrom) x.ship = c.shipFrom;
    if (!x.shelf && c.shelfLife) x.shelf = c.shelfLife;
    x.modes = uniq_(x.modes.concat([c.priceMode.toUpperCase()]));
  });
  return Object.keys(items).sort().map(function (k) {
    var x = items[k];
    return [k, x.desc, x.ship, x.modes.join('/'), x.shelf];
  });
}

/** 구매내역서에 들어갈 업체·통화별 요약 (L/I 순서 = 업체 등장 순서). */
function summarize_(lines) {
  var groups = [], map = {};
  lines.forEach(function (l) {
    var k = l.supplier + '|' + l.cur;
    if (!map[k]) {
      map[k] = { supplier: l.supplier, cur: l.cur, priceMode: l.priceMode, count: 0, qty: 0, amount: 0, materials: {} };
      groups.push(map[k]);
    }
    var g = map[k];
    g.count++;
    g.qty += l.qty;
    g.amount += l.qty * (l.price || 0);
    if (l.material) g.materials[l.material] = (g.materials[l.material] || 0) + 1;
  });
  groups.forEach(function (g) {
    var mats = Object.keys(g.materials).sort(function (a, b) { return g.materials[b] - g.materials[a]; });
    g.mainMaterial = mats[0] || '';
  });
  return groups;
}

/* ------------------------------------------------------------------
 * 구매내역서 작성 계획
 * 템플릿(templates/구매내역서_템플릿.xlsx)을 복사한 뒤 아래 작업 목록(ops)을 그대로 적용한다.
 * Apps Script(applyPlan_)와 테스트용 파이썬 도구가 같은 목록을 쓰므로 결과가 같다.
 *
 * op 종류
 *   fmt    {s, r, from, c1, c2}       _proto 탭 from행의 c1..c2 서식을 s탭 r행에 복사
 *   cellFmt{s, r, c, fs, fr, fc}      fs탭 (fr,fc) 셀 서식을 (r,c)에 복사
 *   set    {s, r, c, v}               값. '='로 시작하면 수식, {date:'yyyy-mm-dd'}는 날짜
 *   nf     {s, r, c, f}               표시 형식
 *   bg     {s, r, c, color}
 *   merge  {s, r1, c1, r2, c2} / unmerge (같은 인자)
 *   height {s, r, pt}
 *   bottom {s, r, c1, c2}             아래 테두리(얇은 선)
 *   list   {s, r1, r2, c, items}      드롭다운
 *   freeze {s, rows, cols}
 * ------------------------------------------------------------------ */
var TPL = {
  SUMMARY: '구매요약', DETAIL: '구매세부내역', SCHEDULE: '구매일정', CONFIRM: '사업부 담당 확인', PROTO: '_proto',
  // _proto 탭 행 번호
  P_S_VENDOR: 1, P_S_CUR: 2, P_S_TOTAL: 3, P_S_NOTE: 4,
  P_D_DATE: 11, P_D_RATE: 12, P_D_TTL: 13, P_D_TTLKRW: 14, P_D_TOTAL: 15, P_D_SPACER: 16, P_D_HEAD: 17, P_D_DATA: 18,
  FILL: '#fff2cc'
};

/**
 * 구매일정 막대. 칸 단위는 반 달(0 = 작성일 달 전반), 12개월 = 24칸.
 * 사용자가 화면에서 줄마다 끌어서 기간을 고른다. 고르지 않은 줄은 {start:0, len:0} (시트에서 빈 줄).
 * SCHEDULE_EXAMPLE 은 화면의 "예시로 채우기"와, 화면이 일정을 보내지 않았을 때(이전 화면) 쓴다.
 */
var SCHEDULE_TASKS = ['현 재고+Open PO', 'PO Coverage', '품의 및 PR/PO', '업체생산 L/T', '운송, 통관, 검사'];
var SCHEDULE_EXAMPLE = [{ start: 0, len: 9 }, { start: 8, len: 10 }, { start: 0, len: 1 }, { start: 1, len: 6 }, { start: 7, len: 1 }];
var SCHEDULE_SLOTS = 24;
var SCHEDULE_COLORS = ['#cccccc', '#fff2cc', '#fff2cc', '#fff2cc', '#fff2cc'];

/** 화면에서 받은 막대 값 검사·보정. len 0 = 비어 있음 */
function normSchedule_(sched) {
  if (!sched) sched = SCHEDULE_EXAMPLE;
  return SCHEDULE_TASKS.map(function (t, i) {
    var x = sched[i] || {};
    var start = Math.round(Number(x.start)), len = Math.round(Number(x.len));
    if (!(len > 0) || isNaN(start)) return { start: 0, len: 0 };
    start = Math.max(0, Math.min(SCHEDULE_SLOTS - 1, start));
    return { start: start, len: Math.max(1, Math.min(SCHEDULE_SLOTS - start, len)) };
  });
}

var CUR_FMT = { USD: '[$$]#,##0.00', EUR: '[$€]#,##0.00', SGD: '"S$"#,##0.00', JPY: '[$¥]#,##0', KRW: '[$₩-412]#,##0' };
function curFmt_(cur) { return CUR_FMT[cur] || '#,##0.00'; }

function colA_(n) { var s = ''; while (n) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }

function rateNote_(curs, rates, base) {
  return '※ 적용환율: ' + curs.map(function (c) {
    var v = Number(rates[c]);
    return 'KRW ' + v.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '/' + c;
  }).join(', ') + ' (' + fmtDate_(base).replace(/-/g, '/') + ' 기준)';
}

/**
 * lines: [{item, qty, needBy(Date), supplier, cur, price, priceMode, material, desc, product, shipFrom, leadtime,
 *          uom, shelfLife, taxCode, remark}]
 * curs: KRW 제외 통화 목록(등장 순), rates: {USD: 1339.3}, base: 작성일(Date)
 * schedule: 구매일정 막대 [{start, len}] × 5 (반 달 단위, len 0 = 빈 줄, 없으면 예시)
 * extra: {title: 구매내역서 제목(사업부 담당 확인 B2), confirmer: '홍길동 과장'(A2)}
 */
function buildPurchasePlan_(lines, curs, rates, base, schedule, extra) {
  extra = extra || {};
  var ops = [];
  var S = TPL.SUMMARY, D = TPL.DETAIL;
  var add = function (o) { ops.push(o); };
  var set = function (s, r, c, v) { add({ op: 'set', s: s, r: r, c: c, v: v }); };
  var nf = function (s, r, c, f) { add({ op: 'nf', s: s, r: r, c: c, f: f }); };
  var hasKRW = lines.some(function (l) { return l.cur === 'KRW'; });

  /* ---------- 구매세부내역 ---------- */
  // 1~2행: 작성일, 통화별 환율 (O1 날짜, P1.. 통화, P2.. 환율)
  add({ op: 'fmt', s: D, r: 1, from: TPL.P_D_DATE, c1: 2, c2: 20 });
  add({ op: 'fmt', s: D, r: 2, from: TPL.P_D_RATE, c1: 2, c2: 20 });
  set(D, 1, 15, { date: fmtDate_(base) });
  var rateCell = {};
  curs.forEach(function (c, i) {
    add({ op: 'cellFmt', s: D, r: 1, c: 16 + i, fs: TPL.PROTO, fr: TPL.P_D_DATE, fc: 16 });
    add({ op: 'cellFmt', s: D, r: 2, c: 16 + i, fs: TPL.PROTO, fr: TPL.P_D_RATE, fc: 16 });
    set(D, 1, 16 + i, c);
    set(D, 2, 16 + i, Number(rates[c]));
    rateCell[c] = '$' + colA_(16 + i) + '$2';
  });

  var head = 6 + curs.length, first = head + 1, last = head + lines.length;
  var I = '$I$' + first + ':$I$' + last, K = '$K$' + first + ':$K$' + last;
  var r = 3, parts = [];
  curs.forEach(function (c) {
    add({ op: 'fmt', s: D, r: r, from: TPL.P_D_TTL, c1: 2, c2: 20 });
    set(D, r, 10, 'TTL(' + c + ')');
    set(D, r, 11, '=SUMIF(' + I + ',"' + c + '",' + K + ')');
    nf(D, r, 11, curFmt_(c));
    parts.push('K' + r + '*' + rateCell[c]);
    r++;
  });
  add({ op: 'fmt', s: D, r: r, from: TPL.P_D_TTLKRW, c1: 2, c2: 20 });
  set(D, r, 10, 'TTL(KRW)');
  set(D, r, 11, '=SUMIF(' + I + ',"KRW",' + K + ')');
  parts.push('K' + r);
  r++;
  var note = rateNote_(curs, rates, base);
  add({ op: 'fmt', s: D, r: r, from: TPL.P_D_TOTAL, c1: 2, c2: 20 });
  set(D, r, 10, 'Total');
  set(D, r, 11, '=' + parts.join('+'));
  set(D, r, 12, note);
  var detailTotalRow = r;
  r++;
  add({ op: 'fmt', s: D, r: r, from: TPL.P_D_SPACER, c1: 2, c2: 20 });
  add({ op: 'height', s: D, r: r, pt: 8.25 });

  add({ op: 'fmt', s: D, r: head, from: TPL.P_D_HEAD, c1: 2, c2: 20 });
  add({ op: 'height', s: D, r: head, pt: 39.75 });
  ['L/I', 'Program', 'Item No.', 'Material', 'Description', 'Product\n Name', 'Qty', 'CUR', 'Unit Price', 'Amount',
    'Ship From', 'Leadtime', 'UoM', 'Shelf Life', 'Supplier', 'Tax\nCode', 'Price \n Mode', '비고', 'Need-by\nDate']
    .forEach(function (h, i) { set(D, head, 2 + i, h); });

  lines.forEach(function (l, i) {
    var n = first + i;
    add({ op: 'fmt', s: D, r: n, from: TPL.P_D_DATA, c1: 2, c2: 20 });
    add({ op: 'height', s: D, r: n, pt: 15 });
    var row = [i + 1, l.program || '', l.item, l.material, l.desc, l.product, l.qty, l.cur, l.price, '=J' + n + '*H' + n,
      l.shipFrom, l.leadtime, l.uom, l.shelfLife, l.supplier, l.taxCode, l.priceMode, l.remark || '',
      l.needBy ? { date: fmtDate_(l.needBy) } : ''];
    row.forEach(function (v, j) { if (v !== '' && v !== null && v !== undefined) set(D, n, 2 + j, v); });
    nf(D, n, 10, curFmt_(l.cur));
    nf(D, n, 11, curFmt_(l.cur));
    nf(D, n, 20, 'yyyy-mm-dd');
    // 담당자 입력 칸
    if (!l.program) add({ op: 'bg', s: D, r: n, c: 3, color: TPL.FILL });
    if (!l.shipFrom) add({ op: 'bg', s: D, r: n, c: 12, color: TPL.FILL });
    if (!l.leadtime) add({ op: 'bg', s: D, r: n, c: 13, color: TPL.FILL });
  });
  add({ op: 'bottom', s: D, r: last, c1: 2, c2: 20 });
  add({ op: 'list', s: D, r1: first, r2: last, c: 12, items: ['Air', 'Ocean'] });
  add({ op: 'freeze', s: D, rows: head, cols: 4 });

  /* ---------- 구매요약 ---------- */
  var groups = summarize_(lines);
  set(S, 3, 8, curs.length + (hasKRW ? 1 : 0) === 1 ? '구매금액(' + (curs[0] || 'KRW') + ')' : '구매금액');
  var DP = "'" + D + "'!";
  var P = DP + '$P$' + first + ':$P$' + last, IC = DP + I, H = DP + '$H$' + first + ':$H$' + last, KC = DP + K;
  r = 4;
  groups.forEach(function (g, i) {
    add({ op: 'fmt', s: S, r: r, from: TPL.P_S_VENDOR, c1: 2, c2: 12 });
    add({ op: 'height', s: S, r: r, pt: 24.75 });
    var crit = P + ',$L' + r + ',' + IC + ',$G' + r;
    set(S, r, 2, i + 1);
    set(S, r, 3, g.supplier);
    set(S, r, 4, g.mainMaterial);
    set(S, r, 5, '=COUNTIFS(' + crit + ')');
    set(S, r, 6, '=SUMIFS(' + H + ',' + crit + ')');
    set(S, r, 7, g.cur);
    set(S, r, 8, '=SUMIFS(' + KC + ',' + crit + ')');
    nf(S, r, 8, curFmt_(g.cur));
    set(S, r, 9, g.priceMode);
    set(S, r, 12, g.supplier);   // 수식용 도우미(회색 글자). C열은 "PPG(제작사, US)"처럼 자유롭게 고쳐도 됨
    r++;
  });
  var gEnd = r - 1, krw = [];
  var allCurs = curs.concat(hasKRW ? ['KRW'] : []);
  allCurs.forEach(function (c) {
    var rng = function (col) { return '$' + col + '$4:$' + col + '$' + gEnd; };
    add({ op: 'fmt', s: S, r: r, from: TPL.P_S_CUR, c1: 2, c2: 12 });
    add({ op: 'height', s: S, r: r, pt: 24.75 });
    add({ op: 'merge', s: S, r1: r, c1: 2, r2: r, c2: 4 });
    add({ op: 'merge', s: S, r1: r, c1: 8, r2: r, c2: 9 });
    set(S, r, 2, 'CUR 별 합계');
    set(S, r, 5, '=SUMIF(' + rng('G') + ',G' + r + ',' + rng('E') + ')');
    set(S, r, 6, '=SUMIF(' + rng('G') + ',G' + r + ',' + rng('F') + ')');
    set(S, r, 7, c);
    set(S, r, 8, '=SUMIF(' + rng('G') + ',G' + r + ',' + rng('H') + ')');
    nf(S, r, 8, curFmt_(c));
    krw.push(c === 'KRW' ? 'H' + r : 'H' + r + '*' + DP + rateCell[c]);
    r++;
  });
  add({ op: 'fmt', s: S, r: r, from: TPL.P_S_TOTAL, c1: 2, c2: 12 });
  add({ op: 'height', s: S, r: r, pt: 24.75 });
  add({ op: 'merge', s: S, r1: r, c1: 2, r2: r, c2: 7 });
  add({ op: 'merge', s: S, r1: r, c1: 8, r2: r, c2: 9 });
  set(S, r, 2, '총 합계 (KRW 표기)');
  set(S, r, 8, '=' + krw.join('+'));
  r++;
  add({ op: 'fmt', s: S, r: r, from: TPL.P_S_NOTE, c1: 2, c2: 12 });
  add({ op: 'merge', s: S, r1: r, c1: 6, r2: r, c2: 9 });
  set(S, r, 6, '=' + DP + 'L' + detailTotalRow);

  /* ---------- 구매일정: 작성일 달부터 12개월 머리글 (E~AB열, 달마다 2칸) ---------- */
  var G = TPL.SCHEDULE;
  add({ op: 'unmerge', s: G, r1: 4, c1: 5, r2: 4, c2: 28 });
  for (var c = 5; c <= 28; c++) set(G, 4, c, '');
  var y = base.getFullYear(), m = base.getMonth() + 1, segStart = 5;
  for (var k = 0; k < 12; k++) {
    var col = 5 + k * 2;
    set(G, 5, col, m);
    var endOfYear = m === 12 || k === 11;
    if (endOfYear) {
      add({ op: 'cellFmt', s: G, r: 4, c: segStart, fs: G, fr: 4, fc: 5 });
      set(G, 4, segStart, y + '년');
      add({ op: 'merge', s: G, r1: 4, c1: segStart, r2: 4, c2: col + 1 });
      segStart = col + 2;
    }
    m++;
    if (m > 12) { m = 1; y++; }
  }

  // 막대: 6~10행, E(5)~AB(28)열. L/T(D열)는 막대 길이로 계산
  normSchedule_(schedule).forEach(function (b, i) {
    var row = 6 + i;
    set(G, row, 4, b.len ? (b.len / 2).toFixed(1) + 'M' : '');
    for (var k = 0; k < SCHEDULE_SLOTS; k++) {
      add({ op: 'bg', s: G, r: row, c: 5 + k, color: k >= b.start && k < b.start + b.len ? SCHEDULE_COLORS[i] : '#ffffff' });
    }
  });

  /* ---------- 사업부 담당 확인 ---------- */
  // A2 사업부 담당자(이름 직급), B2 제목 = 구매내역서 제목, C2 확인 날짜는 담당자가 직접 입력(노란 칸)
  var who = normPerson_(extra.confirmer);
  if (who) set(TPL.CONFIRM, 2, 1, who);
  else add({ op: 'bg', s: TPL.CONFIRM, r: 2, c: 1, color: TPL.FILL });
  set(TPL.CONFIRM, 1, 2, '구매내역서 제목');
  set(TPL.CONFIRM, 2, 2, extra.title || '');
  add({ op: 'bg', s: TPL.CONFIRM, r: 2, c: 3, color: TPL.FILL });

  return ops;
}

// ================================================================
// 2부. 시트·드라이브 입출력과 웹앱 함수
// ================================================================
/**
 * 소요제기 판정 웹앱 — 계약(LTA/3PP) 품목은 구매내역서를 만들고, Spot 품목은 견적 대상으로 분류한다.
 * 판정·계산은 1부(순수 함수), 여기 2부는 시트·드라이브 입출력만 맡는다.
 *
 * 스크립트 속성
 *   TEMPLATE_ID      (필수) 구매내역서 템플릿 구글 시트 ID — templates/구매내역서_템플릿.xlsx 를 드라이브에 올려 구글 시트로 저장
 *   OUTPUT_FOLDER_ID (권장) 생성한 구매내역서를 넣을 드라이브 폴더 ID
 *   SHEET_ID         (선택) 자재정보가 있는 스프레드시트 ID (없으면 이 스크립트가 붙은 시트)
 */
var SHEET_MASTER = '자재정보';
var SHEET_RATE = '환율';

/* ================================================================
 * 설정·시트 접근
 * ================================================================ */
function prop_(key) { return PropertiesService.getScriptProperties().getProperty(key); }

function db_() {
  var id = prop_('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function masterSheet_() {
  var sh = db_().getSheetByName(SHEET_MASTER);
  if (!sh) throw new Error('"' + SHEET_MASTER + '" 시트가 없습니다.');
  return sh;
}

function masterValues_() { return masterSheet_().getDataRange().getValues(); }

function loadContracts_() { return parseContracts_(masterValues_()); }

/** 자재정보 머리글 위치 {row(1부터), headers} */
function masterHeader_(values) {
  for (var r = 0; r < Math.min(values.length, 10); r++) {
    if (values[r].some(function (h) { return keyForHeader_(h) === 'item'; })) return { row: r + 1, headers: values[r] };
  }
  throw new Error('자재정보 시트에서 "Item" 머리글을 찾지 못했습니다.');
}

/** 신규 등록 양식에 쓸 머리글 목록 (시트 열 순서) */
function masterFields_(values) {
  return masterHeader_(values).headers
    .map(function (h) { return { header: String(h).replace(/\s+/g, ' ').trim(), key: keyForHeader_(h) }; })
    .filter(function (f) { return f.key; });
}

/** 환율 시트: A열 통화, B열 환율(KRW), C열 수정일. 없으면 만든다. */
function rateSheet_() {
  var ss = db_(), sh = ss.getSheetByName(SHEET_RATE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_RATE);
    sh.getRange(1, 1, 1, 3).setValues([['CUR', 'KRW 환율', '수정일']]);
  }
  return sh;
}

function loadRates_() {
  var rates = {};
  rateSheet_().getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[0]) rates[String(r[0]).trim().toUpperCase()] = r[1];
  });
  return rates;
}

function saveRates_(rates) {
  var sh = rateSheet_(), rows = sh.getDataRange().getValues(), now = new Date();
  Object.keys(rates).forEach(function (cur) {
    var i = -1;
    rows.forEach(function (r, n) { if (i < 0 && n > 0 && String(r[0]).toUpperCase() === cur) i = n; });
    if (i < 0) { rows.push([cur, rates[cur], now]); sh.appendRow([cur, rates[cur], now]); }
    else sh.getRange(i + 1, 2, 1, 2).setValues([[rates[cur], now]]);
  });
}

/** google.script.run 으로 보낼 수 있게 Date → 'yyyy-mm-dd' */
function plainDates_(o) {
  var out = {};
  Object.keys(o).forEach(function (k) { out[k] = Object.prototype.toString.call(o[k]) === '[object Date]' ? fmtDate_(o[k]) : o[k]; });
  return out;
}

/* ================================================================
 * 웹앱
 * ================================================================ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('소요제기 판정')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 화면이 처음 열릴 때: 작성일, 저장된 환율, 구매일정 설정, Item 안내, 신규 등록 양식 */
function apiInit() {
  var values = masterValues_();
  return {
    today: fmtDate_(new Date()),
    rates: loadRates_(),
    schedule: { tasks: SCHEDULE_TASKS, example: SCHEDULE_EXAMPLE, slots: SCHEDULE_SLOTS },
    items: itemIndex_(parseContracts_(values)),   // [Item, Description, Ship From, 계약 구분들, Shelf Life]
    fields: masterFields_(values)
  };
}

/** 소요제기 판정. input: 화면 표(행 배열) 또는 붙여넣기 글. baseDateStr: 작성일 (계약 판정은 품목별 Need-by 기준) */
function apiClassify(input, baseDateStr) {
  var parsed = Array.isArray(input) ? parseRows_(input) : parseRequest_(input);
  var base = toDate_(baseDateStr) || toDate_(new Date());
  var results = classify_(parsed.requests, loadContracts_(), base).map(function (r) {
    r.needBy = fmtDate_(r.needBy);
    r.candidates = r.candidates.map(plainDates_);
    return r;
  });
  return { results: results, skipped: parsed.skipped, rates: loadRates_() };
}

/** 신규 자재를 자재정보 시트 맨 아래에 추가. rec 키는 masterFields_ 의 key. 머리글 없는 맨 앞 번호 열은 윗줄 + 1 */
function apiRegisterItem(rec) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = masterSheet_(), values = sh.getDataRange().getValues(), head = masterHeader_(values);
    var v = validateNewItem_(rec, parseContracts_(values));
    if (v.errors.length) throw new Error(v.errors.join('\n'));

    var row = newItemRow_(head.headers, v.rec), n = row.length, last = sh.getLastRow(), at = last + 1;
    var prev = values[last - 1] || [];
    if (!String(head.headers[0]).trim() && last > head.row && typeof prev[0] === 'number') row[0] = prev[0] + 1;

    var target = sh.getRange(at, 1, 1, n);
    if (last > head.row) sh.getRange(last, 1, 1, n).copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    head.headers.forEach(function (h, i) {
      var k = keyForHeader_(h);
      if (k === 'item') sh.getRange(at, i + 1).setNumberFormat('@');
      if (k === 'start' || k === 'end') sh.getRange(at, i + 1).setNumberFormat('yyyy-mm-dd');
    });
    target.setValues([row]);
    return { row: at, item: v.rec.item, desc: v.rec.desc || '', ship: v.rec.shipFrom || '', shelf: v.rec.shelfLife || '', mode: v.rec.priceMode };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 구매내역서 생성.
 * payload = { baseDate, rates: {USD: 1339.3}, remarks: {업체명: '비고'},
 *             lines: [{item, qty, needBy, program, shelfLife, shipFrom, taxCode, contractRow}],
 *             schedule: [{start, len}] × 5 (구매일정, 반 달 단위),
 *             name: 구매내역서 제목 (비우면 자동), confirmer: 사업부 담당자 '홍길동 과장' }
 * 단가는 contractRow 로 자재정보를 다시 읽어 서버에서 확정한다(화면 값은 믿지 않음).
 */
function apiCreatePurchase(payload) {
  var lines = purchaseLines_(payload);
  var curs = uniq_(lines.map(function (l) { return l.cur; })).filter(function (c) { return c !== 'KRW'; });
  var rates = checkedRates_(payload.rates, curs);
  saveRates_(rates);

  var base = toDate_(payload.baseDate) || toDate_(new Date());
  var name = purchaseFileName_(payload.name, base, uniq_(lines.map(function (l) { return l.supplier; })));
  var ss = copyTemplate_(name);
  applyPlan_(ss, buildPurchasePlan_(lines, curs, rates, base, payload.schedule, { title: name, confirmer: payload.confirmer }));
  var proto = ss.getSheetByName(TPL.PROTO);
  if (proto) ss.deleteSheet(proto);
  ss.setActiveSheet(ss.getSheetByName(TPL.SUMMARY));
  return { url: ss.getUrl(), name: name, id: ss.getId() };
}

/** 화면에서 받은 줄 + 자재정보 계약 행 → 구매내역서 한 줄 */
function purchaseLines_(payload) {
  var byRow = {};
  loadContracts_().forEach(function (c) { byRow[c.row] = c; });
  var lines = (payload.lines || []).map(function (l) {
    var c = byRow[l.contractRow];
    if (!c || c.item !== l.item) throw new Error(l.item + ': 자재정보가 바뀌었습니다. 다시 판정해 주세요.');
    var o = {};
    Object.keys(c).forEach(function (k) { o[k] = c[k]; });
    o.qty = Number(l.qty);
    o.program = l.program || '';                       // 비우면 시트에서 노란 칸
    if (l.shipFrom) o.shipFrom = l.shipFrom;           // 화면에 적으면 자재정보 값 대신
    if (l.shelfLife) o.shelfLife = l.shelfLife;
    if (l.taxCode !== undefined && l.taxCode !== '') o.taxCode = isNaN(+l.taxCode) ? l.taxCode : +l.taxCode;
    o.needBy = toDate_(l.needBy);
    o.remark = (payload.remarks || {})[c.supplier] || '';
    return o;
  });
  if (!lines.length) throw new Error('구매내역서에 넣을 계약 품목이 없습니다.');
  return lines;
}

/** 필요한 통화의 환율이 모두 양수인지 확인 */
function checkedRates_(input, curs) {
  var rates = {};
  Object.keys(input || {}).forEach(function (k) {
    var n = Number(input[k]);
    if (!(n > 0)) throw new Error(k + ' 환율을 입력해 주세요.');
    rates[k] = n;
  });
  curs.forEach(function (c) { if (!rates[c]) throw new Error(c + ' 환율을 입력해 주세요.'); });
  return rates;
}

/** 템플릿 시트를 복사해 새 스프레드시트를 연다 */
function copyTemplate_(name) {
  var templateId = prop_('TEMPLATE_ID');
  if (!templateId) throw new Error('스크립트 속성 TEMPLATE_ID(구매내역서 템플릿 시트 ID)를 설정해 주세요. README 참고.');
  var folderId = prop_('OUTPUT_FOLDER_ID'), tpl = DriveApp.getFileById(templateId);
  var file = folderId ? tpl.makeCopy(name, DriveApp.getFolderById(folderId)) : tpl.makeCopy(name);
  return SpreadsheetApp.openById(file.getId());
}

/* ================================================================
 * 시트 쓰기: buildPurchasePlan_ 의 작업 목록 적용
 * ================================================================ */
function applyPlan_(ss, ops) {
  var cache = {};
  var sh = function (n) {
    if (!cache[n]) {
      cache[n] = ss.getSheetByName(n);
      if (!cache[n]) throw new Error('템플릿에 "' + n + '" 탭이 없습니다.');
    }
    return cache[n];
  };
  var FORMAT = SpreadsheetApp.CopyPasteType.PASTE_FORMAT;
  var area = function (o) { return sh(o.s).getRange(o.r1, o.c1, o.r2 - o.r1 + 1, o.c2 - o.c1 + 1); };
  ops.forEach(function (o) {
    switch (o.op) {
      case 'fmt': sh(TPL.PROTO).getRange(o.from, o.c1, 1, o.c2 - o.c1 + 1).copyTo(sh(o.s).getRange(o.r, o.c1), FORMAT, false); break;
      case 'cellFmt': sh(o.fs).getRange(o.fr, o.fc).copyTo(sh(o.s).getRange(o.r, o.c), FORMAT, false); break;
      case 'set': setCell_(sh(o.s).getRange(o.r, o.c), o.v); break;
      case 'nf': sh(o.s).getRange(o.r, o.c).setNumberFormat(o.f); break;
      case 'bg': sh(o.s).getRange(o.r, o.c).setBackground(o.color); break;
      case 'merge': area(o).merge(); break;
      case 'unmerge': area(o).breakApart(); break;
      case 'height': sh(o.s).setRowHeight(o.r, Math.round(o.pt * 4 / 3)); break;   // pt → px
      case 'bottom':
        sh(o.s).getRange(o.r, o.c1, 1, o.c2 - o.c1 + 1)
          .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
        break;
      case 'list':
        sh(o.s).getRange(o.r1, o.c, o.r2 - o.r1 + 1, 1).setDataValidation(
          SpreadsheetApp.newDataValidation().requireValueInList(o.items, true).setAllowInvalid(true).build());
        break;
      case 'freeze': sh(o.s).setFrozenRows(o.rows); sh(o.s).setFrozenColumns(o.cols); break;
      default: throw new Error('알 수 없는 작업: ' + o.op);
    }
  });
}

/** 값 넣기: {date:'yyyy-mm-dd'} 는 날짜, '='로 시작하면 수식 */
function setCell_(cell, v) {
  if (v && typeof v === 'object' && v.date) cell.setValue(toDate_(v.date));
  else if (typeof v === 'string' && v.charAt(0) === '=') cell.setFormula(v);
  else cell.setValue(v);
}

/* ================================================================
 * 미리보기: 만든 구매내역서 탭을 HTML로
 * 1순위 구글 시트 "웹페이지(zip)" 내보내기, 실패하면 값·배경·병합만 그린 간단한 표
 * ================================================================ */
function apiPreview(id) {
  SpreadsheetApp.flush();
  var exported = {};
  try { exported = exportHtml_(id); } catch (e) { console.warn('zip 내보내기 실패: ' + e); }
  var ss = SpreadsheetApp.openById(id);
  return [TPL.SUMMARY, TPL.DETAIL, TPL.SCHEDULE, TPL.CONFIRM].map(function (name) {
    if (exported[name]) return { name: name, html: exported[name], mode: 'export' };
    var sh = ss.getSheetByName(name);
    return { name: name, html: sh ? simpleHtml_(sh) : '<p>탭이 없습니다.</p>', mode: 'simple' };
  });
}

function exportHtml_(id) {
  var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + id + '/export?format=zip', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());
  var css = '', pages = {};
  Utilities.unzip(resp.getBlob().setContentType('application/zip')).forEach(function (f) {
    var n = f.getName();
    if (/\.css$/i.test(n)) css += f.getDataAsString('UTF-8');
    else if (/\.html$/i.test(n)) pages[n.replace(/^.*\//, '').replace(/\.html$/i, '')] = f.getDataAsString('UTF-8');
  });
  var extra = 'body{margin:8px;background:#fff}.row-headers-background,.column-headers-background,' +
    '.row-header-wrapper,.freezebar-cell,.freezebar-origin-ltr{display:none!important}';
  Object.keys(pages).forEach(function (k) {
    pages[k] = pages[k].replace(/<link[^>]+sheet\.css[^>]*>/i, '').replace(/<\/head>/i, '<style>' + css + extra + '</style></head>');
  });
  return pages;
}

function esc_(s) {
  return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
}

function simpleHtml_(sh) {
  var nr = Math.min(sh.getLastRow(), 80), nc = Math.min(sh.getLastColumn(), 30);
  if (!nr || !nc) return '<p>내용 없음</p>';
  var rg = sh.getRange(1, 1, nr, nc);
  var v = rg.getDisplayValues(), bg = rg.getBackgrounds(), fw = rg.getFontWeights(), fc = rg.getFontColors(),
    al = rg.getHorizontalAlignments(), fs = rg.getFontSizes(), wrap = rg.getWraps();
  var skip = {}, span = {};
  rg.getMergedRanges().forEach(function (m) {
    var r0 = m.getRow(), c0 = m.getColumn();
    if (r0 > nr || c0 > nc) return;
    span[r0 + ',' + c0] = [Math.min(m.getNumRows(), nr - r0 + 1), Math.min(m.getNumColumns(), nc - c0 + 1)];
    for (var i = 0; i < m.getNumRows(); i++) for (var j = 0; j < m.getNumColumns(); j++) if (i || j) skip[(r0 + i) + ',' + (c0 + j)] = 1;
  });
  var html = '<table style="border-collapse:collapse;font-family:Malgun Gothic,Arial,sans-serif;font-size:12px;table-layout:fixed"><colgroup>';
  for (var c = 1; c <= nc; c++) html += '<col style="width:' + sh.getColumnWidth(c) + 'px">';
  html += '</colgroup>';
  for (var r = 0; r < nr; r++) {
    html += '<tr style="height:' + sh.getRowHeight(r + 1) + 'px">';
    for (var k = 0; k < nc; k++) {
      var key = (r + 1) + ',' + (k + 1);
      if (skip[key]) continue;
      var sp = span[key], filled = v[r][k] !== '' || (bg[r][k] && bg[r][k] !== '#ffffff');
      var st = 'padding:2px 4px;vertical-align:middle;overflow:hidden;' + (wrap[r][k] ? 'white-space:pre-wrap;' : 'white-space:nowrap;') +
        'background:' + bg[r][k] + ';color:' + fc[r][k] + ';font-weight:' + fw[r][k] + ';font-size:' + fs[r][k] + 'pt;' +
        'text-align:' + (al[r][k] || 'general').replace('general-', '').replace('general', 'left') + ';' +
        (filled ? 'border:1px solid #bbb;' : '');
      html += '<td' + (sp ? ' rowspan="' + sp[0] + '" colspan="' + sp[1] + '"' : '') + ' style="' + st + '">' + esc_(v[r][k]) + '</td>';
    }
    html += '</tr>';
  }
  return html + '</table>';
}

/* ================================================================
 * 설정 점검 — 편집기에서 checkSetup 을 선택하고 ▶ 실행, 아래 "실행 로그"에서 결과 확인
 * ================================================================ */
function checkSetup() {
  var lines = [], ok = function (m) { lines.push('✅ ' + m); }, bad = function (m) { lines.push('❌ ' + m); };
  try {
    var values = masterValues_(), head = masterHeader_(values), contracts = parseContracts_(values);
    ok('자재정보: 머리글 ' + head.row + '행, 품목 ' + contracts.length + '줄');
    var keys = masterFields_(values).map(function (f) { return f.key; });
    ['item', 'priceMode', 'supplier', 'price', 'cur', 'start', 'end'].forEach(function (k) {
      if (keys.indexOf(k) < 0) bad('자재정보에 필수 열이 없습니다: ' + MASTER_COLUMNS[k][0]);
    });
  } catch (e) { bad('자재정보: ' + e.message); }

  var tid = prop_('TEMPLATE_ID');
  if (!tid) bad('스크립트 속성 TEMPLATE_ID 가 없습니다');
  else {
    try {
      var tpl = SpreadsheetApp.openById(tid);
      var missing = [TPL.SUMMARY, TPL.DETAIL, TPL.SCHEDULE, TPL.CONFIRM, TPL.PROTO].filter(function (n) { return !tpl.getSheetByName(n); });
      if (missing.length) bad('템플릿에 없는 탭: ' + missing.join(', ') + ' — templates/구매내역서_템플릿.xlsx 를 다시 올려 주세요');
      else ok('템플릿: ' + tpl.getName());
    } catch (e) { bad('템플릿을 열 수 없습니다 (ID·권한 확인): ' + e.message); }
  }
  var fid = prop_('OUTPUT_FOLDER_ID');
  if (fid) { try { ok('저장 폴더: ' + DriveApp.getFolderById(fid).getName()); } catch (e) { bad('OUTPUT_FOLDER_ID 폴더를 열 수 없습니다: ' + e.message); } }
  else lines.push('ℹ️ OUTPUT_FOLDER_ID 없음 — 구매내역서가 템플릿과 같은 위치에 저장됩니다');

  console.log(lines.join('\n'));
  return lines.join('\n');
}
