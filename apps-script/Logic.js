/**
 * 순수 로직 (Google 서비스 호출 없음) — Node 테스트로도 검증 가능.
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
    var valid = all.filter(function (c) {
      return (!c.start || c.start <= d) && (!c.end || d <= c.end);
    }).sort(function (a, b) { return (a.price || 0) - (b.price || 0); });
    var res = { line: q.line, item: q.item, qty: q.qty, needBy: q.needBy, program: q.program, candidates: valid, warnings: [] };
    if (valid.length) {
      res.status = '계약';
      res.reason = valid.length > 1 ? '유효 계약 ' + valid.length + '건 — 업체 선택 필요' : '';
      res.warnings = needByWarnings_(q, valid[0], today);
    } else {
      res.status = 'Spot';
      if (!all.length) {
        res.reason = '계약 없음';
      } else {
        var last = all.reduce(function (a, b) { return (b.end || 0) > (a.end || 0) ? b : a; });
        res.reason = last.end && last.end < d
          ? 'Need-by 시점 계약 만료 (' + last.supplier + ', ' + fmtDate_(last.end) + ' 종료)'
          : 'Need-by 시점 계약 시작 전 (' + last.supplier + ', ' + fmtDate_(last.start) + ' 시작)';
        res.desc = last.desc;
      }
    }
    return res;
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
 */
function buildPurchasePlan_(lines, curs, rates, base) {
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

  /* ---------- 구매일정: 작성일 달부터 11개월 머리글 ---------- */
  var G = TPL.SCHEDULE;
  add({ op: 'unmerge', s: G, r1: 4, c1: 5, r2: 4, c2: 26 });
  for (var c = 5; c <= 26; c++) set(G, 4, c, '');
  var y = base.getFullYear(), m = base.getMonth() + 1, segStart = 5;
  for (var k = 0; k < 11; k++) {
    var col = 5 + k * 2;
    set(G, 5, col, m);
    var endOfYear = m === 12 || k === 10;
    if (endOfYear) {
      add({ op: 'cellFmt', s: G, r: 4, c: segStart, fs: G, fr: 4, fc: 5 });
      set(G, 4, segStart, y + '년');
      add({ op: 'merge', s: G, r1: 4, c1: segStart, r2: 4, c2: col + 1 });
      segStart = col + 2;
    }
    m++;
    if (m > 12) { m = 1; y++; }
  }

  /* ---------- 사업부 담당 확인 ---------- */
  var modes = [], sups = [];
  lines.forEach(function (l) {
    if (l.priceMode && modes.indexOf(l.priceMode) < 0) modes.push(l.priceMode);
    if (sups.indexOf(l.supplier) < 0) sups.push(l.supplier);
  });
  set(TPL.CONFIRM, 2, 2, sups.join(', ') + ' ' + modes.join('/') + ' 자재 구매 품의');
  set(TPL.CONFIRM, 2, 3, { date: fmtDate_(base) });
  add({ op: 'bg', s: TPL.CONFIRM, r: 2, c: 1, color: TPL.FILL });   // 담당자 이름 입력 칸

  return ops;
}

if (typeof module !== 'undefined') {
  module.exports = { buildPurchasePlan_: buildPurchasePlan_, curFmt_: curFmt_, leadDays_: leadDays_, parseContracts_: parseContracts_, parseRequest_: parseRequest_, classify_: classify_, summarize_: summarize_, toDate_: toDate_, fmtDate_: fmtDate_ };
}
