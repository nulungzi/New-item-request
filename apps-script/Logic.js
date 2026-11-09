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
  product: ['Product Name'],
  shipFrom: ['Ship From'],
  leadtime: ['Leadtime'],
  shelfLife: ['Shelf Life'],
  taxCode: ['Tax Code']
};

function headerKey_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, '').toLowerCase();
}

function itemKey_(s) {
  return String(s == null ? '' : s).trim().toUpperCase();
}

function toDate_(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
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
      product: String(get('product')).trim(),
      shipFrom: String(get('shipFrom')).trim(),
      leadtime: String(get('leadtime')).trim(),
      shelfLife: String(get('shelfLife')).trim(),
      taxCode: get('taxCode') === '' ? '' : get('taxCode')
    });
  }
  return out;
}

/**
 * 소요제기 붙여넣기 파싱. 한 줄에 "품번  수량  Program(선택)".
 * 엑셀에서 복사한 탭 구분, 쉼표, 공백 구분 모두 허용. 수량이 숫자가 아닌 줄(머리글 등)은 건너뛴다.
 */
function parseRequest_(text) {
  var out = [], skipped = [];
  String(text || '').split(/\r?\n/).forEach(function (line, n) {
    if (!line.trim()) return;
    var cells = line.indexOf('\t') >= 0 ? line.split('\t') : line.split(/\s*,\s*|\s+/);
    cells = cells.map(function (c) { return c.trim(); });
    var qty = toNumber_(cells[1]);
    if (!cells[0] || qty === null) { skipped.push(n + 1); return; }
    out.push({ line: n + 1, item: itemKey_(cells[0]), qty: qty, program: cells.slice(2).join(' ').trim() });
  });
  return { requests: out, skipped: skipped };
}

/** 각 소요 품목을 기준일 기준으로 계약/Spot 판정. 유효 계약이 여럿이면 단가 낮은 순으로 후보를 준다. */
function classify_(requests, contracts, baseDate) {
  var byItem = {};
  contracts.forEach(function (c) { (byItem[c.item] = byItem[c.item] || []).push(c); });
  return requests.map(function (q) {
    var all = byItem[q.item] || [];
    var valid = all.filter(function (c) {
      return (!c.start || c.start <= baseDate) && (!c.end || baseDate <= c.end);
    }).sort(function (a, b) { return (a.price || 0) - (b.price || 0); });
    var res = { line: q.line, item: q.item, qty: q.qty, program: q.program, candidates: valid };
    if (valid.length) {
      res.status = '계약';
      res.reason = valid.length > 1 ? '유효 계약 ' + valid.length + '건 — 업체 선택 필요' : '';
    } else {
      res.status = 'Spot';
      if (!all.length) {
        res.reason = '계약 없음';
      } else {
        var last = all.reduce(function (a, b) { return (b.end || 0) > (a.end || 0) ? b : a; });
        res.reason = last.end && last.end < baseDate
          ? '계약 만료 (' + last.supplier + ', ' + fmtDate_(last.end) + ' 종료)'
          : '계약 시작 전 (' + last.supplier + ', ' + fmtDate_(last.start) + ' 시작)';
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

if (typeof module !== 'undefined') {
  module.exports = { parseContracts_: parseContracts_, parseRequest_: parseRequest_, classify_: classify_, summarize_: summarize_, toDate_: toDate_, fmtDate_: fmtDate_ };
}
