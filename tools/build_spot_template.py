"""
Spot 견적 구매내역서 샘플에서 템플릿을 만든다.
- 전체구매요약 / 구매내역서 / 견적종합 / 구매일정 / 사업부 담당 확인 5개 탭만 남기고 데이터를 비운다.
- 행 서식 원본은 숨김 탭 `_proto` 에 모은다 (Apps Script가 복사한 뒤 지운다).
    1  전체구매요약 품목 행      2  CUR 별 합계 첫 행   3  CUR 별 합계 다음 행
    4  총 합계(KRW) 행           5  적용환율 메모 행
    11 구매내역서 품목 행 (A~CX)
    21 견적종합 업체 행 (A~AS)
- 구매내역서 통화 합계를 USD·EUR·JPY·SGD·KRW 로, 환율 칸을 USD·JPY·EUR·SGD 로 바꾼다.

사용: python3 -I tools/build_spot_template.py <Spot구매내역서샘플.xlsx> templates/Spot_구매내역서_템플릿.xlsx
"""
import sys
from copy import copy
import openpyxl
from openpyxl.utils import get_column_letter as L

KEEP = ['전체구매요약', '구매내역서', '견적종합', '구매일정', '사업부 담당 확인']
src, out = sys.argv[1:3]
wb = openpyxl.load_workbook(src)
for n in wb.sheetnames:
    if n not in KEEP:
        del wb[n]
wb._sheets = [wb[n] for n in KEEP]
for ws in wb:
    ws.sheet_state = 'visible'
S, D, Q, G, C = (wb[n] for n in KEEP)
P = wb.create_sheet('_proto')
P.sheet_state = 'hidden'


def cp(s, d):
    d._style = copy(s._style)


def proto(ws, src_row, dst_row, c1, c2):
    for c in range(c1, c2 + 1):
        cp(ws.cell(src_row, c), P.cell(dst_row, c))
    if ws.row_dimensions[src_row].height:
        P.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def clear(ws, r1, c1=1, c2=None):
    """r1 행부터 아래를 값·서식·병합 모두 비운다"""
    c2 = c2 or ws.max_column
    for m in list(ws.merged_cells.ranges):
        if m.max_row >= r1:
            ws.unmerge_cells(str(m))
    blank = ws.cell(ws.max_row + 50, 1)._style
    for row in ws.iter_rows(min_row=r1, max_row=ws.max_row, min_col=c1, max_col=c2):
        for cell in row:
            cell.value = None
            cell._style = copy(blank)
    for r in range(r1, ws.max_row + 1):
        ws.row_dimensions[r].height = None


# ---------- 서식 원본 ----------
proto(S, 5, 1, 1, 16)
proto(S, 13, 2, 1, 16)
proto(S, 14, 3, 1, 16)
proto(S, 15, 4, 1, 16)
proto(S, 16, 5, 1, 16)
proto(D, 11, 11, 1, 102)
proto(Q, 5, 21, 1, 45)


def wrap(cell):
    a = copy(cell.alignment)
    a.wrap_text, a.vertical = True, 'center'
    cell.alignment = a


# Supplier Remarks 는 줄바꿈 (추가 비용 등을 여러 줄로 적음)
wrap(P.cell(21, 40))                     # 견적종합 AN
for c in [22] + list(range(26, 103, 4)):  # 구매내역서 V(선정 Remark) · 업체별 Remark
    wrap(P.cell(11, c))

# ---------- 전체구매요약: 4행(머리글)까지만 ----------
clear(S, 5)

# ---------- 구매내역서 ----------
for m in list(D.merged_cells.ranges):
    if m.min_row <= 7 and m.min_col >= 11 and m.max_col <= 13:
        D.unmerge_cells(str(m))
cur_style = D['L2']._style
val_style = {'USD': D['M2']._style, 'KRW': D['M3']._style, 'JPY': D['M4']._style}
lbl_style, tot_lbl, tot_val = D['K2']._style, D['K5']._style, D['M5']._style
curs = ['USD', 'EUR', 'JPY', 'SGD', 'KRW']
for r in range(2, 8):
    for c in (11, 12, 13):
        D.cell(r, c).value = None
fmt = {'USD': '[$$]#,##0.00', 'EUR': '[$€]#,##0.00', 'JPY': '[$¥]#,##0', 'SGD': '"S$"#,##0.00', 'KRW': '[$₩-412]#,##0'}
for i, cur in enumerate(curs):
    r = 2 + i
    D.cell(r, 11)._style = copy(lbl_style)
    D.cell(r, 12)._style = copy(cur_style)
    D.cell(r, 12).value = cur
    D.cell(r, 13)._style = copy(val_style.get(cur, val_style['USD']))
    D.cell(r, 13).number_format = fmt[cur]
    D.cell(r, 13).value = '=SUMIF($K$11:$K$1000,"%s",$M$11:$M$1000)' % cur
    D.row_dimensions[r].height = D.row_dimensions[2].height
D.cell(2, 11).value = ' 총 합계'
D.merge_cells('K2:K6')
for c in (11, 12, 13):
    D.cell(7, c)._style = copy([tot_lbl, tot_lbl, tot_val][c - 11])
D['K7'].value, D['L7'].value = '총 합계', 'KRW'
D['M7'].value = '=M6+M2*$P$3+M3*$R$3+M4*$Q$3+M5*$S$3'
D.row_dimensions[7].height = D.row_dimensions[2].height
# 환율: P USD · Q JPY · R EUR · S SGD
for c, cur in zip(range(16, 20), ['USD', 'JPY', 'EUR', 'SGD']):
    cp(D['R2'], D.cell(2, c)); cp(D['R3'], D.cell(3, c))
    D.cell(2, c).value = cur
    D.cell(3, c).value = None
D['O2'].value = None
clear(D, 11, 1, 102)
for c in range(23, 103, 4):   # 업체 이름은 만들 때 채움
    D.cell(8, c).value = None

# ---------- 견적종합 ----------
Q['G1'].value = None
for c, cur, ref in zip(range(42, 46), ['USD', 'JPY', 'EUR', 'SGD'], ['P3', 'Q3', 'R3', 'S3']):
    cp(Q['AR1'], Q.cell(1, c)); cp(Q['AR2'], Q.cell(2, c))
    Q.cell(1, c).value = cur
    Q.cell(2, c).value = "='구매내역서'!" + ref
Q.column_dimensions['AS'].width = Q.column_dimensions['AR'].width or 9
for c in range(12, 24):
    Q.cell(3, c).value = None
clear(Q, 5, 1, 45)

# ---------- 구매일정: 12개월 (계약 구매내역서 템플릿과 같게) ----------
# 샘플은 10개월(E~X)까지만 서식이 있어 Y~AB(11·12번째 달)에 W·X 서식을 복사
for r in range(4, 11):
    for c in (25, 27):
        cp(G.cell(r, 23), G.cell(r, c))
        cp(G.cell(r, 24), G.cell(r, c + 1))
if 'Y5:Z5' not in [str(m) for m in G.merged_cells.ranges]:
    G.merge_cells('Y5:Z5')
if 'AA5:AB5' not in [str(m) for m in G.merged_cells.ranges]:
    G.merge_cells('AA5:AB5')
slot_w = G.column_dimensions['E'].width
for key in [k for k, d in G.column_dimensions.items() if d.min and d.max and d.max >= 5]:
    del G.column_dimensions[key]
for c in range(5, 29):
    G.column_dimensions[L(c)].width = slot_w

# ---------- 사업부 담당 확인 ----------
for a in ['A2', 'B2', 'C2']:
    C[a].value = None
C['B1'].value = '구매내역서 제목'
C._images = []

for name in list(wb.defined_names.keys()):
    del wb.defined_names[name]
for ws in wb:
    ws.auto_filter.ref = None
    ws.defined_names.clear()
    ws.print_area = None
wb.properties.creator = None
wb.properties.lastModifiedBy = None
wb.properties.title = 'Spot 구매내역서 템플릿'
wb.active = 0
wb.save(out)
print('saved', out)
