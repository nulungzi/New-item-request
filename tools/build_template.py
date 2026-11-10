"""
샘플 구매내역서(xlsx)에서 템플릿을 만든다.
- 구매요약 / 구매세부내역 / 구매일정 / 사업부 담당 확인 4개 탭만 남기고 데이터는 비운다.
- 서식 원본 행은 숨김 탭 `_proto` 에 모은다 (Apps Script가 여기서 서식을 복사한 뒤 지운다).

사용: python3 -I tools/build_template.py <샘플.xlsx> templates/구매내역서_템플릿.xlsx
"""
import sys
from copy import copy
import openpyxl

KEEP = ['구매요약', '구매세부내역', '구매일정', '사업부 담당 확인']
src_path, out_path = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(src_path)
for n in wb.sheetnames:
    if n not in KEEP:
        del wb[n]
wb._sheets = [wb[n] for n in KEEP]

S, D = wb['구매요약'], wb['구매세부내역']
P = wb.create_sheet('_proto')
P.sheet_state = 'hidden'


def copy_style(src, dst):
    dst._style = copy(src._style)


def proto_row(ws, src_row, dst_row, col_map):
    """col_map: {대상 열: 원본 열}"""
    for dc, sc in col_map.items():
        copy_style(ws.cell(src_row, sc), P.cell(dst_row, dc))
    h = ws.row_dimensions[src_row].height
    if h:
        P.row_dimensions[dst_row].height = h


same = lambda a, b: {c: c for c in range(a, b + 1)}
# 구매요약 (B..L)
proto_row(S, 4, 1, same(2, 12))   # 업체 행 (L열: 회색 도우미 글자)
proto_row(S, 5, 2, same(2, 12))   # CUR 별 합계
proto_row(S, 6, 3, same(2, 12))   # 총 합계
proto_row(S, 7, 4, same(2, 12))   # 환율 메모
# 구매세부내역 (B..T). 원본의 S(구분) 열은 빼고, 비고는 S, Need-by Date는 T.
detail_map = {**same(2, 18), 19: 20, 20: 18}
for i, r in enumerate([1, 2, 3, 4, 5, 6]):
    proto_row(D, r, 11 + i, same(2, 20))
proto_row(D, 7, 17, {**same(2, 18), 19: 18, 20: 20})   # 머리글: S=비고(테두리 있음), T=원본 비고 서식
proto_row(D, 8, 18, {**same(2, 18), 19: 18, 20: 20})   # 데이터 행

# 구매요약: 4행 이하 비우기
for rng in list(S.merged_cells.ranges):
    if rng.min_row >= 4:
        S.unmerge_cells(str(rng))
for row in S.iter_rows(min_row=4, max_row=S.max_row):
    for c in row:
        c.value = None
        c._style = copy(S.cell(400, 1)._style)
for r in range(4, S.max_row + 1):
    S.row_dimensions[r].height = None

# 구매세부내역: 전부 비우기 (열 너비만 유지)
for row in D.iter_rows(min_row=1, max_row=D.max_row):
    for c in row:
        c.value = None
        c._style = copy(D.cell(400, 1)._style)
for r in range(1, D.max_row + 1):
    D.row_dimensions[r].height = None
D.column_dimensions['S'].width = D.column_dimensions['T'].width   # 비고
D.column_dimensions['T'].width = 12                                 # Need-by Date
D.freeze_panes = None

# 구매일정: 12개월이 되도록 AA·AB열(13번째 달) 서식 추가
G = wb['구매일정']
for r in range(4, 11):
    copy_style(G.cell(r, 25), G.cell(r, 27))   # Y → AA
    copy_style(G.cell(r, 26), G.cell(r, 28))   # Z → AB
G.merge_cells('AA5:AB5')
# 일정 칸(E~AB) 너비: 원본은 E~V, W~AH 처럼 묶음 범위로 지정되어 있어 Y 이후가 기본 너비로 넓어진다.
# 묶음을 지우고 24칸을 같은 너비로 하나씩 지정한다.
from openpyxl.utils import get_column_letter
slot_w = G.column_dimensions['E'].width
for key in [k for k, d in G.column_dimensions.items() if d.min and d.max and d.max >= 5]:
    del G.column_dimensions[key]
for c in range(5, 29):
    G.column_dimensions[get_column_letter(c)].width = slot_w

# 사업부 담당 확인: 담당자·제목·날짜·날인 비우기
B = wb['사업부 담당 확인']
for a in ['A2', 'B2', 'C2']:
    B[a].value = None
B._images = []

# 원본에 남아 있던 이름 정의(다른 탭 참조, 개인 이름 등) 제거
for name in list(wb.defined_names.keys()):
    del wb.defined_names[name]
for ws in wb:
    ws.auto_filter.ref = None
    ws.defined_names.clear()
wb.properties.creator = None
wb.properties.lastModifiedBy = None
wb.properties.title = '구매내역서 템플릿'

wb.active = 0
wb.save(out_path)
print('saved', out_path)
