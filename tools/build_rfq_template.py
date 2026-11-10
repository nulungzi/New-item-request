"""
견적요청서(RFQ) 샘플에서 12개월 템플릿을 만든다.
원본: B~H 품목 정보, I~Q 9개월, R RFQ Qty(ttl), S~AB 업체 기입란
결과: B~H 그대로, I~T 12개월, U RFQ Qty(ttl), V~AE 업체 기입란. 3행은 서식 원본(값 없음)

사용: python3 -I tools/build_rfq_template.py <RFQ샘플.xlsx> templates/견적요청서_템플릿.xlsx
"""
import sys
from copy import copy
import openpyxl
from openpyxl.utils import get_column_letter as L

src_path, out_path = sys.argv[1:3]
src = openpyxl.load_workbook(src_path)['RFQ_List']


def cp(s, d):
    """다른 통합 문서로 셀 서식 복사"""
    d.font, d.fill, d.border = copy(s.font), copy(s.fill), copy(s.border)
    d.alignment, d.number_format, d.protection = copy(s.alignment), s.number_format, copy(s.protection)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'RFQ_List'

# 새 열 → 원본 열
cmap = {c: c for c in range(1, 9)}
cmap.update({c: 9 for c in range(9, 21)})          # 12개월 (원본 I 서식)
cmap[21] = 18                                       # RFQ Qty(ttl)
cmap.update({c: c - 3 for c in range(22, 32)})      # 업체 기입란 S~AB → V~AE

for r in (1, 2, 3):
    for dc, sc in cmap.items():
        s, d = src.cell(r, sc), ws.cell(r, dc)
        cp(s, d)
        if r == 2 and not (9 <= dc <= 20):
            d.value = s.value
    ws.row_dimensions[r].height = src.row_dimensions[r].height
# 1행 머리글 "Required Shipping date" 를 12개월 칸 위에 병합
ws.cell(1, 9).value = src.cell(1, 9).value
for c in range(10, 21):
    cp(src.cell(1, 17), ws.cell(1, c))
ws.merge_cells('I1:T1')
# 12개월 머리글 칸은 서식만 (날짜는 만들 때 채움)
for c in range(9, 21):
    cp(src.cell(2, 9), ws.cell(2, c))
# 열 너비
src_w = {}
for k, d in src.column_dimensions.items():
    for c in range(d.min or 0, (d.max or 0) + 1):
        src_w[c] = d.width
for dc, sc in cmap.items():
    if sc in src_w and src_w[sc]:
        ws.column_dimensions[L(dc)].width = src_w[sc]
ws.freeze_panes = 'D3'
ws.sheet_view.showGridLines = False
wb.properties.creator = None
wb.properties.title = '견적요청서 템플릿'
wb.save(out_path)
print('saved', out_path)
