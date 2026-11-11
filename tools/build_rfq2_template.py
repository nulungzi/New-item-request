"""
견적요청서 템플릿 2 (품목 목록형, Need-by 마다 한 줄)를 샘플에서 만든다.
- 1행 머리글은 그대로, 2행은 품목 행 서식 원본 (값만 지움), 3행부터는 비움
- 열: A L/I · B Part Number · C Part Name(Description) · D PO Qty · E UOM · F U/P · G T/P · H Lead Time · I MOQ
      J MFG(QPL) · K Need by date · L Quallity Doc · M NOTE · N Program · O Supplier
- Description·QPL 은 줄바꿈, 열 너비를 넓혀 잘리지 않게

사용: python3 -I tools/build_rfq2_template.py <RFQ샘플.xlsx> templates/견적요청서_템플릿2_목록형.xlsx
"""
import sys
from copy import copy
import openpyxl

src, out = sys.argv[1:3]
wb = openpyxl.load_workbook(src)
ws = wb.worksheets[0]
for extra in wb.worksheets[1:]:
    del wb[extra.title]
ws.title = 'RFQ'

blank = ws.cell(ws.max_row + 50, 1)._style
for r in range(2, ws.max_row + 1):
    for c in range(1, 16):
        cell = ws.cell(r, c)
        cell.value = None
        if r > 2:
            cell._style = copy(blank)
    ws.row_dimensions[r].height = None

# 2행(품목 행 원본): 가운데 맞춤·줄바꿈·숫자 서식
def align(c, h=None, wrap=False):
    a = copy(ws.cell(2, c).alignment)
    a.vertical = 'center'
    if h:
        a.horizontal = h
    a.wrap_text = wrap
    ws.cell(2, c).alignment = a

for c in range(1, 16):
    align(c)
for c in (3, 10, 13):            # Description · QPL · NOTE 줄바꿈
    align(c, wrap=True)
for c in (8, 9, 11, 12):         # Lead Time · MOQ · Need by · Quality Doc 가운데
    align(c, 'center')
ws.cell(2, 6).number_format = '#,##0.00'      # U/P (통화는 업체가 적음)
ws.cell(2, 7).number_format = '#,##0.00'      # T/P = 수량 × U/P

widths = {'A': 5.5, 'B': 22, 'C': 46, 'D': 10, 'E': 7, 'F': 11, 'G': 13, 'H': 11, 'I': 9,
          'J': 30, 'K': 13, 'L': 13, 'M': 18, 'N': 14, 'O': 16}
for k, w in widths.items():
    ws.column_dimensions[k].width = w
ws.freeze_panes = 'A2'

wb.properties.creator = None
wb.properties.lastModifiedBy = None
wb.properties.title = '견적요청서 템플릿 2 (목록형)'
wb.active = 0
wb.save(out)
print('saved', out)
