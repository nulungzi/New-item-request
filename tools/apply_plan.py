"""
Code.gs 의 buildPurchasePlan_ 결과(JSON)를 엑셀 템플릿에 적용해 미리보기 파일을 만든다.
Apps Script의 applyPlan_ 과 같은 동작을 openpyxl로 흉내 낸다 (테스트·미리보기용).

사용: python3 -I tools/apply_plan.py <템플릿.xlsx> <plan.json> <출력.xlsx>
"""
import sys, json, datetime
from copy import copy
import openpyxl
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

tpl, plan_path, out = sys.argv[1:4]
wb = openpyxl.load_workbook(tpl)
ops = json.load(open(plan_path, encoding='utf-8'))
P = wb['_proto'] if '_proto' in wb.sheetnames else None


def rng(o):
    return '%s%d:%s%d' % (get_column_letter(o['c1']), o['r1'], get_column_letter(o['c2']), o['r2'])


for o in ops:
    op = o['op']
    if op == 'fmt':
        ws = wb[o['s']]
        for c in range(o['c1'], o['c2'] + 1):
            ws.cell(o['r'], c)._style = copy(P.cell(o['from'], c)._style)
    elif op == 'rowFmt':
        ws, src = wb[o['s']], wb[o['fs']]
        for c in range(o['c1'], o['c2'] + 1):
            ws.cell(o['r'], c)._style = copy(src.cell(o['fr'], c)._style)
    elif op == 'cellFmt':
        wb[o['s']].cell(o['r'], o['c'])._style = copy(wb[o['fs']].cell(o['fr'], o['fc'])._style)
    elif op == 'set':
        v = o['v']
        if isinstance(v, dict) and 'date' in v:
            v = datetime.datetime.strptime(v['date'], '%Y-%m-%d')
        if v == '':
            v = None
        wb[o['s']].cell(o['r'], o['c']).value = v
    elif op == 'nf':
        wb[o['s']].cell(o['r'], o['c']).number_format = o['f']
    elif op == 'bg':
        wb[o['s']].cell(o['r'], o['c']).fill = PatternFill('solid', fgColor=o['color'].lstrip('#').upper())
    elif op == 'merge':
        wb[o['s']].merge_cells(rng(o))
    elif op == 'unmerge':
        ws = wb[o['s']]
        for m in list(ws.merged_cells.ranges):
            if m.min_row >= o['r1'] and m.max_row <= o['r2'] and m.min_col >= o['c1'] and m.max_col <= o['c2']:
                ws.unmerge_cells(str(m))
    elif op == 'height':
        wb[o['s']].row_dimensions[o['r']].height = o['pt']
    elif op == 'bottom':
        ws = wb[o['s']]
        for c in range(o['c1'], o['c2'] + 1):
            cell = ws.cell(o['r'], c)
            b = copy(cell.border)
            cell.border = Border(left=b.left, right=b.right, top=b.top, bottom=Side(style='thin', color='FF000000'))
    elif op == 'list':
        ws = wb[o['s']]
        dv = DataValidation(type='list', formula1='"' + ','.join(o['items']) + '"', showErrorMessage=False)
        ws.add_data_validation(dv)
        col = get_column_letter(o['c'])
        dv.add('%s%d:%s%d' % (col, o['r1'], col, o['r2']))
    elif op == 'freeze':
        wb[o['s']].freeze_panes = wb[o['s']].cell(o['rows'] + 1, o['cols'] + 1)
    elif op == 'width':
        wb[o['s']].column_dimensions[get_column_letter(o['c'])].width = round(o['px'] / 7, 2)
    elif op == 'checkbox':
        pass   # 엑셀 미리보기에서는 TRUE/FALSE 값 그대로
    elif op == 'deleteCols':
        ws = wb[o['s']]
        for m in list(ws.merged_cells.ranges):
            if m.max_col >= o['c']:
                ws.unmerge_cells(str(m))
        ws.delete_cols(o['c'], o['n'])
    else:
        raise SystemExit('unknown op ' + op)

if P is not None:
    del wb['_proto']
wb.active = 0
wb.save(out)
print('saved', out)
