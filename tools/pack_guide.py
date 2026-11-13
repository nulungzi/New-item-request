"""
capture_guide.js 로 찍은 사진(PNG)을 WebP 로 줄여 apps-script/Guide.html (JSON) 로 묶는다.
사용: python3 -I tools/pack_guide.py <작업 폴더>/guide apps-script/Guide.html
"""
import sys, json, base64, io
from PIL import Image

d, out = sys.argv[1], sys.argv[2]
meta = json.load(open(d + '/shots.json'))
lines = ['{"hot":' + json.dumps(meta['hots'], separators=(',', ':')) + ',', '"img":{']
items = list(meta['shots'].items())
for i, (k, v) in enumerate(items):
    im = Image.open(v['file']).convert('RGB')
    if im.width > 1300:
        im = im.resize((1300, round(im.height * 1300 / im.width)), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, 'WEBP', quality=60, method=6)
    src = 'data:image/webp;base64,' + base64.b64encode(buf.getvalue()).decode()
    lines.append(json.dumps(k) + ':' + json.dumps({'src': src, 'w': v['w'], 'h': v['h']}, separators=(',', ':')) + (',' if i < len(items) - 1 else ''))
lines.append('}}')
open(out, 'w', encoding='utf-8').write('\n'.join(lines) + '\n')
json.load(open(out, encoding='utf-8'))
print('saved', out)
