// Apps Script에 붙여넣을 배포용 파일을 만든다: dist/Code.gs (Logic.js + Code.js), dist/Index.html, dist/appsscript.json
// 사용: node tools/bundle.js   (npm run build)
const fs = require('fs'), path = require('path');
const root = path.join(__dirname, '..'), src = path.join(root, 'apps-script'), out = path.join(root, 'dist');

function bundle() {
  const strip = s => s.replace(/\n*if \(typeof module !== 'undefined'\) \{[\s\S]*?\n\}\s*$/, '\n');   // Node 테스트용 export 제거
  const logic = strip(fs.readFileSync(path.join(src, 'Logic.js'), 'utf8'));
  const code = fs.readFileSync(path.join(src, 'Code.js'), 'utf8');
  const head = [
    '// ===== 소요제기 판정 웹앱 · Apps Script 배포용 (자동 생성: npm run build) =====',
    '// 이 파일 전체를 Apps Script 의 Code.gs 에 붙여넣으세요. Logic.gs 파일이 따로 있으면 삭제하세요.',
    '// 원본: apps-script/Logic.js + apps-script/Code.js',
    ''
  ].join('\n');
  return {
    'Code.gs': head + '\n// ---------- Logic.js ----------\n' + logic.trimEnd() + '\n\n// ---------- Code.js ----------\n' + code.trimEnd() + '\n',
    'Index.html': fs.readFileSync(path.join(src, 'Index.html'), 'utf8'),
    'appsscript.json': fs.readFileSync(path.join(src, 'appsscript.json'), 'utf8')
  };
}

if (require.main === module) {
  fs.mkdirSync(out, { recursive: true });
  const files = bundle();
  for (const [name, body] of Object.entries(files)) {
    fs.writeFileSync(path.join(out, name), body);
    console.log('dist/' + name, body.split('\n').length + '줄');
  }
}
module.exports = { bundle };
