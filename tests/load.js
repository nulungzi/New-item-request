// apps-script/Code.gs 를 테스트에서 불러온다.
// - load(): 이 Node 환경에 그대로 올려 1부(순수 함수)를 바로 호출 (배열·객체 비교가 자연스럽게 됨)
// - inContext(ctx): 가짜 SpreadsheetApp 등을 넣은 별도 환경에 올려 2부(시트 입출력)를 시험
const fs = require('fs'), vm = require('vm'), path = require('path');
const FILE = path.join(__dirname, '..', 'apps-script', 'Code.gs');
const source = () => fs.readFileSync(FILE, 'utf8');
let loaded = null;
function load() {
  if (!loaded) { vm.runInThisContext(source(), { filename: 'Code.gs' }); loaded = globalThis; }
  return loaded;
}
function inContext(ctx) {
  vm.createContext(ctx);
  vm.runInContext(source(), ctx, { filename: 'Code.gs' });
  return ctx;
}
module.exports = { load, inContext, source, FILE };
