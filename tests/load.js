// apps-script/Logic.gs + Code.gs 를 테스트에서 불러온다.
// - load(): 이 Node 환경에 그대로 올려 순수 함수(Logic.gs)를 바로 호출
// - inContext(ctx, order): 가짜 SpreadsheetApp 등을 넣은 별도 환경에 두 파일을 올려 Code.gs 를 시험
const fs = require('fs'), vm = require('vm'), path = require('path');
const DIR = path.join(__dirname, '..', 'apps-script');
const FILES = ['Logic.gs', 'Code.gs'];
const source = f => fs.readFileSync(path.join(DIR, f), 'utf8');
let loaded = null;
function load() {
  if (!loaded) { FILES.forEach(f => vm.runInThisContext(source(f), { filename: f })); loaded = globalThis; }
  return loaded;
}
function inContext(ctx, order) {
  vm.createContext(ctx);
  (order || FILES).forEach(f => vm.runInContext(source(f), ctx, { filename: f }));
  return ctx;
}
module.exports = { load, inContext, source, FILES };
