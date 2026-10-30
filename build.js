// Builds dist/ReportBuilder.html: a single offline file with both libraries inlined.
const fs = require('fs');
const path = require('path');
const LIBS = '/tmp/claude-0/-home-user-Claude/3457d303-2a1d-5cfb-99a3-368bc144b851/scratchpad/libs/node_modules';
const read = (p) => fs.readFileSync(p, 'utf8');
const guard = (name, code) => {
  if (/<\/script/i.test(code)) throw new Error(name + ' contains </script>');
  return code;
};
let html = read(path.join(__dirname, 'src/index.html'));
html = html.replace('/*__CSS__*/', () => read(path.join(__dirname, 'src/styles.css')));
html = html.replace('/*__XLSX__*/', () => guard('xlsx', read(path.join(LIBS, 'xlsx/dist/xlsx.full.min.js'))));
html = html.replace('/*__EXCELJS__*/', () => guard('exceljs', read(path.join(LIBS, 'exceljs/dist/exceljs.min.js'))));
html = html.replace('/*__APP__*/', () => guard('app', read(path.join(__dirname, 'src/app.js'))));
fs.mkdirSync(path.join(__dirname, 'dist'), { recursive: true });
fs.writeFileSync(path.join(__dirname, 'dist/ReportBuilder.html'), html);
console.log('dist/ReportBuilder.html', (html.length / 1048576).toFixed(2) + ' MB');

let mask = read(path.join(__dirname, 'src/mask.html'));
mask = mask.replace('/*__CSS__*/', () => read(path.join(__dirname, 'src/styles.css')));
mask = mask.replace('/*__XLSX__*/', () => guard('xlsx', read(path.join(LIBS, 'xlsx/dist/xlsx.full.min.js'))));
mask = mask.replace('/*__EXCELJS__*/', () => guard('exceljs', read(path.join(LIBS, 'exceljs/dist/exceljs.min.js'))));
mask = mask.replace('/*__APP__*/', () => guard('maskapp', read(path.join(__dirname, 'src/mask.js'))));
fs.writeFileSync(path.join(__dirname, 'dist/MaskKit.html'), mask);
console.log('dist/MaskKit.html', (mask.length / 1048576).toFixed(2) + ' MB');
