const XLSX = require('xlsx');
const path = require('path');

const excelPath = path.join(__dirname, '../MASTER Agricola Santa Elisa.xlsx');

console.log('Checking banco sheet for matched movements...');
const workbook = XLSX.readFile(excelPath);
const sheet = workbook.Sheets['Cuenta Banco'];

let auto_matched = 0;
let ambiguo = 0;
let unlinked = 0;

for (let row = 2; row < 5000; row++) {
    const tipoCell = sheet[XLSX.utils.encode_col(6) + row];  // Col G = Tipo

    if (tipoCell && tipoCell.v) {
        const tipo = tipoCell.v;
        if (tipo === 'factura') {
            auto_matched++;
        } else if (tipo === 'ambiguo') {
            ambiguo++;
        }
    } else if (!tipoCell || tipoCell.v === '') {
        unlinked++;
    }
}

console.log(`Auto-matched (Tipo=factura): ${auto_matched}`);
console.log(`Ambiguos (Tipo=ambiguo): ${ambiguo}`);
console.log(`Unlinked: ${unlinked}`);

workbook.close = () => {};
