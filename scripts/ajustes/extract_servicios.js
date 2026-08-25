const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const excelPath = path.join(__dirname, '../MASTER Agricola Santa Elisa.xlsx');

console.log('Reading Excel file...');
const workbook = XLSX.readFile(excelPath);
const sheet = workbook.Sheets['Facturas'];

const servicios = [];

for (let row = 2; row < 2000; row++) {
    const catCell = sheet[XLSX.utils.encode_col(16) + row];

    if (catCell && catCell.v === 'SERVICIOS') {
        const filaCell = sheet[XLSX.utils.encode_col(0) + row];
        const fechaCell = sheet[XLSX.utils.encode_col(0) + row];
        const provCell = sheet[XLSX.utils.encode_col(3) + row];
        const nroCell = sheet[XLSX.utils.encode_col(6) + row];
        const detailCell = sheet[XLSX.utils.encode_col(7) + row];
        const totalCell = sheet[XLSX.utils.encode_col(15) + row];

        servicios.push({
            fila: row,
            fecha: fechaCell ? fechaCell.v : '',
            proveedor: provCell ? provCell.v : '',
            nro: nroCell ? nroCell.v : '',
            detalle: detailCell ? detailCell.v : '',
            total: totalCell ? parseFloat(totalCell.v) : 0
        });
    }
}

// Sort by amount descending
servicios.sort((a, b) => b.total - a.total);

console.log(`\nFound ${servicios.length} invoices categorized as SERVICIOS\n`);
console.log('TOP 30 by amount:');
console.log('='.repeat(120));

servicios.slice(0, 30).forEach((s, idx) => {
    console.log(`${(idx + 1).toString().padStart(2)}. Fila ${s.fila} | ${s.proveedor.substring(0, 30).padEnd(30)} | ${s.nro.toString().substring(0, 15).padEnd(15)} | $${s.total.toLocaleString('es-CL').padStart(15)} CLP`);
    if (s.detalle) {
        console.log(`    Detalle: ${s.detalle.substring(0, 100)}`);
    }
    console.log('');
});

// Also save to JSON for reference
const jsonPath = path.join(__dirname, 'servicios_review.json');
fs.writeFileSync(jsonPath, JSON.stringify(servicios, null, 2));
console.log(`Full list saved to: servicios_review.json`);
