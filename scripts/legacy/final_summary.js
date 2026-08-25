const XLSX = require('xlsx');
const path = require('path');

const excelPath = path.join(__dirname, '../MASTER Agricola Santa Elisa.xlsx');

console.log('Reading Excel file...');
const workbook = XLSX.readFile(excelPath);
const sheet = workbook.Sheets['Facturas'];

const categories = {};
let revisar_count = 0;

for (let row = 2; row < 2000; row++) {
    const catCell = sheet[XLSX.utils.encode_col(16) + row];

    if (catCell && catCell.v) {
        const cat = catCell.v;

        if (cat === 'REVISAR') {
            revisar_count++;
        } else {
            if (!categories[cat]) {
                categories[cat] = { count: 0, total: 0 };
            }
            categories[cat].count++;

            const totalCell = sheet[XLSX.utils.encode_col(15) + row];
            const amount = totalCell && totalCell.v ? parseFloat(totalCell.v) : 0;
            categories[cat].total += amount;
        }
    }
}

console.log('═'.repeat(80));
console.log('FINAL CATEGORIZATION SUMMARY');
console.log('═'.repeat(80));
console.log(`\nRemaining REVISAR: ${revisar_count}`);
console.log(`Total categorized invoices: ${Object.values(categories).reduce((sum, c) => sum + c.count, 0)}\n`);

const sorted = Object.entries(categories)
    .sort((a, b) => b[1].total - a[1].total);

let grand_total = 0;
sorted.forEach(([cat, data]) => {
    console.log(`${cat.padEnd(40)} ${data.count.toString().padStart(4)} items    $${data.total.toLocaleString('es-CL').padStart(18)} CLP`);
    grand_total += data.total;
});

console.log('─'.repeat(80));
console.log(`${'GRAND TOTAL'.padEnd(40)} ${Object.values(categories).reduce((sum, c) => sum + c.count, 0).toString().padStart(4)} items    $${grand_total.toLocaleString('es-CL').padStart(18)} CLP`);
console.log('═'.repeat(80));
