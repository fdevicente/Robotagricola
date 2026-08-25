const XLSX = require('xlsx');
const path = require('path');

const excelPath = path.join(__dirname, '../MASTER Agricola Santa Elisa.xlsx');

console.log('Reading Excel file...');
const workbook = XLSX.readFile(excelPath);
const sheet = workbook.Sheets['Facturas'];

const categories = {};
let revisar_count = 0;
let total_invoices = 0;

for (let row = 2; row < 2000; row++) {
    const catCell = sheet[XLSX.utils.encode_col(16) + row];

    if (catCell && catCell.v) {
        total_invoices++;
        const cat = catCell.v;

        if (cat === 'REVISAR') {
            revisar_count++;
        } else {
            if (!categories[cat]) {
                categories[cat] = { count: 0, total: 0 };
            }
            categories[cat].count++;

            // Get amount
            const totalCell = sheet[XLSX.utils.encode_col(15) + row];  // Column P = index 15
            const amount = totalCell && totalCell.v ? parseFloat(totalCell.v) : 0;
            categories[cat].total += amount;
        }
    }
}

console.log(`\nTotal invoices scanned: ${total_invoices}`);
console.log(`Remaining REVISAR: ${revisar_count}\n`);

console.log('Category Distribution:');
const sorted = Object.entries(categories)
    .sort((a, b) => b[1].total - a[1].total);

let grand_total = 0;
sorted.forEach(([cat, data]) => {
    console.log(`${cat}: ${data.count} items, $${data.total.toLocaleString('es-CL')} CLP`);
    grand_total += data.total;
});

console.log(`\nGrand Total: ${grand_total.toLocaleString('es-CL')} CLP`);
console.log(`\nSUMMARY: ${total_invoices - revisar_count} categorized, ${revisar_count} remaining`);
