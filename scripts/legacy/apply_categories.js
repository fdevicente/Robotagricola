const XLSX = require('xlsx');
const path = require('path');

const excelPath = path.join(__dirname, '../MASTER Agricola Santa Elisa.xlsx');

function getCategory(proveedor, detalle) {
    if (!proveedor) proveedor = '';
    if (!detalle) detalle = '';

    const p = proveedor.toUpperCase();
    const d = detalle.toUpperCase();

    if (p.includes('ECOSMART')) return 'SERVICIOS DE EXPORTACION';
    if (p.includes('S-INVEST') || p.includes('S.INVEST')) return 'COSTO ENERGETICO';
    if (p.includes('SCOTIABANK')) return 'CAMBIO DIVISA';
    if (p.includes('LUIS QUIROZ') || p.includes('LUIS ALBERTO')) return 'TRANSPORTE';

    if (p.includes('CGE') || p.includes('ELECTRICIDAD')) return 'ENERGIA';
    if (p.includes('AGUA') || p.includes('RIEGO')) return 'AGUA';
    if (p.includes('TRANSPORTE') || p.includes('FLETE') || p.includes('DHL') || p.includes('LATAM')) return 'TRANSPORTE';
    if (p.includes('AGROKIMUN') || p.includes('FERTILI') || p.includes('ABONO')) return 'INSUMOS AGRICOLAS';
    if (p.includes('FERRETERIA') || p.includes('AGROEQUIPOS')) return 'HERRAMIENTAS';
    if (p.includes('ARIDOS') || p.includes('GRAVA') || p.includes('ARENA')) return 'MATERIALES';
    if (p.includes('LIBRERIA') || p.includes('TUCAN') || p.includes('MERCADOLIBRE') || p.includes('OFFICE')) return 'SUMINISTROS';
    if (d.includes('ARRENDAMIENTO') || d.includes('ARRIENDO')) return 'ARRIENDOS';
    if (p.includes('PULLMAN') || p.includes('RESTAURANT') || p.includes('CAFE')) return 'ALIMENTACION';
    if (p.includes('COMERCIAL') || p.includes('COMERCIALIZADORA')) return 'SUMINISTROS';

    return 'SERVICIOS';
}

console.log('Reading Excel file:', excelPath);
const workbook = XLSX.readFile(excelPath);
const sheet = workbook.Sheets['Facturas'];

console.log('Processing data...');
let updateCount = 0;
let skiCount = 0;

for (let row = 2; row < 2000; row++) {
    const catCell = sheet[XLSX.utils.encode_col(16) + row];  // Column Q (17th) = index 16

    if (catCell && catCell.v === 'REVISAR') {
        const provCell = sheet[XLSX.utils.encode_col(3) + row];   // Column D (4th) = index 3
        const detailCell = sheet[XLSX.utils.encode_col(7) + row]; // Column H (8th) = index 7

        const proveedor = provCell ? provCell.v : '';
        const detalle = detailCell ? detailCell.v : '';

        const newCat = getCategory(proveedor, detalle);
        catCell.v = newCat;

        updateCount++;
        if (row % 100 === 0) {
            console.log(`  Row ${row}: ${proveedor} -> ${newCat}`);
        }
    }
}

console.log(`\nUpdated ${updateCount} invoices`);
console.log('Saving file...');
XLSX.writeFile(workbook, excelPath);
console.log('Done!');
