const XLSX = require('xlsx');
const path = require('path');

const excelPath = path.join(__dirname, '../MASTER Agricola Santa Elisa.xlsx');

function reclassifyServicio(proveedor, detalle) {
    if (!proveedor) proveedor = '';
    if (!detalle) detalle = '';

    const p = (proveedor || '').toUpperCase();
    const d = (detalle || '').toUpperCase();

    // Maxisacos, pallets, etc. → INSUMOS AGRICOLAS
    if (p.includes('CARMEN') && d.includes('MAXISACO')) return 'INSUMOS AGRICOLAS';
    if (p.includes('CONFECCION') && d.includes('MAXISACO')) return 'INSUMOS AGRICOLAS';
    if (p.includes('SANTA TERESITA') && d.includes('MAXISACO')) return 'INSUMOS AGRICOLAS';
    if (p.includes('PACIFOR') && d.includes('PALLET')) return 'INSUMOS AGRICOLAS';

    // Pinturas → INSUMOS AGRICOLAS
    if (d.includes('PINTURA') && d.includes('DECORATIVA')) return 'INSUMOS AGRICOLAS';
    if (d.includes('PINTURAS')) return 'INSUMOS AGRICOLAS';

    // Soplador, herramientas → HERRAMIENTAS
    if (d.includes('SOPLADOR') || d.includes('HUSQVARNA')) return 'HERRAMIENTAS';
    if (d.includes('HERRAMIENT')) return 'HERRAMIENTAS';

    // Equipos de red, WiFi, cámaras de seguridad → SEGURIDAD
    if (d.includes('SWITCH') || d.includes('POE')) return 'SEGURIDAD';
    if (d.includes('WIFI') || d.includes('INALAMBR')) return 'SEGURIDAD';
    if (d.includes('CAMARA') || d.includes('CÁMARA')) return 'SEGURIDAD';
    if (d.includes('EAP') || d.includes('PUNTO DE ACCESO')) return 'SEGURIDAD';
    if (p.includes('NASJAC') && d.includes('ACCESO')) return 'SEGURIDAD';
    if (p.includes('PCPLAY') && (d.includes('SWITCH') || d.includes('CAMARA'))) return 'SEGURIDAD';

    // Transporte
    if (d.includes('FLETE') || d.includes('TRANSPORTE')) return 'TRANSPORTE';
    if (d.includes('TRASLADO')) return 'TRANSPORTE';

    // Mantenimiento
    if (d.includes('CERCO') || d.includes('POLIN')) return 'MANTENIMIENTO';
    if (d.includes('ACEITE') || d.includes('NEUMATICO') || d.includes('REPUESTO')) return 'MANTENIMIENTO';

    // Materiales
    if (d.includes('MAICILLO') || d.includes('ARIDO')) return 'MATERIALES';
    if (d.includes('TETRAPACK')) return 'MATERIALES';

    // Inspecciones/servicios profesionales
    if (d.includes('INSPECCION') || d.includes('INSPECCIONES')) return 'SERVICIOS PROFESIONALES';

    // Default: keep as SERVICIOS
    return 'SERVICIOS';
}

console.log('Reading Excel file...');
const workbook = XLSX.readFile(excelPath);
const sheet = workbook.Sheets['Facturas'];

let updateCount = 0;
let reclassByCategory = {};

for (let row = 2; row < 2000; row++) {
    const catCell = sheet[XLSX.utils.encode_col(16) + row];

    if (catCell && catCell.v === 'SERVICIOS') {
        const provCell = sheet[XLSX.utils.encode_col(3) + row];
        const detailCell = sheet[XLSX.utils.encode_col(7) + row];

        const proveedor = provCell ? provCell.v : '';
        const detalle = detailCell ? detailCell.v : '';

        const newCat = reclassifyServicio(proveedor, detalle);

        if (newCat !== 'SERVICIOS') {
            catCell.v = newCat;
            updateCount++;

            if (!reclassByCategory[newCat]) {
                reclassByCategory[newCat] = 0;
            }
            reclassByCategory[newCat]++;

            if (updateCount <= 10) {
                console.log(`  Row ${row}: ${proveedor.substring(0, 30).padEnd(30)} → ${newCat}`);
            }
        }
    }
}

console.log(`\nReclassified ${updateCount} invoices:`);
Object.entries(reclassByCategory).forEach(([cat, count]) => {
    console.log(`  ${cat}: ${count} items`);
});

console.log('\nSaving file...');
XLSX.writeFile(workbook, excelPath);
console.log('Done!');
