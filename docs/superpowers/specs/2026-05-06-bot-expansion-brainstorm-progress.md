# Brainstorm — Expansión Bot Agrícola Santa Elisa

**Fecha inicio:** 2026-05-06
**Estado:** En progreso — fase de preguntas clarificadoras
**Próximo paso al retomar:** Continuar preguntas para Fase 1 (flujo de caja)

---

## 1. Objetivo general

Expandir el bot de Telegram existente (que procesa facturas vía OCR + Claude AI y guarda en Excel MASTER) para incluir:

1. 📦 **Inventario automático** — facturas con fertilizantes → agregar stock; rebajar via mensajes Telegram
2. 🚜 **Maquinaria** — pestaña con tractores, mantenciones, cambios de aceite
3. 🏖️ **Vacaciones mejoradas** — mejor control del personal
4. 📊 **Reportes mensuales** — trabajos hechos, gastos, resúmenes
5. 💰 **Dashboard con flujo de dinero** — presupuesto, cosechas, comparación anual
6. 📋 **Multi-Excel output** — al guardar factura → Master + Planilla de Gastos + FXP
7. 🗂️ **Reorganización del Excel** — optimizar estructura

---

## 2. Prioridad definida (FASE 1)

**🚨 URGENCIA: Flujo de caja y proyección de gastos**

### Contexto financiero actual
- Recién terminó la temporada
- Producción: **240.000 kg de nueces**
- Precio: **2.4 USD/kg**
- Adelanto recibido: **$223 millones CLP**
- Falta pagar al menos **$50 millones CLP** en facturas
- Quedarán aprox **$120 millones CLP** para el resto del año
- Hay que replantar → necesita saber cuánto puede gastar

### Evolución histórica de hectáreas

| Año | Nogales | Cerezos | Avellanos | Cambio |
|-----|---------|---------|-----------|--------|
| 2024 | ~64 hc | 4 hc (1ra cosecha, perdió 17.000 USD) | — | — |
| 2025 | ~52 hc | 4 hc | nuevos | -12 hc nogales |
| 2026 | ~38 hc | 8 hc | + posible más | -14 hc nogales, +4 cerezas |

**Implicación clave:** No se puede repetir gastos del año pasado — hay que **escalar costos por hectárea** según superficie cambiante cada año.

---

## 3. Estado del proyecto actual

### Bot existente — funcionalidades ya implementadas
- Telegram bot recibe fotos/PDFs de facturas
- Pipeline: Escaneo (CamScanner-style) → OCR Tesseract → Claude AI → Preview → Excel
- Guardado con reintentos (Excel locked file handling)
- Alertas automáticas: vencimientos 8am, banco 18:00
- Edición de campos antes de guardar
- Detección de duplicados
- Módulos existentes: tareas, inventario, vacaciones, dashboard, caja chica
- Scraper Scotiabank parcial/experimental

### Excel MASTER actual
- Ubicación: `C:\Users\Windows\...\MASTER Agricola Santa Elisa.xlsx`
- Datos completos desde **2024**
- Datos parciales **2022-2023**
- Hojas: Facturas, Boletas, Caja Chica, Cuenta Banco, Proveedores
- 16 columnas en Facturas

### Categorización actual de gastos
- **Pocos gastos están etiquetados por cultivo**
- Cerezas: usan sobrante de nogales en general; algunos gastos específicos sí marcados (cosecha, etc.)
- Mayoría de gastos son generales del campo

---

## 4. Datos históricos disponibles en Dropbox

**Ruta:** `C:\Users\Windows\Dropbox\CAMARICO 2023\`

### Archivos clave identificados

#### `PLANILLA GASTOS CAMARICO 2023-2026.xlsx` ⭐
**Hojas:** RESUMEN, CARGO II, PROVEEDOR, FF, Hoja1, Hoja2, figueroa, PAGOS, CAJA CHICA, Hoja3, CARGO, Hoja4, Hoja5, MANO OBRA, RESS TEMP 2024, Hoja7, Hoja8, DATOS, PROVEEDORES

**Contenido:**
- RESUMEN tiene tabla dinámica con gastos 2021-2022 por mes
- CARGO II agrupa por categoría (ABOGADO, etc.)
- FF agrupa por temporada → CAMPO, ARRIENDOS
- Estructura columnas en CARGO: TEMPEMPLO, PROVI, ARE_-SUB-, cargo, Cargo B, Iteme, Proveedor, RUT, Document-, GLORIA, GLORIA-, Fecha-, MES, MES C-I, año, N_DOC, U UM, Cantidad, TOTAL NES, IVA, ESPECIF-, TOTI-, TOTAL

⚠️ **Este es el "Planilla de Gastos" externo donde el bot debe duplicar el guardado.**

#### `PRESUPUESTO 2024-2025.xlsx`
**Hojas:** Hoja2, Hoja1
**Estructura columnas:** ITEM, MES, AÑO, CARGO, RR, CC, CCC, UNIDAD, CANTIDAD, VALOR, TOTAL
**Ejemplo:** REMUNERACIONES, MAYO, 2024, NOGALES, GENERAL, GENTE PLANTA, PERSONAS, 10, 31564, 315640

⭐ **Tiene categorización por CARGO (NOGALES) — esto es clave para proyecciones.**

#### `DATOS COSECHA.xlsx`
**Hojas:** SECADO Y PACKING, SALIDA CAJONES, ENTRADA ACOPIO

**Columnas relevantes:**
- SALIDA CAJONES: FECHA, CAJONES, SACAS, KILOS, EQUIPO, SECTOR, DESTINO, HA
- ENTRADA ACOPIO: FECHA, OPERADOR, MM, N COLOSADAS, EQUIPO, SECTOR, KG UNIT, KG TOTAL

⭐ **Datos por SECTOR y EQUIPO — permite costos/rendimiento por zona.**

#### `BODEGA ENTRADAS-SALIDAS fda.xlsx`
**Hojas:** PORTADA, SAG, Historial de plaguicidas usados, Hoja1, STOK BODEGA JUAN P, **Nogales**, **Cerezos**, RESUMEN, Hoja2, STOCK, HERBICIDAS, Hoja3, Hoja5, TRABAJO, Hoja4, PRODUCTOS, C COSTOS

⭐ **Tiene hojas separadas por Nogales vs Cerezos — categorización ya existente.**
⭐ Contiene listado SAG completo de pesticidas autorizados (referencia para el bot).

#### `CAMPO SITUACION MARZO 2025.xlsx`
**Hojas:** E1 (2), RESS, E1, E1 PURO, E2, E2 PURO, E3, 2025, E4

Datos por **EQUIPO** (E1, E2, E3, E4) y **SECTOR** — útil para asignar gastos.

### Otros archivos relevantes en Dropbox
- `MAQUINARIAS1 fda.xlsx` — datos de tractores y maquinaria
- `ASISTENCIA TEMP 2023-2024 fda.xlsx` — asistencia del personal temporal
- `COSTO EMPRESA/` — carpeta con costos mensuales 2023
- `CAJA CHICA TEMP 2023-2024.xlsx` — caja chica histórica
- `EXPORTADORAS/EXPOTADA.xlsx` — datos de exportación
- `MARCO PLANTACION.xlsx` — datos de plantación

---

## 5. Decisiones de diseño tomadas hasta ahora

1. ✅ Visual companion aceptado (puede usar mockups en navegador para diseño visual)
2. ✅ Fase 1 = Flujo de caja / proyección de gastos
3. ✅ Proyección debe ajustar por hectáreas variables año a año
4. ✅ Hay 3 fuentes de datos: MASTER, Dropbox CAMARICO, FXP (formato propio del usuario)
5. ✅ Nivel de proyección: **Opción C — por categoría + mes** (estacional)
6. ✅ Categorización: **Opción 3 (híbrido)** — categorías nuevas orientadas a flujo de caja, conservando campo CARGO libre para retrocompatibilidad
7. ✅ **Visión a futuro: retirar FXP**. Master se vuelve fuente única. Scraper Scotiabank rellena Cuenta Banco. Cuando se paga una factura, bot agrega Fecha Pago a Master.Facturas con formato FXP.
8. ✅ **Ingresos vienen de FXP.ScotiaBCO** (columna Deposito) + **FXP.Ingresos** (matriz exportadora×mes). No hace falta ingreso manual.
9. ✅ **Lista de 11 categorías** + cruce con cultivo (NOGALES/CEREZOS/AVELLANOS/GENERAL): Mano de obra planta, Mano de obra temporal, Fertilizantes, Fitosanitarios, Combustible, Maquinaria-mantención, Riego, Servicios profesionales, Arriendos/Patentes/Seguros, Inversión/Replante, Caja chica/Imprevistos
10. ✅ **Categorización del histórico: Claude AI** (~1300 facturas, ~USD $5 una sola vez)
11. ✅ **Año base proyección: 2025**, con 2024 como referencia visual comparativa, y capa de ajustes manuales encima
12. ✅ **Master = fuente única de verdad** (actualizado al día por el usuario). NO hay migración masiva desde FXP.
13. ✅ **Mano de obra híbrida**: liquidaciones masivas se categorizan desde Master.Cuenta Banco (descripción del cargo); honorarios de Francisco se cargan en Master.Facturas como una factura más.
14. ✅ **Matching banco↔factura: B + Telegram para ambiguos**. Match automático cuando hay match único por monto+proveedor+nº factura; cuando hay duda, bot pregunta por Telegram con opciones.
15. ✅ **Dashboard táctico (B) + extras**: comparación gasto 2025 ajustado por hectáreas, "cuánto queda para el año", simulador de replante interactivo, reporte mensual formal para directorio
16. ✅ **Saldo mínimo de seguridad: 10% del gasto anual proyectado** (~$36M CLP con base 2025), recalculado anualmente
17. ✅ **Alertas estándar (B)** + reporte cierre de mes formato directorio
18. ✅ **Simulador replante**: simple "affordability check" (¿alcanza la plata?), no ROI calculator. Costo por hectárea calculado desde facturas+banco categorizadas como Inversión/Replante.
19. ✅ **Wizard post-cosecha**: cuando termina cosecha, bot pregunta por Telegram (kg, exportadora, precio, cuotas, fechas). Vos confirmás/ajustás. Misma lógica para cerezas. Avellanos placeholder hasta 2028.
20. ✅ **Refresh banco: 1×/día a las 18:00**
21. ✅ **Dashboard combinado Telegram + browser local** (puerto 5000), con visión futura de app móvil
22. ✅ **Arquitectura: Enfoque 1 — extensión modular del bot actual, Excel-only** (módulo `cash_flow/` paralelo a tareas, inventario, vacaciones)

## 5b. Estructura de archivos descubierta

### FXP.xlsx — `C:\Users\Windows\Dropbox\Agricola Santa Elisa\FXP.xlsx`
12 hojas: FXP, **ScotiaBCO**, ScotiaUSD, **Ingresos**, **Flujo 2025-26**, Flujo 2024-5, Info Campo, Flujo 2023-4, 2023 y 2024 cosecha, BCOChile, BCOChile USD, BCOSantander

**ScotiaBCO** (4828 filas, desde 2016, hasta 2026-05-06):
- Cols: Emisión, Vencimiento, **Pago**, Año, N, Descripción, **Monto** (cargo), **Deposito** (abono), **Saldo**, Asig Cta, Notas
- Saldo actual: **~$130.621.871 CLP**

**FXP** (hoja, 1361 filas, formato "1 línea"):
- Cols: Fecha Emision, Fecha Vencimiento, **Fecha Pago**, N, Mes, Año, Nombre Factura, Numero Factura, **Monto**, Ingresado, Deposito, **Saldo**, Notas
- Notas guarda detalle del producto (ej: "Insumo Agricola Fosf Monoam")
- Estado en columna "Notas" tipo "Pagada"

**Ingresos** (matriz):
- Filas: Vitakai, Valbifrut, Biopinon
- Sub-filas: Pesos / Usd
- Columnas: meses 2023-2024+ (Mayo, Junio, Julio… Abril)

**Flujo 2025-26** (estructura de proyección actual):
- Cabecera: 2025 (Mayo-Diciembre) + 2026 (Enero-Abril) — temporada agrícola
- Sección "Ingresos" con líneas por exportadora (Vitakai, Biopinon)
- Probablemente sigue con sección Egresos (verificar)

### Master.xlsx — `C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER Agricola Santa Elisa.xlsx`
Hojas: Facturas, Proveedores, **Cuenta Banco**, Tareas, Bitácora, Inventario, Aplicaciones, Personal, Vacaciones

**Cuenta Banco** (4710 filas, ya parcialmente migrada):
- Cols: Fecha, Descripcion, Referencia, **Cargo**, **Abono**, Saldo
- ⚠️ Esquema más simple que ScotiaBCO (no tiene Año, N, Asig Cta, Notas)

**Facturas** (1351 filas):
- 16 cols incluyendo **Fecha Pago** ya existente → el campo de linking ya está presente

---

## 6. Preguntas pendientes (retomar aquí)

### Para Fase 1 — Flujo de caja
- [ ] ¿Cómo se categorizan los gastos en Master actualmente? (ya parcialmente respondido)
- [ ] **¿Qué nivel de proyección quiere?** 3 opciones a evaluar:
  - **A. Proyección simple:** total año pasado escalado por hectáreas actuales
  - **B. Proyección por categoría:** desglosado por tipo de gasto (fertilizante, mano obra, combustible, mantención, etc.) — más preciso
  - **C. Proyección por mes:** patrón estacional del año pasado escalado al actual — capta que ciertos meses son más caros (cosecha, fumigación)
- [ ] ¿El dashboard debe mostrar saldo en tiempo real (cuánto queda) o también predicción mensual?
- [ ] ¿Quiere alertas de Telegram cuando se acerque al límite presupuestario?
- [ ] ¿Cómo se ingresan los ingresos? (adelantos de exportadora, ventas) — ¿manualmente o detectar de algún archivo?

### Para Fases siguientes (después de cash flow)
- [ ] **Inventario automático:**
  - ¿Cómo detectar que una factura es de fertilizante? (palabra clave, proveedor, manual)
  - ¿Comando Telegram para rebajar stock? (`/uso fertilizante X cantidad Y sector Z`)
- [ ] **Maquinaria:**
  - ¿Lista actual de tractores/máquinas?
  - ¿Qué se trackea? (horas uso, kilómetros, fechas mantención, tipo)
  - ¿Cómo asociar factura de mantención con máquina específica?
- [ ] **Vacaciones:**
  - ¿Qué falta del módulo actual? (problemas concretos)
- [ ] **Multi-Excel output:**
  - Ver formato exacto de FXP (todo en una línea sin items)
  - Mapeo de columnas Master ↔ Planilla Gastos ↔ FXP

---

## 7. Estructura propuesta de fases (preliminar)

```
FASE 1 — Flujo de caja / proyección (URGENTE)
├── Categorizar gastos históricos en Master
├── Importar histórico de Planilla Gastos (Dropbox)
├── Cálculo de costo por hectárea por categoría
├── Pestaña dashboard "Flujo de Caja" con proyección
└── Alertas Telegram de presupuesto

FASE 2 — Multi-Excel output
├── Guardar factura en Planilla Gastos (formato CARGO/CC/CCC)
└── Guardar factura en FXP (formato 1 línea)

FASE 3 — Inventario automático
├── Detección automática de fertilizantes/insumos
├── Auto-agregar al inventario al guardar factura
└── Comando Telegram para rebajar uso

FASE 4 — Maquinaria
├── Pestaña Maquinaria con tractores
├── Asociar mantenciones a máquina
└── Reporte de costo por máquina

FASE 5 — Reportes mensuales + dashboard avanzado
├── Reporte mensual automático Telegram
├── Pestañas dashboard: Cosechas, Comparación anual
└── Mejoras vacaciones

FASE 6 — Reorganización Excel (consolidación final)
```

---

## 8. Información que se le pidió al usuario

- ✅ Años de datos en Master (respuesta: 2022+, completo desde 2024)
- ✅ Categorización de gastos (respuesta: pocos por cultivo)
- ✅ Datos históricos en Dropbox (respuesta: sí, en `CAMARICO 2023`)
- ✅ Cosechas pasadas (respuesta: 240k kg nueces 2026)

### Pendiente de pedir al usuario
- Lista actual de tractores/maquinaria
- Formato exacto del FXP (planilla del usuario)
- Categorías de gasto que quiere ver en proyección
- Lista de personal con días de vacación pendientes
- Ejemplos de productos de inventario más comunes

---

## 9. Próximos pasos al retomar

1. **Continuar preguntas Fase 1** — definir nivel de proyección (A/B/C arriba)
2. Revisar más a fondo `PLANILLA GASTOS CAMARICO 2023-2026.xlsx` (las hojas RESUMEN y CARGO tienen las categorizaciones)
3. Revisar `PRESUPUESTO 2024-2025.xlsx` para ver categorías que el usuario ya usaba
4. Una vez claras las preguntas: presentar 2-3 enfoques arquitectónicos
5. Presentar diseño por secciones para aprobación
6. Escribir spec final en `docs/superpowers/specs/2026-05-06-flujo-caja-design.md`
7. Pasar a `writing-plans` para implementación

---

## 10. Notas técnicas

- Bot corre en `py -3.11 main.py`
- Excel MASTER tiene retry logic (5 intentos, 4 seg espera)
- ANTHROPIC_API_KEY ya configurada
- Credenciales en Windows Credential Manager
- Dashboard Flask en puerto 5000
- Se prefiere español en toda interacción
- No usar merge_cells en Excel
- Hacer backup antes de modificar

---

**Última actualización:** 2026-05-06
**Próxima sesión:** Continuar desde sección 6 — preguntas pendientes Fase 1
