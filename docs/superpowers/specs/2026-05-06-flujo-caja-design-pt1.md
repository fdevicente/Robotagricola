# Spec — Fase 1: Flujo de Caja — Parte 1/4

## 1. Objetivo

Módulo de flujo de caja con proyección por categoría × mes × cultivo para el bot agrícola.
Responde: cuánta plata hay, cuánta entra, en qué se gasta, si alcanza para replantar.

## 2. Contexto

- Cosecha 2026: 240k kg nueces (Valbifrut 140k @ 1.8 USD, Pacific Nuts 100k @ 1.6 USD)
- Cosecha 2025: 290k kg nueces (referencia histórica)
- Saldo banco: $130.6M CLP. Falta pagar ~$50M. Replante avellanos en curso.
- Hectáreas: 2026 = 38 nogales + 8 cerezos + avellanos (datos exactos pendientes Daniel)

## 3. Decisiones de diseño (31 total)

### Categorización
- D1: Proyección por categoría + mes (estacional)
- D2: Categorías híbridas (nuevas + CARGO libre retrocompatible)
- D3: 11 categorías × 4 cultivos (NOGALES/CEREZOS/AVELLANOS/GENERAL)
- D4: Claude AI categoriza histórico (~$5 USD una vez)
- D8: Mano obra: masiva desde banco, honorarios Francisco en Facturas

### Proyección
- D5: Año base 2025 + 2024 referencia + ajustes manuales
- D6: Master = fuente única de verdad
- D10: Ingresos USD en ScotiaUSD, CLP (Vitakai) en ScotiaBCO

### Infraestructura
- D11: Match banco↔factura auto + Telegram para ambiguos
- D13: Saldo mínimo = 10% gasto anual (~$36M)
- D17: Refresh banco 1×/día 18:00
- D19: Enfoque 1 — extensión modular, Excel-only
- D21: Backup automático Dropbox (Master + código)
- D22: Manual Telegram auto-generado

### Features
- D12: Dashboard táctico + comparación 2025 ajustado por hc
- D14: Alertas estándar + reporte mensual PDF directorio
- D15: Simulador replante = affordability check (no ROI)
- D16: Wizard post-cosecha por Telegram
- D25: Router documentos: factura/boleta/guía_despacho/otro
- D27: Fechas límite cosecha: cerezas 15-dic, nueces 30-may
- D28: Import automático cosechas históricas desde Dropbox
- D29: Flujo Caja contrastado con saldo real banco
- D30: Línea ingresos extraordinarios (detección automática)
- D31: Replante con escenario deuda

### Fases futuras
- D23: Bitácora Inteligente NLP = Fase 2
- D24: Secuencia: F1 Cash → F2 NLP → F3 Multi-Excel → F4 Inventario → F5 Maquinaria → F6 Reportes → F7 App
