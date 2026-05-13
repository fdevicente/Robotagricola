# 03b — Arquitectura: componentes en detalle

## `cash_flow/categorizer.py`
Recibe factura/cargo banco → llama a Claude → devuelve (categoria, cultivo, confianza).

```python
categorize_invoice(row) -> CategorizationResult
categorize_bank_movement(row) -> CategorizationResult
batch_categorize_history() -> Report  # corre 1× en onboarding
```

- Prompt con proveedor + glosa I + glosa II + monto + fecha
- Cache de prompts (proveedor+glosa similares cacheados)
- Confianza <0.85 → marca "REVISAR" + manda Telegram

## `cash_flow/matcher.py`
Cuando scraper trae movimientos, busca la factura asociada y escribe Fecha Pago.

```python
match_new_bank_movements(movements) -> MatchReport
```

Algoritmo:
1. Descripción contiene N° factura → match directo
2. Si no, busca pendientes con TOTAL = monto ±0.01% en ventana ±15 días, mismo proveedor
3. Match único → escribe Fecha Pago auto
4. >1 candidato → inline keyboard Telegram

## `cash_flow/projector.py`
Cerebro del dashboard. Calcula proyección por mes × categoría × cultivo.

```python
get_cash_flow(start_date, end_date) -> CashFlowProjection
```

Devuelve:
- opening_balance (real del banco)
- monthly: dict[mes, MonthBreakdown]
- alerts: list[Alert]
- comparativos: {2024_ajustado, 2025_real, 2026_proyectado}

Memoizado 30 min. Invalidable.

## `cash_flow/income_wizard.py`
FSM por usuario en Telegram. Pregunta cosecha por cultivo:
kg → exportadoras → precio → cuotas → fechas → liquidación final.

Output: hoja `Cosechas` del Master.

## `cash_flow/replante.py`
Affordability check.

```python
afford_check(cultivo, hectareas=None) -> ReplanteResult
```

Lógica:
- Costo/hc = promedio facturas Inversión/Replante año anterior / hc plantadas
- Disponible = saldo_proyectado_fin_año - saldo_min_seguridad - pendientes
- Si no alcanza → calcula endeudamiento necesario

## `cash_flow/reporter.py`
Genera PDF mensual formato directorio. WeasyPrint HTML→PDF.

Secciones:
1. Resumen ejecutivo
2. Tabla ingresos del mes
3. Tabla egresos por categoría × cultivo
4. Top 10 proveedores
5. Comparativo vs año anterior ajustado
6. Notas / desviaciones
7. Proyección próximos 3 meses

Output: `Reportes/Cierre_YYYY-MM.pdf`

## `cash_flow/alerts.py`
Genera y dispara notificaciones Telegram según reglas. Dedupe por tipo+mes.

## `infrastructure/backups.py`
```python
backup_master(reason: str)
backup_codebase()
daily_master_snapshot()
```

## `infrastructure/manual_gen.py`
Escanea handlers + crons + alertas → genera `docs/MANUAL_TELEGRAM.md`.

## `documentos/router.py`
Claude detecta tipo: factura | boleta | guía_despacho | otro → rutea.

## `guias_despacho/process_guia.py`
Claude extrae: fecha, cultivo, kg, exportadora, n_guia, camión, sector.
Guarda en hoja `Guias Despacho` + copia PDF a Dropbox.

## `historical_importer.py`
Onboarding 1×:
1. Categoriza facturas históricas con Claude
2. Importa cosechas desde DATOS COSECHA + FXP
3. Importa guías de despacho históricas
4. Detecta ingresos pasados del banco (venta de dólares + Vitakai CLP)
