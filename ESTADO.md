# Estado del Bot — Agrícola Santa Elisa

**Al 2026-08-24.** Este es el documento **vivo**: se actualiza.
`HANDOFF.md` y `CAMBIOS_SESION.md` están **congelados** como históricos — no los edites ni los uses como contexto.

La lista de trabajo pendiente vive en la memoria del asistente:
`~/.claude/projects/C--Users-Windows-Desktop-Workflow/memory/project_pendientes_roadmap.md`

---

## Cómo se opera

| | |
|---|---|
| Arranque | Tarea de Windows `AgricolaBotWatchdog` → `watchdog_bot.ps1` (AtLogon) |
| Reiniciar | `Stop-Process` del proceso `main.py`; el watchdog lo relanza en 10 s |
| Python | `%LOCALAPPDATA%\Python\bin\python3.11.exe` (**no hay `python` en PATH**) |
| Tests | `python -m pytest tests/ -q` — **404 pasando** |
| Dashboard | `http://localhost:5000` (login con roles admin/gerencia/campo) |

⚠️ Al identificar el proceso del bot para matarlo, **verifica que el padre sea el watchdog de Agrícola**: en la misma máquina corren otros bots (Pokémon, Wix) con `main.py`.

## Jobs programados

| job | cuándo |
|---|---|
| `banco_viernes` | **viernes 08:00** — un intento; si falla, pide la cartola |
| `resumen_semanal` | **lunes 08:00** |
| `bodega_check` | **lunes 08:30** |
| `heartbeat` | diario 20:00 |
| `latido` | cada 5 min (silencioso) |
| `sync_db` | diario 21:00 |
| `vacaciones_mensuales` · `reporte_mensual` | día 1 |

⚠️ **En python-telegram-bot ≥ 20, `days` va de 0=domingo a 6=sábado** (antes 0=lunes).
Viernes = `days=(5,)`, lunes = `days=(1,)`. `resumen_semanal` y `bodega_check` estuvieron
meses corriendo en domingo por un `days=(0,)` heredado. Hay un test que falla si
alguien vuelve a escribir `days=(0,)` en código.

## Reglas que no se negocian

- **Respaldar el Master antes de modificarlo** (`infrastructure/backups.py`).
- **Español siempre**, en código, mensajes y commits.
- **El scraper del banco no fuerza CAPTCHA ni antibot.** Si el portal bloquea, la vía es
  la carga manual de cartola por Telegram, que funciona.
- **Las credenciales bancarias no salen del PC local.**
- **En los tests que escriben Excel, pasar SIEMPRE la ruta explícita.** Un test destruyó
  el Master real por confiar en el default de `_save_wb`.
- **El bot nunca inventa correlativos de FXP.**
- En FXP, columna Saldo: `Pagada` = listo · `NN` = no se paga · monto o `#VALUE!` = por pagar.

## Piezas principales

```
main.py                     arranque, registro de handlers y jobs, menú de comandos
handlers/                   telegram: facturas, finanzas, bitácora, maquinaria,
                            personal, conciliación, banco_upload, monitoreo
modules/
  cash_flow/projector.py    proyección mes × categoría × cultivo
  conciliacion_*.py         conciliador estilo Chipax (5 fases, N:M y parcial)
  banco_import.py           carga manual de cartola, dedup por nº doc y multiconjunto
  telegram_backup.py        respaldo CRUDO de todo lo que entra por Telegram
  bitacora_asistencia.py    parte de asistencia → una fila por actividad
  maquinaria.py             horómetros, fichas, mantenciones
  db/                       espejo SQLAlchemy del Excel (modo paralelo)
infrastructure/backups.py   respaldo del Master a Dropbox
src/dashboard.py            Flask + login por roles + guardián central por prefijo
```

## Lo que hay que saber antes de tocar datos

- **La columna `Fecha` de `Bitácora` guarda TEXTO `"YYYY-MM-DD"`**, no fechas de Excel.
  Cualquier comparación tiene que parsear el string.
- **La fecha de la bitácora es la del TRABAJO**, no la del mensaje: Juan reporta días después.
- **`Días Cubiertos`** (columna 18) dice cuántos días abarca una lectura de horómetro.
  Con huecos de varios días las horas se atribuyen al **mes**, no al día.
- **La columna L de `Facturas` (`TOTAL NETO`, fórmulas `=J*K`) no tiene valores en caché**:
  leerla con `data_only=True` devuelve `None`. El proyector usa la columna O.
- **El tipo de cambio vive en DOS lugares** que deben coincidir: `config.py`
  (`CASH_FLOW_CONFIG`, lo lee el proyector) y la hoja `Config` del Master (lo lee
  `cuentas._tipo_cambio()`, y **manda**). Hoy: **910**.
- **`Richard Padilla` y `Richard Padilla Crespo` son dos personas** (padre e hijo).
  "richard" a secas es el padre.

## Respaldo de Telegram

`files/telegram/YYYY-MM.jsonl` — una línea JSON por mensaje recibido, escrita **antes**
de que ningún handler interprete nada. Está en `.gitignore` porque contiene el contenido
de los mensajes. Acumula **desde el 24-ago-2026**; para lo anterior hay que exportar el
chat desde Telegram.

## No subir al repositorio

`.env` · `bot.log` · `*.xlsx` · `files/` · `.bot_state.json` · `.bot_persistence.pickle*` ·
`Claude/` (capturas del portal del banco: muestran saldos) · `node_modules/` ·
`agricola.db` · `.flask_secret` · `.dashboard_token`
