# 05a — UI Telegram: comandos nuevos

| Comando | Qué hace |
|---|---|
| `/saldo` | Saldo actual banco + comparación con mínimo seguridad |
| `/proyeccion [meses]` | Flujo de caja próximos N meses (default 6) |
| `/proyeccion_completa` | Manda PDF del flujo completo año + comparativa 2025 |
| `/categoria [nombre]` | Detalle categoría: gastado mes, anterior, año, % presupuesto |
| `/cosecha [cultivo]` | Inicia/reabre wizard de cierre de cosecha |
| `/cosecha_actual [cultivo]` | Kg despachados temporada actual vs estimado |
| `/replante [cultivo] [hc]` | Affordability check; sin args = wizard interactivo |
| `/reporte [YYYY-MM]` | Genera/manda reporte mensual PDF (default: mes anterior) |
| `/revisar` | Lista pendientes: categorización dudosa, matches ambiguos, ingresos sin clasificar |
| `/ajuste [+/-monto] [cat] [mes]` | Ajuste manual a proyección |
| `/refresh` | Fuerza recalcular proyección + recargar Master |
| `/manual` | Manda el MANUAL_TELEGRAM.md actualizado |
| `/guias [cultivo]` | Resumen guías despacho temporada actual |
| `/cuota [exportadora] [#] [ajuste]` | Actualizar cuota específica |

Todos usan menus inline cuando los argumentos son ambiguos o faltan.
