# 06b — Gráfico presupuesto vs real

## Gráfico principal: barras + línea

Eje X: meses (May-26 a Abr-27)
Eje Y izquierdo: montos CLP

Elementos:
- Barra verde: egresos proyectados del mes
- Barra roja/naranja encima: exceso sobre proyección (si hubo)
- Barra azul clara: ingresos del mes
- Línea negra: saldo proyectado cierre mes
- Línea punteada gris: saldo real (meses pasados)
- Línea roja horizontal: saldo mínimo seguridad ($36M)

## Drill-down "¿por qué?"

Click en cualquier barra de un mes → panel lateral muestra:
- Desglose por categoría (11 categorías)
- Por cada categoría: proyectado vs real
- Semáforo: verde (<90%), amarillo (90-100%), rojo (>100%)
- Top 3 facturas que más aportaron al desvío
- Texto: "Fertilizantes +23% → factura Cals $1.29M no presupuestada"

## Comparación año anterior

Toggle en el gráfico: "Mostrar 2025 ajustado"
- Agrega barras semi-transparentes con el gasto 2025 escalado por hc
- Permite ver visualmente si 2026 va por encima o debajo

## Implementación

- Chart.js con plugin de anotaciones para línea saldo mínimo
- Click handler abre panel lateral (div oculto, sin JS framework)
- Datos vienen de projector.get_cash_flow() vía endpoint JSON
