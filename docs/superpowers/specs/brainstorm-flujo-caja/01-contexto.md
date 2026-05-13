# 01 — Contexto

## Objetivo de Fase 1

Construir un módulo de **flujo de caja con proyección por categoría × mes × cultivo** que le permita al usuario saber, en cualquier momento:

- Cuánta plata tiene
- Cuánta plata va a entrar y cuándo
- En qué se va a gastar mes a mes
- Si va a alcanzar para replantar avellanos
- Si hay alguna alerta financiera por venir

## Urgencia financiera (al 2026-05-06)

- Cosecha 2026: **240.000 kg nueces** (140k Valbifrut @ 1.8 USD + 100k Pacific Nuts @ 1.6 USD + liquidación dic)
- Cosecha 2025: **290.000 kg nueces** (referencia histórica)
- Adelanto recibido: $223M CLP
- Falta pagar: ~$50M CLP en facturas
- Saldo banco actual: $130.6M CLP
- Hay replante de avellanos en curso y planificado a varios años

## Evolución de hectáreas

| Año | Nogales | Cerezos | Avellanos | Total |
|-----|---------|---------|-----------|-------|
| 2024 | 65 hc | 1.8 hc | 0 hc | 66.8 hc |
| 2025 | 54 hc | 3.8 hc | 11.5 hc | 69.3 hc |
| 2026 | 43 hc | 3.8 hc | 26.5 hc | 73.3 hc |

## Fuentes de datos

- **Master.xlsx** (`C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\MASTER...`) — fuente única de verdad
- **FXP.xlsx** (`C:\Users\Windows\Dropbox\Agricola Santa Elisa\FXP.xlsx`) — histórico, se retira
- **CAMARICO 2023/** (Dropbox) — datos cosecha, presupuesto, bodega históricos
