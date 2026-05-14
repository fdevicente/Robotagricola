# Scheduler Setup

## Registrar cron 18:00 (Windows Task Scheduler)

1. Abrir **PowerShell como Administrador**
2. Ir a la carpeta del proyecto:
   ```
   cd "C:\Users\Windows\Desktop\Workflow\Agricola Santa Elisa\Robot"
   ```
3. Permitir scripts (una vez):
   ```
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
4. Registrar:
   ```
   .\scripts\register_daily_banco_task.ps1
   ```
5. Verificar en `taskschd.msc` que aparece `AgricolaSantaElisa-DailyBanco-18h`.

## Reintentos

Settings incluye `-RestartCount 3 -RestartInterval 60 min`. Si el scraper
falla a las 18:00, reintenta 19:00 y 20:00 automaticamente.

## Logs

El runner escribe a stdout. Para capturar a archivo, editar la accion
en Task Scheduler y redirigir: `> daily_banco.log 2>&1`.

## Desactivar

```
Unregister-ScheduledTask -TaskName "AgricolaSantaElisa-DailyBanco-18h" -Confirm:$false
```
